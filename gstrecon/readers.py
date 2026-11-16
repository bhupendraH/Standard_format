"""Read the purchase register and GSTR-2B workbooks into Docs."""
from __future__ import annotations

import io
import re

import openpyxl

from .model import Doc, G2bFile, InputError, PrFile
from .normalise import is_gstin, normalise_number, parse_date, to_float

MONTHS = ("january", "february", "march", "april", "may", "june", "july", "august",
          "september", "october", "november", "december")
_FILENAME_MONTH = re.compile(
    r"(?<![A-Za-z])(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s_-]+(\d{4})(?!\d)", re.I)

PR_REQUIRED = (
    "Transaction Type", "Party GSTIN", "Doc Ref", "Transaction Date", "Doc Date",
    "ARAP", "Reverse Charge", "Taxable Value", "CGST Amount", "SGST Amount",
    "IGST Amount", "Invoice Amount", "Party Name", "Party Region",
)


def _open(data: bytes):
    try:
        return openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception as exc:  # corrupt / not an xlsx
        raise InputError(f"not a readable Excel file ({type(exc).__name__})") from None


def _text(value) -> str:
    return "" if value is None else str(value).strip()


def read_pr(data: bytes) -> PrFile:
    wb = _open(data)
    rows = wb.worksheets[0].iter_rows(values_only=True)
    header = next(rows, None) or ()
    col = {_text(h): i for i, h in enumerate(header) if _text(h)}
    for name in PR_REQUIRED:
        if name not in col:
            raise InputError(f"PR missing column: {name}")

    def get(row, name):
        i = col[name]
        return row[i] if i < len(row) else None

    groups: dict = {}
    months = set()
    for row in rows:
        if all(v is None or _text(v) == "" for v in row):
            continue
        if _text(get(row, "Transaction Type")).lower() == "journal":
            continue
        booking = parse_date(get(row, "Transaction Date"))
        if booking is None:
            raise InputError("PR row without Transaction Date")
        months.add((booking.year, booking.month))
        key = (_text(get(row, "Party GSTIN")).upper(), _text(get(row, "ARAP")), _text(get(row, "Doc Ref")))
        g = groups.get(key)
        if g is None:
            groups[key] = g = {
                "row": row, "booking": booking,
                "taxable": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0,
            }
        g["taxable"] += to_float(get(row, "Taxable Value"))
        g["igst"] += to_float(get(row, "IGST Amount"))
        g["cgst"] += to_float(get(row, "CGST Amount"))
        g["sgst"] += to_float(get(row, "SGST Amount"))

    if not groups:
        raise InputError("PR has no rows")
    if len(months) > 1:
        listed = ", ".join(f"{m:02d}/{y}" for y, m in sorted(months))
        raise InputError(f"PR spans several months: {listed}")

    docs = []
    for (gstin, arap, doc_ref), g in groups.items():
        row = g["row"]
        total = g["taxable"] + g["igst"] + g["cgst"] + g["sgst"]
        credit = total < 0
        sign = -1 if credit else 1  # store credit notes as positive amounts, like GSTR-2B
        docs.append(Doc(
            gstin=gstin,
            vendor_name=_text(get(row, "Party Name")),
            number=arap,
            norm=normalise_number(arap),
            doc_type="Credit Note" if credit else "Invoice",
            doc_date=parse_date(get(row, "Doc Date")),
            value=abs(to_float(get(row, "Invoice Amount"))),
            taxable=round(sign * g["taxable"], 2),
            igst=round(sign * g["igst"], 2),
            cgst=round(sign * g["cgst"], 2),
            sgst=round(sign * g["sgst"], 2),
            pos_name=_text(get(row, "Party Region")),
            reverse_charge=_text(get(row, "Reverse Charge")) or "N",
            rcm="/RCM/" in arap.upper(),
            extra={"erp_ref": doc_ref, "booking_date": g["booking"]},
        ))
    return PrFile(month=months.pop(), docs=docs)


# --------------------------------------------------------------------------- GSTR-2B

_B2B_FIELDS = {
    "gstin": "GSTIN of supplier", "name": "Trade/Legal name", "number": "Invoice number",
    "subtype": "Invoice type", "date": "Invoice Date", "value": "Invoice Value",
}
_CDNR_FIELDS = {
    "gstin": "GSTIN of supplier", "name": "Trade/Legal name", "number": "Note number",
    "note_type": "Note type", "subtype": "Note Supply type", "date": "Note date",
    "value": "Note Value",
}
_COMMON_FIELDS = {
    "pos": "Place of supply", "rc": "Supply Attract Reverse Charge",
    "taxable": "Taxable Value", "igst": "Integrated Tax", "cgst": "Central Tax",
    "sgst": "State/UT Tax", "cess": "Cess", "period": "GSTR-1/IFF/GSTR-5 Period",
    "filed": "GSTR-1/IFF/GSTR-5 Filing", "itc": "ITC Availability", "reason": "Reason",
    "pct": "Applicable %", "source": "Source", "irn": "IRN", "irn_date": "IRN Date",
}
_READ_SHEETS = {"Read me", "B2B", "B2B-CDNR"}


def _label(text) -> str:
    return re.sub(r"\s+", " ", _text(text)).lower()


def _columns(rows5_6, fields: dict, sheet: str) -> dict:
    row5, row6 = rows5_6
    width = max(len(row5), len(row6))
    labels = []
    for i in range(width):
        sub = row6[i] if i < len(row6) else None
        top = row5[i] if i < len(row5) else None
        labels.append(_label(sub if _text(sub) else top))
    out = {}
    for key, want in fields.items():
        want = _label(want)
        idx = next((i for i, l in enumerate(labels) if l == want), None)
        if idx is None:
            idx = next((i for i, l in enumerate(labels) if l.startswith(want)), None)
        if idx is None:
            raise InputError(f"GSTR-2B sheet {sheet} missing column: {fields[key]}")
        out[key] = idx
    return out


def _pct(value) -> float:
    text = _text(value).rstrip("%")
    if not text:
        return 0.0
    n = to_float(text)
    return n / 100 if n > 1 else n


def _g2b_docs(ws, fields: dict, sheet: str) -> list:
    rows = ws.iter_rows(min_row=5, values_only=True)
    head = (next(rows, ()) or (), next(rows, ()) or ())
    col = _columns(head, {**fields, **_COMMON_FIELDS}, sheet)

    def get(row, key):
        i = col[key]
        return row[i] if i < len(row) else None

    docs = []
    for row in rows:
        if not row or not is_gstin(row[0]):
            break
        number = _text(get(row, "number"))
        doc_type = _text(get(row, "note_type")).title() if "note_type" in col else "Invoice"
        docs.append(Doc(
            gstin=_text(get(row, "gstin")).upper(),
            vendor_name=_text(get(row, "name")),
            number=number,
            norm=normalise_number(number),
            doc_type=doc_type or "Credit Note",
            doc_date=parse_date(get(row, "date")),
            value=to_float(get(row, "value")),
            taxable=to_float(get(row, "taxable")),
            igst=to_float(get(row, "igst")),
            cgst=to_float(get(row, "cgst")),
            sgst=to_float(get(row, "sgst")),
            cess=to_float(get(row, "cess")),
            pos_name=_text(get(row, "pos")),
            reverse_charge=_text(get(row, "rc")) or "No",
            extra={
                "invoice_type": _text(get(row, "subtype")),
                "filing_period": _text(get(row, "period")),
                "filing_date": parse_date(get(row, "filed")),
                "itc_available": _text(get(row, "itc")),
                "reason": _text(get(row, "reason")),
                "applicable_pct": _pct(get(row, "pct")),
                "source": _text(get(row, "source")),
                "irn": _text(get(row, "irn")),
                "irn_date": parse_date(get(row, "irn_date")),
            },
        ))
    return docs


def _read_me(ws) -> dict:
    out = {}
    # The value is the first non-empty cell after the label (merged cells push it to C).
    for row in ws.iter_rows(max_row=30, max_col=6, values_only=True):
        if row and _text(row[0]):
            out[_label(row[0])] = next((v for v in row[1:] if _text(v)), None)
    return out


def read_gstr2b(data: bytes) -> G2bFile:
    wb = _open(data)
    if "Read me" not in wb.sheetnames:
        raise InputError("GSTR-2B file has no 'Read me' sheet")
    info = _read_me(wb["Read me"])
    period = _label(info.get("tax period"))
    fy = _text(info.get("financial year"))
    m = re.fullmatch(r"(\d{4})-(\d{2})", fy)
    if period not in MONTHS or not m:
        raise InputError("GSTR-2B 'Read me' has no readable Tax Period / Financial Year")
    month = MONTHS.index(period) + 1
    year = int(m.group(1)) + (1 if month <= 3 else 0)

    docs = []
    if "B2B" in wb.sheetnames:
        docs += _g2b_docs(wb["B2B"], _B2B_FIELDS, "B2B")
    if "B2B-CDNR" in wb.sheetnames:
        docs += _g2b_docs(wb["B2B-CDNR"], _CDNR_FIELDS, "B2B-CDNR")

    ignored = {}
    for name in wb.sheetnames:
        if name in _READ_SHEETS or name.startswith("ITC "):
            continue
        n = sum(
            1 for row in wb[name].iter_rows(min_row=6, values_only=True)
            if any(isinstance(v, (int, float)) and not isinstance(v, bool) for v in row)
        )
        if n:
            ignored[name] = n

    return G2bFile(
        month=(year, month),
        gstin=_text(info.get("gstin")).upper(),
        financial_year=fy,
        generated=parse_date(info.get("date of generation")),
        docs=docs,
        ignored=ignored,
    )


def month_from_filename(name: str):
    base = re.split(r"[\\/]", name)[-1]
    m = _FILENAME_MONTH.search(base)
    if not m:
        return None
    month = [x[:3] for x in MONTHS].index(m.group(1).lower()) + 1
    return int(m.group(2)), month
