"""Read the purchase register and GSTR-2B workbooks into Docs."""
from __future__ import annotations

import io

import openpyxl

from .model import Doc, InputError, PrFile
from .normalise import normalise_number, parse_date, to_float

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
