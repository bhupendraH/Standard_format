"""Builders for synthetic input workbooks used by the tests (no real data)."""
from __future__ import annotations

import io

import openpyxl

PR_COLUMNS = [
    "Office", "Office_Region", "Office_GSTIN", "Transaction Type", "Party Type",
    "Party GSTIN", "Doc Ref", "Transaction Date", "Doc Date", "FA Vou Ref", "ARAP",
    "Invoice Type", "Invoice Source", "Reverse Charge", "Tariff/Account", "SAC/HSN",
    "Invoice Value", "Taxable Value", "CGST Rate", "CGST Amount", "SGST Rate",
    "SGST Amount", "IGST Rate", "IGST Amount", "Invoice Amount", "Service Name",
    "Job Ref", "Party Name", "Party Region", "GST Group",
]

PR_COLUMNS_APRIL = [
    "Office", "Office_Region", "Office_GSTIN", "Transaction Type", "Party Type",
    "Party Name", "Party Region", "Party GSTIN", "Doc Ref", "Transaction Date",
    "Doc Date", "FA Vou Ref", "ARAP", "Invoice Type", "Invoice Source",
    "Reverse Charge", "Tariff/Account", "SAC/HSN", "Invoice Value", "Taxable Value",
    "GST Group", "CGST Rate", "CGST Amount", "SGST Rate", "SGST Amount",
    "IGST Rate", "IGST Amount", "Invoice Amount",
]


def pr_line(**kw) -> dict:
    """One PR line with sensible defaults; override any column by keyword."""
    row = {
        "Office": "ARC Mumbai",
        "Office_GSTIN": "27AADCA0878J1ZI",
        "Transaction Type": "Invoice",
        "Party Type": "Vendor",
        "Party GSTIN": "27AEEPN0699C1ZH",
        "Doc Ref": "AMULP00889/24-25",
        "Transaction Date": "28-Mar-2025",
        "Doc Date": "28-Mar-2025",
        "ARAP": "DN/003/25-26",
        "Reverse Charge": "N",
        "Taxable Value": 17600,
        "CGST Amount": 1584,
        "SGST Amount": 1584,
        "IGST Amount": 0,
        "Invoice Amount": 20768,
        "Party Name": "DN SHIPPING & LOGISTICS",
        "Party Region": "MAHARASHTRA",
    }
    row.update(kw)
    return row


def _xlsx(wb) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def make_pr_xlsx(rows, columns=None) -> bytes:
    columns = columns or PR_COLUMNS
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(columns)
    for r in rows:
        ws.append([r.get(c) for c in columns])
    return _xlsx(wb)


B2B_ROW5 = ["GSTIN of supplier", "Trade/Legal name", "Invoice Details", None, None, None,
            "Place of supply", "Supply Attract Reverse Charge", "Taxable Value (₹)", "Tax Amount",
            None, None, None, "GSTR-1/IFF/GSTR-5 Period", "GSTR-1/IFF/GSTR-5 Filing Date",
            "ITC Availability", "Reason", "Applicable % of Tax Rate", "Source", "IRN", "IRN Date"]
B2B_ROW6 = [None, None, "Invoice number", "Invoice type", "Invoice Date", "Invoice Value(₹)",
            None, None, None, "Integrated Tax(₹)", "Central Tax(₹)", "State/UT Tax(₹)", "Cess(₹)"]
CDNR_ROW5 = ["GSTIN of supplier", "Trade/Legal name", "Credit note/Debit note details", None, None,
             None, None, "Place of supply", "Supply Attract Reverse Charge", "Taxable Value (₹)",
             "Tax Amount", None, None, None, "GSTR-1/IFF/GSTR-5 Period",
             "GSTR-1/IFF/GSTR-5 Filing Date", "ITC Availability", "Reason",
             "Applicable % of Tax Rate", "Source", "IRN", "IRN Date"]
CDNR_ROW6 = [None, None, "Note number", "Note type", "Note Supply type", "Note date",
             "Note Value (₹)", None, None, None, "Integrated Tax(₹)", "Central Tax(₹)",
             "State/UT Tax(₹)", "Cess(₹)"]


def b2b_row(gstin="27AEEPN0699C1ZH", name="DN Shipping & Logistics", number="DN/1702/24-25",
            date="28/03/2025", value=20768, pos="Maharashtra", rc="No", taxable=17600,
            igst=0, cgst=1584, sgst=1584, cess=0, period="Mar'25", filed="10/04/2025",
            itc="Yes", reason="", pct="100%", source="E-Invoice", irn="abc", irn_date="08/04/2025"):
    return [gstin, name, number, "Regular", date, value, pos, rc, taxable, igst, cgst, sgst, cess,
            period, filed, itc, reason, pct, source, irn, irn_date]


def cdnr_row(gstin="27AACCG1899E1ZH", name="GATEWAY TERMINALS INDIA P LTD", number="3470",
             note_type="Credit Note", date="26/03/2025", value=899.51, pos="Maharashtra", rc="No",
             taxable=762.3, igst=0, cgst=68.61, sgst=68.61, cess=0, period="Mar'25",
             filed="10/04/2025", itc="Yes", reason="", pct="100%", source="", irn="", irn_date=""):
    return [gstin, name, number, note_type, "Regular", date, value, pos, rc, taxable, igst, cgst,
            sgst, cess, period, filed, itc, reason, pct, source, irn, irn_date]


def _portal_sheet(wb, title, row5, row6, rows):
    ws = wb.create_sheet(title)
    ws.append(["Goods and Services Tax  - GSTR-2B"])
    ws.append([])
    ws.append([])
    ws.append(["Taxable inward supplies received from registered persons"])
    ws.append(row5)
    ws.append(row6)
    for r in rows:
        ws.append(r)


def make_g2b_xlsx(b2b=(), cdnr=(), *, period="March", fy="2024-25", gstin="27AADCA0878J1ZI",
                  generated="14/04/2025", extra_sheets=None) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Read me"
    ws.append(["Goods and Services Tax  - GSTR-2B"])
    ws.append([])
    ws.append([])
    ws.append(["Financial Year", None, fy])  # value sits in column C, as on the portal
    ws.append(["Tax Period", None, period])  # value sits in column C, as on the portal
    ws.append(["GSTIN", None, gstin])  # value sits in column C, as on the portal
    ws.append(["Legal Name", "TEST CO"])
    ws.append(["Trade Name (if any)", "TEST CO"])
    ws.append(["Date of generation", None, generated])  # value sits in column C, as on the portal
    summary = wb.create_sheet("ITC Available")
    summary.append(["FORM GSTR-2B"])
    summary.append([1, 2, 3])  # numeric cells in a summary sheet must not count as ignored
    _portal_sheet(wb, "B2B", B2B_ROW5, B2B_ROW6, b2b)
    _portal_sheet(wb, "B2B-CDNR", CDNR_ROW5, CDNR_ROW6, cdnr)
    for title, rows in (extra_sheets or {}).items():
        _portal_sheet(wb, title, ["Header"], ["Sub"], rows)
    return _xlsx(wb)


BOT_CATEGORY = 'IF(reco_datatable[[#This Row],[status_ctgy]]="EXACT MATCH","Matched","Other")'
BOT_REMARKS = 'IF(reco_datatable[[#This Row],[status_ctgy]]="PR+","x","y")'
BOOKS_MONTH = "VALUE(MONTH(BooksCurrent[[#This Row],[doc_date]]))"

_CALC_CT = (b'<Override PartName="/xl/calcChain.xml" ContentType="application/'
            b'vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>')
_CALC_REL = (b'<Relationship Id="rIdCalc" Type="http://schemas.openxmlformats.org/officeDocument/'
             b'2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>')
_CALC_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><calcChain xmlns="http://'
             'schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="E2" i="4"/></calcChain>')


def _table(ws, name, ref, headers, formulas):
    from openpyxl.worksheet.table import Table, TableColumn, TableFormula

    t = Table(displayName=name, name=name, ref=ref)
    t.tableColumns = [
        TableColumn(id=i, name=h, calculatedColumnFormula=(
            TableFormula(attr_text=formulas[h]) if h in formulas else None))
        for i, h in enumerate(headers, 1)
    ]
    ws.add_table(t)


def make_mini_template(formulas=True) -> bytes:
    """A small workbook shaped like Output_template.xlsx: a report sheet with header
    cells, data sheets with tables whose row 2 is the formula prototype, an untouched
    Sheet1, and a calcChain part."""
    import zipfile
    from datetime import date

    from openpyxl.worksheet.formula import ArrayFormula

    wb = openpyxl.Workbook()
    rs = wb.active
    rs.title = "Recon Summary"
    rs["B2"], rs["C2"] = "Office GSTIN", "27AABCB5853G1ZP"
    rs["B3"], rs["C3"] = "Reco Date", date(2025, 3, 20)
    rs["C3"].number_format = "dd-mm-yyyy"
    wb.create_sheet("Sheet1")["A1"] = "keep me"
    wb.create_sheet("EXCLUSIONS").append(["Doc Number", "Vendor Name", "Error Flag"])

    bot = wb.create_sheet("BOT_ReconData")
    heads = ["id", "status_ctgy", "pr_vendor_name", "pr_doc_date", "Category", "Remarks"]
    bot.append(heads)
    bot["D2"].number_format = "dd-mm-yyyy"
    if formulas:
        bot["E2"] = "=" + BOT_CATEGORY
        bot["F2"] = ArrayFormula("F2", "=" + BOT_REMARKS)
    else:
        bot["A2"] = None
    bot.append([99, "OLD", "old row", None, None, None])
    _table(bot, "reco_datatable", "A1:F3", heads,
           {"Category": BOT_CATEGORY, "Remarks": BOT_REMARKS} if formulas else {})

    books = wb.create_sheet("Books_CurrMnth")
    bh = ["id", "doc_date", "Booking_Mnth"]
    books.append(bh)
    books.append([None, None, ("=" + BOOKS_MONTH) if formulas else None])
    _table(books, "BooksCurrent", "A1:C2", bh, {"Booking_Mnth": BOOKS_MONTH} if formulas else {})

    raw = _xlsx(wb)
    src, out = zipfile.ZipFile(io.BytesIO(raw)), io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "[Content_Types].xml":
                data = data.replace(b"</Types>", _CALC_CT)
            if info.filename == "xl/_rels/workbook.xml.rels":
                data = data.replace(b"</Relationships>", _CALC_REL)
            z.writestr(info, data)
        z.writestr("xl/calcChain.xml", _CALC_XML)
    return out.getvalue()
