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
