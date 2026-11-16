"""Turn one month's inputs and pairs into row dicts for the template's data sheets.

Rows are keyed by the sheet's header names. Formula columns (Booking_Mnth,
GST_TGST, BOT Category..Reco Decision) are left out: the writer fills them
from the template.
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from datetime import date
from typing import Optional

from .matching import match
from .model import G2bFile, PrFile
from .normalise import EMPTY_DATE, state_code, tax_period

DIFF_FIELDS = (
    ("diff_doc_value", "value"), ("diff_tax_value", "tax"), ("diff_taxable_value", "taxable"),
    ("diff_cgst_value", "cgst"), ("diff_sgst_value", "sgst"), ("diff_igst_value", "igst"),
    ("diff_cess_value", "cess"),
)


@dataclass
class MonthInputs:
    pr: PrFile
    g2b: G2bFile
    pr_prev: Optional[PrFile] = None
    pr_next: Optional[PrFile] = None
    g2b_prev: Optional[G2bFile] = None


@dataclass
class WorkbookData:
    sheets: dict
    header: dict
    counts: dict
    warnings: list = field(default_factory=list)


def _r2(x: float) -> float:
    return round(x + 0.0, 2)


def _books_rows(pr: Optional[PrFile], start_id: int = 1):
    """Returns (rows, {id(doc): row})."""
    rows, by_doc = [], {}
    if pr is None:
        return rows, by_doc
    period = date(pr.month[0], pr.month[1], 1)
    for n, d in enumerate(pr.docs, start_id):
        row = {
            "id": n,
            "doc_booking_tax_period": period,
            "doc_booking_date": d.extra.get("booking_date"),
            "doc_type": d.doc_type,
            "purch_type": None,
            "doc_number": d.number,
            "doc_date": d.doc_date,
            "erp_ref_number": d.extra.get("erp_ref"),
            "vendor_gstin": d.gstin,
            "vendor_name": d.vendor_name,
            "place_of_supply": d.pos_name.upper(),
            "reverse_charge_flag": d.reverse_charge,
            "doc_value": d.value,
            "taxable_value": d.taxable,
            "tax_value": _r2(d.tax),
            "integrated_tax": d.igst,
            "central_tax": d.cgst,
            "state_tax": d.sgst,
            "cess": 0,
            "doc_date_formatted": EMPTY_DATE,
            "boe_date_formatted": EMPTY_DATE,
            "doc_value_formatted": 0,
            "taxable_value_fromatted": 0,
            "tax_value_formatted": 0,
            "int_doc_number": d.norm,
            "error_flag": "FAIL" if d.rcm else None,
            "error_comment": "RCM Invoice" if d.rcm else None,
        }
        rows.append(row)
        by_doc[id(d)] = row
    return rows, by_doc


def _g2b_rows(g2b: Optional[G2bFile]):
    rows, by_doc = [], {}
    if g2b is None:
        return rows, by_doc
    period = tax_period(*g2b.month)
    for n, d in enumerate(g2b.docs, 1):
        row = {
            "id": n,
            "org_gstin": g2b.gstin,
            "tax_period": period,
            "financial_year": g2b.financial_year,
            "generation_date": g2b.generated,
            "vendor_gstin": d.gstin,
            "vendor_name": d.vendor_name,
            "invoice_number": d.number,
            # notes have no type column on this tab, so the note type goes here
            "invoice_type": d.extra.get("invoice_type") if d.doc_type == "Invoice" else d.doc_type,
            "invoice_date": d.doc_date,
            "invoice_value": d.value,
            "place_of_supply": d.pos_name,
            "reverse_charge_flag": d.reverse_charge,
            "rate": None,
            "taxable_value": d.taxable,
            "integrated_tax": d.igst,
            "central_tax": d.cgst,
            "state_tax": d.sgst,
            "cess": d.cess,
            "gstr_1_5_filing_period": d.extra.get("filing_period"),
            "gstr_1_5_filing_date": d.extra.get("filing_date"),
            "itc_eligibility_flag": d.extra.get("itc_available"),
            "reason": d.extra.get("reason") or None,
            "applicable_tx": d.extra.get("applicable_pct"),
            "source": d.extra.get("source") or None,
            "irn": d.extra.get("irn") or None,
            "irn_date": d.extra.get("irn_date"),
        }
        rows.append(row)
        by_doc[id(d)] = row
    return rows, by_doc


def _exclusion_row(b: dict) -> dict:
    return {
        "Doc Booking Tax Period": b["doc_booking_tax_period"],
        "Doc Booking Date": b["doc_booking_date"],
        "Doc Type": b["doc_type"],
        "Purch Type": None,
        "Doc Number": b["doc_number"],
        "Doc Date": b["doc_date"],
        "Erp Ref Number": b["erp_ref_number"],
        "Vendor Gstin": b["vendor_gstin"],
        "Vendor Name": b["vendor_name"],
        "Place Of Supply": b["place_of_supply"],
        "Reverse Charge Flag": b["reverse_charge_flag"],
        "Doc Value": b["doc_value"],
        "Taxable Value": b["taxable_value"],
        "Tax Value": b["tax_value"],
        "Integrated Tax": b["integrated_tax"],
        "Central Tax": b["central_tax"],
        "State Tax": b["state_tax"],
        "Cess": None,
        "Error Flag": "FAIL",
        "Error Comment": "RCM Invoice",
    }


def _pr_side(doc, b: dict) -> dict:
    return {
        "pr_recon_id": b["id"],
        "pr_doc_booking_tax_period": b["doc_booking_tax_period"],
        "pr_doc_booking_date": b["doc_booking_date"],
        "pr_doc_type": b["doc_type"],
        "pr_purch_type": None,
        "pr_doc_number": b["doc_number"],
        "pr_int_doc_number": b["int_doc_number"],
        "pr_doc_date": b["doc_date"],
        "pr_erp_ref_number": b["erp_ref_number"],
        "pr_vendor_gstin": b["vendor_gstin"],
        "pr_vendor_name": b["vendor_name"],
        "pr_place_of_supply": state_code(doc.pos_name, doc.gstin),
        "pr_reverse_charge_flag": b["reverse_charge_flag"],
        "pr_doc_value": b["doc_value"],
        "pr_taxable_value": b["taxable_value"],
        "pr_integrated_tax": b["integrated_tax"],
        "pr_central_tax": b["central_tax"],
        "pr_state_tax": b["state_tax"],
        "pr_cess": 0,
        "pr_icegate_reference_date": EMPTY_DATE,
    }


def _g2b_side(doc, g: dict) -> dict:
    return {
        "gstr2b_id": g["id"],
        "gstr2b_original_document_type": None,
        "gstr2b_original_document_number": None,
        "gstr2b_original_document_date": EMPTY_DATE,
        "gstr2b_org_gstin": g["org_gstin"],
        "gstr2b_tax_period": g["tax_period"],
        "gstr2b_financial_year": g["financial_year"],
        "gstr2b_generation_date": g["generation_date"],
        "gstr2b_vendor_gstin": g["vendor_gstin"],
        "gstr2b_vendor_name": g["vendor_name"],
        "gstr2b_document_type": doc.doc_type.upper(),
        "gstr2b_document_number": g["invoice_number"],
        "gstr2b_intelligent_doc_number": doc.norm,
        "gstr2b_document_supply_type": doc.extra.get("invoice_type"),
        "gstr2b_document_date": g["invoice_date"],
        "gstr2b_document_value": g["invoice_value"],
        "gstr2b_place_of_supply": state_code(doc.pos_name, doc.gstin),
        "gstr2b_reverse_charge_flag": "Y" if doc.reverse_charge.strip().lower() in ("yes", "y") else "N",
        "gstr2b_rate": None,
        "gstr2b_taxable_value": g["taxable_value"],
        "gstr2b_integrated_tax": g["integrated_tax"],
        "gstr2b_central_tax": g["central_tax"],
        "gstr2b_state_tax": g["state_tax"],
        "gstr2b_cess": g["cess"],
        "gstr2b_gstr_1_5_filing_period": g["gstr_1_5_filing_period"],
        "gstr2b_gstr_1_5_filing_date": g["gstr_1_5_filing_date"],
        "gstr2b_itc_eligibility_flag": g["itc_eligibility_flag"],
        "gstr2b_reason": g["reason"],
        "gstr2b_applicable_tx": g["applicable_tx"],
        "gstr2b_source": g["source"],
        "gstr2b_irn": g["irn"],
        "gstr2b_irn_date": g["irn_date"],
        "gstr2b_icegate_reference_date": EMPTY_DATE,
        "gstr2b_port_code": None,
        "gstr2b_boe_number": None,
        "gstr2b_boe_date": EMPTY_DATE,
        "gstr2b_amended_flag": "N",
    }


def _ok(flag: bool) -> str:
    return "Match" if flag else "Mismatch"


def _compare(a, b, tolerance: float) -> dict:
    out = {name: _r2(getattr(b, attr) - getattr(a, attr)) for name, attr in DIFF_FIELDS}
    out.update({
        "document_vendorGSTIN_status": _ok(a.gstin.strip().upper() == b.gstin.strip().upper()),
        "document_number_status": _ok(a.number.strip() == b.number.strip()),
        "document_date_status": _ok(a.doc_date == b.doc_date),
        "place_of_supply_status": _ok(state_code(a.pos_name, a.gstin) == state_code(b.pos_name, b.gstin)),
        "document_value_status": _ok(abs(b.value - a.value) <= tolerance + 1e-9),
        "taxable_value_status": _ok(abs(b.taxable - a.taxable) <= tolerance + 1e-9),
        "tax_value_status": _ok(abs(b.tax - a.tax) <= tolerance + 1e-9),
    })
    return out


def build_month(inputs: MonthInputs, run_date: date, tolerance: float = 5.0) -> WorkbookData:
    books, books_by_doc = _books_rows(inputs.pr)
    prev_rows, _ = _books_rows(inputs.pr_prev)
    next_rows, _ = _books_rows(inputs.pr_next, start_id=len(prev_rows) + 1)
    g2b, g2b_by_doc = _g2b_rows(inputs.g2b)
    g2b_prev, _ = _g2b_rows(inputs.g2b_prev)

    rcm_docs = [d for d in inputs.pr.docs if d.rcm]
    pairs, warnings = match([d for d in inputs.pr.docs if not d.rcm], inputs.g2b.docs, tolerance)

    bot = []
    for n, p in enumerate(pairs, 1):
        row = {"id": n, "recon_id": n, "status_ctgy": p.status, "status_sub_ctgy": None}
        if p.pr is not None:
            row.update(_pr_side(p.pr, books_by_doc[id(p.pr)]))
        if p.g2b is not None:
            row.update(_g2b_side(p.g2b, g2b_by_doc[id(p.g2b)]))
        if p.pr is not None and p.g2b is not None:
            row.update(_compare(p.pr, p.g2b, tolerance))
        bot.append(row)

    counts = Counter(p.status for p in pairs)
    counts["RCM excluded"] = len(rcm_docs)
    return WorkbookData(
        sheets={
            "Books_CurrMnth": books,
            "Books_PrevMnth_NextMnth": prev_rows + next_rows,
            "GSTR2B": g2b,
            "GSTR2B_PrevMonth": g2b_prev,
            "EXCLUSIONS": [_exclusion_row(books_by_doc[id(d)]) for d in rcm_docs],
            "BOT_ReconData": bot,
        },
        header={"C2": inputs.g2b.gstin, "C3": run_date},
        counts=dict(counts),
        warnings=warnings,
    )
