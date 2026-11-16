from datetime import date

from gstrecon.model import Doc, G2bFile, PrFile
from gstrecon.normalise import normalise_number
from gstrecon.sheets import MonthInputs, build_month

RUN = date(2026, 11, 16)
SHEETS = {"Books_CurrMnth", "Books_PrevMnth_NextMnth", "GSTR2B", "GSTR2B_PrevMonth",
          "EXCLUSIONS", "BOT_ReconData"}


def pr_doc(number="DN/003/25-26", gstin="27AEEPN0699C1ZH", tax=3168.0, rcm=False,
           booking=date(2025, 3, 28), d=date(2025, 3, 28)):
    return Doc(gstin=gstin, vendor_name="DN SHIPPING & LOGISTICS", number=number,
               norm=normalise_number(number), doc_type="Invoice", doc_date=d, value=20768.0,
               taxable=17600.0, igst=0.0, cgst=tax / 2, sgst=tax / 2, pos_name="MAHARASHTRA",
               reverse_charge="N", rcm=rcm,
               extra={"erp_ref": "AMULP00889/24-25", "booking_date": booking})


def g2b_doc(number="DN/003/25-26", gstin="27AEEPN0699C1ZH", tax=3168.0):
    return Doc(gstin=gstin, vendor_name="DN Shipping & Logistics", number=number,
               norm=normalise_number(number), doc_type="Invoice", doc_date=date(2025, 3, 28),
               value=20768.0, taxable=17600.0, igst=0.0, cgst=tax / 2, sgst=tax / 2,
               pos_name="Maharashtra", reverse_charge="No",
               extra={"invoice_type": "Regular", "filing_period": "Mar'25",
                      "filing_date": date(2025, 4, 10), "itc_available": "Yes", "reason": "",
                      "applicable_pct": 1.0, "source": "E-Invoice", "irn": "abc",
                      "irn_date": date(2025, 4, 8)})


def inputs(pr_docs=None, g2b_docs=None, pr_prev=None, pr_next=None, g2b_prev=None):
    pr = PrFile(month=(2025, 3), docs=pr_docs if pr_docs is not None else [pr_doc()])
    g2b = G2bFile(month=(2025, 3), gstin="27AADCA0878J1ZI", financial_year="2024-25",
                  generated=date(2025, 4, 14),
                  docs=g2b_docs if g2b_docs is not None else [g2b_doc()], ignored={})
    return MonthInputs(pr=pr, g2b=g2b, pr_prev=pr_prev, pr_next=pr_next, g2b_prev=g2b_prev)


def test_build_month_sheet_keys_and_header():
    wd = build_month(inputs(), RUN)
    assert set(wd.sheets) == SHEETS
    assert wd.header == {"C2": "27AADCA0878J1ZI", "C3": RUN}


def test_books_row_fields():
    row = build_month(inputs(), RUN).sheets["Books_CurrMnth"][0]
    assert row["id"] == 1
    assert row["doc_booking_tax_period"] == date(2025, 3, 1)
    assert row["doc_booking_date"] == date(2025, 3, 28)
    assert row["doc_type"] == "Invoice"
    assert row["doc_number"] == "DN/003/25-26"
    assert row["erp_ref_number"] == "AMULP00889/24-25"
    assert row["place_of_supply"] == "MAHARASHTRA"
    assert row["tax_value"] == 3168.0
    assert row["int_doc_number"] == normalise_number("DN/003/25-26")
    assert row["doc_date_formatted"] == date(1900, 1, 1)
    assert row["taxable_value_fromatted"] == 0
    assert "Booking_Mnth" not in row


def test_gstr2b_row_fields():
    row = build_month(inputs(), RUN).sheets["GSTR2B"][0]
    assert row["org_gstin"] == "27AADCA0878J1ZI"
    assert row["tax_period"] == 32025
    assert row["financial_year"] == "2024-25"
    assert row["generation_date"] == date(2025, 4, 14)
    assert row["reverse_charge_flag"] == "No"
    assert row["place_of_supply"] == "Maharashtra"
    assert row["applicable_tx"] == 1.0
    assert "GST_TGST" not in row


def test_rcm_in_books_and_exclusions_not_bot():
    rcm = pr_doc(number="ARC2425/RCM/M047, MT/BILL/2174/JAN/2025", gstin="UNREGISTERED", rcm=True)
    wd = build_month(inputs(pr_docs=[rcm, pr_doc()]), RUN)
    assert len(wd.sheets["Books_CurrMnth"]) == 2
    assert wd.sheets["Books_CurrMnth"][0]["error_flag"] == "FAIL"
    assert wd.sheets["Books_CurrMnth"][0]["error_comment"] == "RCM Invoice"
    ex = wd.sheets["EXCLUSIONS"]
    assert len(ex) == 1
    assert ex[0]["Error Flag"] == "FAIL" and ex[0]["Error Comment"] == "RCM Invoice"
    assert ex[0]["Doc Booking Tax Period"] == date(2025, 3, 1)
    assert len(wd.sheets["BOT_ReconData"]) == 1
    assert wd.counts["RCM excluded"] == 1


def test_bot_exact_row():
    wd = build_month(inputs(), RUN)
    row = wd.sheets["BOT_ReconData"][0]
    assert row["id"] == 1 and row["recon_id"] == 1
    assert row["status_ctgy"] == "EXACT MATCH"
    assert row["pr_recon_id"] == 1 and row["gstr2b_id"] == 1
    assert row["pr_vendor_gstin"] == row["gstr2b_vendor_gstin"] == "27AEEPN0699C1ZH"
    assert row["pr_place_of_supply"] == 27 and row["gstr2b_place_of_supply"] == 27
    assert row["gstr2b_tax_period"] == 32025
    assert row["gstr2b_document_type"] == "INVOICE"
    assert row["gstr2b_reverse_charge_flag"] == "N"
    assert row["gstr2b_original_document_date"] == date(1900, 1, 1)
    assert row["gstr2b_amended_flag"] == "N"
    assert row["diff_tax_value"] == 0
    assert row["tax_value_status"] == "Match"
    assert row["document_number_status"] == "Match"
    assert wd.counts["EXACT MATCH"] == 1


def test_bot_diff_is_2b_minus_books():
    wd = build_month(inputs(g2b_docs=[g2b_doc(tax=3200.0)]), RUN)
    row = wd.sheets["BOT_ReconData"][0]
    assert row["status_ctgy"] == "MISMATCH"
    assert row["diff_tax_value"] == 32.0
    assert row["diff_cgst_value"] == 16.0
    assert row["tax_value_status"] == "Mismatch"


def test_bot_pr_plus_row_has_blank_2b_side():
    wd = build_month(inputs(g2b_docs=[]), RUN)
    row = wd.sheets["BOT_ReconData"][0]
    assert row["status_ctgy"] == "PR+"
    assert all(v is None for k, v in row.items() if k.startswith("gstr2b_"))
    assert all(row.get(k) is None for k in row if k.startswith("diff_") or k.endswith("_status"))
    assert row["pr_vendor_gstin"] == "27AEEPN0699C1ZH"


def test_bot_g2b_plus_row_has_blank_pr_side():
    wd = build_month(inputs(pr_docs=[]), RUN)
    row = wd.sheets["BOT_ReconData"][0]
    assert row["status_ctgy"] == "GSTR2B+"
    assert all(v is None for k, v in row.items() if k.startswith("pr_"))


def test_prev_next_books_concatenated():
    prev = PrFile(month=(2025, 2), docs=[pr_doc("A1", booking=date(2025, 2, 3)),
                                         pr_doc("A2", booking=date(2025, 2, 4))])
    nxt = PrFile(month=(2025, 4), docs=[pr_doc("B1", booking=date(2025, 4, 2))])
    rows = build_month(inputs(pr_prev=prev, pr_next=nxt), RUN).sheets["Books_PrevMnth_NextMnth"]
    assert [r["id"] for r in rows] == [1, 2, 3]
    assert [r["doc_booking_tax_period"] for r in rows] == [date(2025, 2, 1)] * 2 + [date(2025, 4, 1)]


def test_prev_month_gstr2b_rows():
    prev = G2bFile(month=(2025, 2), gstin="27AADCA0878J1ZI", financial_year="2024-25",
                   generated=date(2025, 3, 14), docs=[g2b_doc("P1")], ignored={})
    rows = build_month(inputs(g2b_prev=prev), RUN).sheets["GSTR2B_PrevMonth"]
    assert rows[0]["tax_period"] == 22025 and rows[0]["invoice_number"] == "P1"


def test_missing_neighbours_give_empty_sheets():
    wd = build_month(inputs(), RUN)
    assert wd.sheets["Books_PrevMnth_NextMnth"] == []
    assert wd.sheets["GSTR2B_PrevMonth"] == []
