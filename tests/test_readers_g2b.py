from datetime import date

import pytest

from gstrecon.model import InputError
from gstrecon.normalise import normalise_number
from gstrecon.readers import month_from_filename, read_gstr2b
from tests.helpers import b2b_row, cdnr_row, make_g2b_xlsx


def test_read_g2b_header_fields():
    f = read_gstr2b(make_g2b_xlsx([b2b_row()]))
    assert f.month == (2025, 3)
    assert f.gstin == "27AADCA0878J1ZI"
    assert f.financial_year == "2024-25"
    assert f.generated == date(2025, 4, 14)


def test_read_g2b_january_is_second_fy_year():
    assert read_gstr2b(make_g2b_xlsx([b2b_row()], period="January")).month == (2025, 1)
    assert read_gstr2b(make_g2b_xlsx([b2b_row()], period="April")).month == (2024, 4)


def test_read_g2b_b2b_row():
    d = read_gstr2b(make_g2b_xlsx([b2b_row()])).docs[0]
    assert d.gstin == "27AEEPN0699C1ZH"
    assert d.vendor_name == "DN Shipping & Logistics"
    assert d.number == "DN/1702/24-25"
    assert d.norm == normalise_number("DN/1702/24-25")
    assert d.doc_type == "Invoice"
    assert d.doc_date == date(2025, 3, 28)
    assert d.value == 20768 and d.taxable == 17600
    assert d.tax == 3168
    assert d.pos_name == "Maharashtra"
    assert d.reverse_charge == "No"
    assert d.extra["invoice_type"] == "Regular"
    assert d.extra["filing_period"] == "Mar'25"
    assert d.extra["filing_date"] == date(2025, 4, 10)
    assert d.extra["itc_available"] == "Yes"
    assert d.extra["applicable_pct"] == 1.0
    assert d.extra["source"] == "E-Invoice"
    assert d.extra["irn"] == "abc"
    assert d.extra["irn_date"] == date(2025, 4, 8)


def test_read_g2b_credit_note():
    f = read_gstr2b(make_g2b_xlsx([b2b_row()], [cdnr_row()]))
    cn = f.docs[1]
    assert cn.doc_type == "Credit Note"
    assert cn.number == "3470"
    assert cn.taxable == pytest.approx(762.3)
    assert cn.extra["irn_date"] is None


def test_read_g2b_debit_note_type():
    f = read_gstr2b(make_g2b_xlsx([], [cdnr_row(note_type="Debit Note")]))
    assert f.docs[0].doc_type == "Debit Note"


def test_read_g2b_stops_at_non_gstin_rows():
    f = read_gstr2b(make_g2b_xlsx([b2b_row(), ["Total", None, None]]))
    assert len(f.docs) == 1


def test_read_g2b_counts_ignored_sheets():
    f = read_gstr2b(make_g2b_xlsx([b2b_row()], extra_sheets={"B2BA": [["27X", "n", 5]], "ECO": []}))
    assert f.ignored == {"B2BA": 1}


def test_read_g2b_requires_read_me():
    import io
    import openpyxl
    wb = openpyxl.Workbook()
    buf = io.BytesIO()
    wb.save(buf)
    with pytest.raises(InputError):
        read_gstr2b(buf.getvalue())


def test_month_from_filename():
    assert month_from_filename("ARC MAHARASHTRA GSTR2B March 2025.xlsx") == (2025, 3)
    assert month_from_filename("x Feb 2025.xlsx") == (2025, 2)
    assert month_from_filename("GSTR2B/input/PR/ARC MAHARASHTRA April 2025.xlsx") == (2025, 4)
    assert month_from_filename("PR_Mar_2025.xlsx") == (2025, 3)
    assert month_from_filename("no month.xlsx") is None
    assert month_from_filename("Remarch 20251.xlsx") is None
