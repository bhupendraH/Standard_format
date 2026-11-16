from datetime import date, datetime

import pytest

from gstrecon.model import InputError
from gstrecon.readers import read_pr
from tests.helpers import PR_COLUMNS, PR_COLUMNS_APRIL, make_pr_xlsx, pr_line

JMJ = dict(
    **{"Party GSTIN": "27AAAPX3996J2ZX", "Party Name": "JMJ CONTAINER SOLUTIONS",
       "ARAP": "JMJACNH2425FEB07", "Doc Ref": "AMULP00773/24-25",
       "Doc Date": "06-Mar-2025", "Transaction Date": "06-Mar-2025",
       "Invoice Amount": 121068, "IGST Amount": 0}
)


def test_read_pr_groups_lines_into_invoice():
    data = make_pr_xlsx([
        pr_line(**JMJ, **{"Taxable Value": 84600, "CGST Amount": 7614, "SGST Amount": 7614}),
        pr_line(**JMJ, **{"Taxable Value": 18000, "CGST Amount": 1620, "SGST Amount": 1620}),
    ])
    pr = read_pr(data)
    assert len(pr.docs) == 1
    d = pr.docs[0]
    assert d.taxable == 102600
    assert d.cgst == 9234 and d.sgst == 9234
    assert d.value == 121068
    assert d.doc_type == "Invoice"
    assert d.number == "JMJACNH2425FEB07"
    assert d.extra["erp_ref"] == "AMULP00773/24-25"
    assert d.extra["booking_date"] == date(2025, 3, 6)


def test_read_pr_drops_journals():
    data = make_pr_xlsx([
        pr_line(**{"Transaction Type": "Journal", "ARAP": None, "Party GSTIN": None}),
        pr_line(),
    ])
    assert len(read_pr(data).docs) == 1


def test_read_pr_credit_note_by_negative_amount():
    data = make_pr_xlsx([pr_line(**{
        "Party GSTIN": "27AACCG1899E1ZH", "ARAP": "GTICD1168657", "Doc Ref": "AMULV00014/24-25",
        "Taxable Value": -762.3, "CGST Amount": -68.61, "SGST Amount": -68.61,
        "Invoice Amount": 899.52,
    })])
    d = read_pr(data).docs[0]
    assert d.doc_type == "Credit Note"
    assert d.taxable == pytest.approx(762.3)
    assert d.cgst == pytest.approx(68.61)
    assert d.value == pytest.approx(899.52)


def test_read_pr_flags_rcm():
    data = make_pr_xlsx([pr_line(ARAP="ARC2425/RCM/M047, MT/BILL/2174/JAN/2025"), pr_line()])
    docs = read_pr(data).docs
    assert [d.rcm for d in docs] == [True, False]


def test_read_pr_month():
    assert read_pr(make_pr_xlsx([pr_line(**{"Transaction Date": "03-Mar-2025"})])).month == (2025, 3)
    two = make_pr_xlsx([pr_line(), pr_line(ARAP="X1", **{"Transaction Date": "01-Apr-2025"})])
    with pytest.raises(InputError):
        read_pr(two)


def test_read_pr_column_order_independent():
    rows = [pr_line(), pr_line(ARAP="X/9", **{"Doc Ref": "R2"})]
    assert read_pr(make_pr_xlsx(rows, PR_COLUMNS)).docs == read_pr(make_pr_xlsx(rows, PR_COLUMNS_APRIL)).docs


def test_read_pr_accepts_text_and_datetime_dates():
    a = read_pr(make_pr_xlsx([pr_line(**{"Doc Date": "28-Feb-2025"})])).docs[0]
    b = read_pr(make_pr_xlsx([pr_line(**{"Doc Date": datetime(2025, 2, 28)})])).docs[0]
    assert a.doc_date == b.doc_date == date(2025, 2, 28)


def test_read_pr_missing_column():
    cols = [c for c in PR_COLUMNS if c != "ARAP"]
    with pytest.raises(InputError, match="ARAP"):
        read_pr(make_pr_xlsx([pr_line()], cols))


def test_read_pr_empty():
    with pytest.raises(InputError):
        read_pr(make_pr_xlsx([]))
