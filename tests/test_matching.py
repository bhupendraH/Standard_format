from datetime import date

from gstrecon.matching import match
from gstrecon.model import Doc
from gstrecon.normalise import normalise_number


def doc(number, gstin="27AEEPN0699C1ZH", tax=3168.0, doc_type="Invoice", d=date(2025, 3, 28)):
    half = round(tax / 2, 2)
    return Doc(gstin=gstin, vendor_name="V", number=number, norm=normalise_number(number),
               doc_type=doc_type, doc_date=d, value=0.0, taxable=0.0,
               igst=0.0, cgst=half, sgst=round(tax - half, 2))


def statuses(pairs):
    return [(p.pr.number if p.pr else None, p.g2b.number if p.g2b else None, p.status) for p in pairs]


def test_exact_match_ignores_punctuation_case_zeros():
    pairs, _ = match([doc("INV-001")], [doc("inv1")])
    assert statuses(pairs) == [("INV-001", "inv1", "EXACT MATCH")]


def test_tolerance_boundary():
    pairs, _ = match([doc("A1", tax=100.0)], [doc("A1", tax=105.0)])
    assert pairs[0].status == "EXACT MATCH"
    pairs, _ = match([doc("A1", tax=100.0)], [doc("A1", tax=105.01)])
    assert pairs[0].status == "MISMATCH"


def test_tolerance_is_configurable():
    pairs, _ = match([doc("A1", tax=100.0)], [doc("A1", tax=103.0)], tolerance=1.0)
    assert pairs[0].status == "MISMATCH"


def test_credit_note_never_pairs_with_invoice():
    pr = [doc("JMJACNH2425MAR11", gstin="27AAAPX3996J2ZX", doc_type="Credit Note")]
    g2b = [doc("JMJACNH2425MAR11", gstin="27AAAPX3996J2ZX")]
    pairs, _ = match(pr, g2b)
    assert [p.status for p in pairs] == ["PR+", "GSTR2B+"]


def test_probable_match_same_date_and_tax():
    pairs, _ = match([doc("DN/003/25-26")], [doc("DN/1702/24-25")])
    assert statuses(pairs) == [("DN/003/25-26", "DN/1702/24-25", "PROBABLE MATCH")]


def test_probable_requires_unique_candidate():
    pairs, _ = match([doc("X1")], [doc("Y1"), doc("Y2")])
    assert sorted(p.status for p in pairs) == ["GSTR2B+", "GSTR2B+", "PR+"]


def test_probable_requires_unique_on_books_side_too():
    pairs, _ = match([doc("X1"), doc("X2")], [doc("Y1")])
    assert sorted(p.status for p in pairs) == ["GSTR2B+", "PR+", "PR+"]


def test_jmj_credit_note_stays_unpaired():
    pr = [doc("JMJACNH2425MAR11", gstin="27AAAPX3996J2ZX", doc_type="Credit Note",
              tax=1293.84, d=date(2025, 3, 31))]
    g2b = [doc("JMJACNH2425FEB11", gstin="27AAAPX3996J2ZX", doc_type="Credit Note",
               tax=1260.0, d=date(2025, 3, 6))]
    pairs, _ = match(pr, g2b)
    assert [p.status for p in pairs] == ["PR+", "GSTR2B+"]


def test_duplicate_key_not_paired_and_warned():
    pairs, warnings = match([doc("INV-1"), doc("INV/1", d=date(2025, 3, 1))], [doc("INV1")])
    assert sorted(p.status for p in pairs) == ["GSTR2B+", "PR+", "PR+"]
    assert len(warnings) == 1 and "INV" in warnings[0]


def test_unregistered_never_pairs():
    pairs, _ = match([doc("A1", gstin="UNREGISTERED")], [doc("A1", gstin="UNREGISTERED")])
    assert [p.status for p in pairs] == ["PR+", "GSTR2B+"]


def test_different_gstin_never_pairs():
    pairs, _ = match([doc("A1")], [doc("A1", gstin="27AAACF2188H1ZM")])
    assert [p.status for p in pairs] == ["PR+", "GSTR2B+"]


def test_output_order():
    pr = [doc("P-ONLY", tax=1.0, d=date(2025, 3, 1)), doc("M1")]
    g2b = [doc("G-ONLY", tax=999.0), doc("M1")]
    pairs, _ = match(pr, g2b)
    assert statuses(pairs) == [
        ("M1", "M1", "EXACT MATCH"),
        ("P-ONLY", None, "PR+"),
        (None, "G-ONLY", "GSTR2B+"),
    ]
