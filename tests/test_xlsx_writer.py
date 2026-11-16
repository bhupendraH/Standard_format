import io
import os
import re
import zipfile
from datetime import date

import openpyxl
import pytest

from gstrecon.xlsx_writer import TemplateError, fill_template
from tests.helpers import BOOKS_MONTH, BOT_CATEGORY, make_mini_template

RUN = date(2026, 11, 16)


def bot_rows(n):
    return [{"id": i, "status_ctgy": "EXACT MATCH", "pr_vendor_name": f"V{i}",
             "pr_doc_date": date(2025, 3, i)} for i in range(1, n + 1)]


def read(data):
    return openpyxl.load_workbook(io.BytesIO(data))


def table_ref(data, sheet, name):
    return read(data)[sheet].tables[name].ref


def test_fill_writes_rows_and_resizes_table():
    out = fill_template(make_mini_template(), {"BOT_ReconData": bot_rows(3)}, {})
    ws = read(out)["BOT_ReconData"]
    assert [ws.cell(r, 1).value for r in range(2, 5)] == [1, 2, 3]
    assert ws.cell(3, 3).value == "V2"
    assert ws.cell(4, 4).value.date() == date(2025, 3, 3)
    assert ws.cell(4, 4).number_format == "dd-mm-yyyy"  # style taken from the prototype row
    assert ws.max_row == 4  # the template's old row 3 is gone
    assert table_ref(out, "BOT_ReconData", "reco_datatable") == "A1:F4"


def test_fill_copies_formulas_to_every_row():
    out = fill_template(make_mini_template(), {"BOT_ReconData": bot_rows(3)}, {})
    ws = read(out)["BOT_ReconData"]
    for r in range(2, 5):
        assert ws.cell(r, 5).value == "=" + BOT_CATEGORY
        assert ws.cell(r, 6).value.ref == f"F{r}"  # array formula re-anchored on its own row


def test_fill_zero_rows_keeps_prototype():
    out = fill_template(make_mini_template(), {"Books_CurrMnth": [], "BOT_ReconData": []}, {})
    ws = read(out)["Books_CurrMnth"]
    assert table_ref(out, "Books_CurrMnth", "BooksCurrent") == "A1:C2"
    assert ws.cell(2, 3).value == "=" + BOOKS_MONTH
    assert ws.cell(2, 1).value is None
    assert read(out)["BOT_ReconData"].max_row == 2


def test_fill_untouched_parts_byte_identical():
    tpl = make_mini_template()
    out = fill_template(tpl, {"BOT_ReconData": bot_rows(4)}, {"C2": "X"})
    a, b = zipfile.ZipFile(io.BytesIO(tpl)), zipfile.ZipFile(io.BytesIO(out))
    assert [n for n in a.namelist() if n != "xl/calcChain.xml"] == b.namelist()
    workbook_level = {"[Content_Types].xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"}
    differing = {n for n in b.namelist() if a.read(n) != b.read(n)} - workbook_level
    # BOT sheet + its table + the report sheet (header cell) - nothing else
    assert len(differing) == 3
    for n in differing:
        assert re.match(r"xl/(worksheets/sheet\d+|tables/table\d+)\.xml$", n), n


def test_fill_sets_full_calc_and_drops_calcchain():
    out = fill_template(make_mini_template(), {"BOT_ReconData": bot_rows(1)}, {})
    z = zipfile.ZipFile(io.BytesIO(out))
    assert 'fullCalcOnLoad="1"' in z.read("xl/workbook.xml").decode()
    assert b"calcChain" not in z.read("xl/_rels/workbook.xml.rels")
    assert b"calcChain" not in z.read("[Content_Types].xml")


def test_fill_header_cells():
    out = fill_template(make_mini_template(), {}, {"C2": "27AADCA0878J1ZI", "C3": RUN})
    ws = read(out)["Recon Summary"]
    assert ws["C2"].value == "27AADCA0878J1ZI"
    assert ws["C3"].value.date() == RUN
    assert ws["C3"].number_format == "dd-mm-yyyy"
    assert ws["B2"].value == "Office GSTIN"


def test_fill_escapes_special_characters():
    name = 'DN SHIPPING & LOGISTICS <"x"> \x0b'
    out = fill_template(make_mini_template(), {"BOT_ReconData": [{"id": 1, "pr_vendor_name": name}]}, {})
    assert read(out)["BOT_ReconData"].cell(2, 3).value == 'DN SHIPPING & LOGISTICS <"x"> '


def test_fill_unknown_sheet_raises():
    with pytest.raises(TemplateError, match="Nope"):
        fill_template(make_mini_template(), {"Nope": []}, {})


def test_fill_missing_formulas_raises():
    with pytest.raises(TemplateError, match="template missing formulas: reco_datatable"):
        fill_template(make_mini_template(formulas=False), {"BOT_ReconData": bot_rows(1)}, {})


@pytest.mark.skipif(not os.environ.get("RECON_TEMPLATE"), reason="set RECON_TEMPLATE to the fixed template")
def test_fill_real_template_if_available():
    tpl = open(os.environ["RECON_TEMPLATE"], "rb").read()
    rows = [{"id": 1, "status_ctgy": "PR+"}, {"id": 2, "status_ctgy": "GSTR2B+"}]
    out = fill_template(tpl, {"BOT_ReconData": rows}, {"C2": "27AADCA0878J1ZI", "C3": RUN})
    z = zipfile.ZipFile(io.BytesIO(out))
    assert "xl/metadata.xml" in z.namelist()
    sheets = "".join(z.read(n).decode("utf-8") for n in z.namelist()
                     if n.startswith("xl/worksheets/sheet"))
    assert sheets.count("_xlfn.PIVOTBY") == 2
    for n in z.namelist():
        if n.startswith("xl/pivotCache/pivotCacheDefinition"):
            assert 'refreshOnLoad="1"' in z.read(n).decode("utf-8")
    assert read(out)["BOT_ReconData"].cell(3, 1).value == 2
