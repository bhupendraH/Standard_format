import io

import openpyxl

from gstrecon.xlsx_writer import fill_template
from tests.helpers import BOOKS_MONTH, BOT_CATEGORY, make_mini_template
from tools.make_template import DATA_SHEETS, make_template


def test_make_template_clears_data_keeps_formulas():
    filled = fill_template(make_mini_template(), {"BOT_ReconData": [{"id": i} for i in range(1, 4)]}, {})
    out = make_template(filled, sheets=("BOT_ReconData", "Books_CurrMnth", "EXCLUSIONS"))
    wb = openpyxl.load_workbook(io.BytesIO(out))
    bot = wb["BOT_ReconData"]
    assert bot.max_row == 2
    assert bot.cell(2, 1).value is None
    assert bot.cell(2, 5).value == "=" + BOT_CATEGORY
    assert bot.tables["reco_datatable"].ref == "A1:F2"
    assert wb["Books_CurrMnth"].cell(2, 3).value == "=" + BOOKS_MONTH


def test_data_sheets_are_the_six_template_tabs():
    assert set(DATA_SHEETS) == {"Books_CurrMnth", "Books_PrevMnth_NextMnth", "GSTR2B",
                                "GSTR2B_PrevMonth", "EXCLUSIONS", "BOT_ReconData"}
