import io
import json
from datetime import datetime
from types import SimpleNamespace

import openpyxl
import pytest
from azure.core.exceptions import HttpResponseError

import reconcile
from gstrecon.normalise import IST
from reconcile import main, run
from tests.helpers import (b2b_row, make_full_template, make_g2b_xlsx, make_pr_xlsx,
                           make_mini_template, pr_line)

SAS_URL = "https://acct.blob.core.windows.net/standardize-test?sv=x&sig=SECRET"
ENV = {"AZURE_BLOB_SAS_URL": SAS_URL}
NOW = datetime(2026, 11, 16, 10, 30, tzinfo=IST)
TPL = "GSTR2B/config/Output_template.xlsx"


class FakeContainer:
    def __init__(self, blobs):
        self.blobs = dict(blobs)
        self.uploads = {}

    def list_blobs(self, name_starts_with=None, **kw):
        prefix = name_starts_with or ""
        return iter(SimpleNamespace(name=n) for n in sorted(self.blobs) if n.startswith(prefix))

    def download_blob(self, name, **kw):
        data = self.blobs[name]
        return SimpleNamespace(readall=lambda: data)

    def upload_blob(self, name, data, overwrite=False, content_settings=None, **kw):
        assert overwrite is True
        self.uploads[name] = (data if isinstance(data, bytes) else data.encode("utf-8"),
                              content_settings.content_type)


def pr_file(month="Mar", year=2025, n=1, **kw):
    lines = [pr_line(ARAP=f"INV/{i}", **{"Doc Ref": f"R{i}", "Transaction Date": f"10-{month}-{year}",
                                         "Doc Date": f"05-{month}-{year}"}, **kw) for i in range(1, n + 1)]
    return make_pr_xlsx(lines)


def g2b_file(period="March", date_text="05/03/2025", n=1):
    return make_g2b_xlsx([b2b_row(number=f"INV/{i}", date=date_text) for i in range(1, n + 1)], period=period)


def blobs_for(**extra):
    blobs = {TPL: make_full_template(), "GSTR2B/config/": b""}
    blobs.update(extra)
    return blobs


def fire(container, argv=("GSTR2B",)):
    return run(list(argv), ENV, make_container=lambda url: container, now=NOW)


def summary(c):
    return c.uploads["GSTR2B/output/run_summary.txt"][0].decode("utf-8")


@pytest.mark.parametrize("argv", [["../x"], [""], ["a/b/c/d"], [], ["a", "b"], ["a b"]])
def test_run_invalid_transaction(argv):
    assert run(argv, ENV, make_container=lambda url: pytest.fail("touched Blob")) == (1, "INVALID_TRANSACTION")


def test_run_missing_secret():
    assert run(["GSTR2B"], {}) == (3, "ERROR AZURE_BLOB_SAS_URL not set")


def test_run_missing_template():
    c = FakeContainer({"GSTR2B/input/PR/a.xlsx": pr_file()})
    assert fire(c) == (2, "NOT_FOUND GSTR2B")
    assert c.uploads == {}


def test_run_produces_month_and_summary():
    c = FakeContainer(blobs_for(**{
        "GSTR2B/input/Gstr2b/ARC GSTR2B March 2025.xlsx": g2b_file(),
        "GSTR2B/input/PR/ARC March 2025.xlsx": pr_file(),
    }))
    assert fire(c) == (0, "DONE 1 month(s): Mar 2025")
    out, ctype = c.uploads["GSTR2B/output/ARC Recon March 2025.xlsx"]
    assert ctype == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ws = openpyxl.load_workbook(io.BytesIO(out))["BOT_ReconData"]
    col = [c.value for c in ws[1]].index("status_ctgy") + 1
    assert [ws.cell(r, col).value for r in range(2, ws.max_row + 1)] == ["EXACT MATCH"]
    text = summary(c)
    assert "Mar 2025" in text and "EXACT MATCH 1" in text and "ARC Recon March 2025.xlsx" in text


def test_run_skips_pr_only_month():
    c = FakeContainer(blobs_for(**{
        "GSTR2B/input/Gstr2b/ARC GSTR2B March 2025.xlsx": g2b_file(),
        "GSTR2B/input/PR/ARC March 2025.xlsx": pr_file(),
        "GSTR2B/input/PR/ARC April 2025.xlsx": pr_file("Apr"),
    }))
    assert fire(c) == (0, "DONE 1 month(s): Mar 2025")
    assert "Apr 2025: purchase register only, no GSTR-2B" in summary(c)


def test_run_month_without_previous_month():
    c = FakeContainer(blobs_for(**{
        "GSTR2B/input/Gstr2b/X GSTR2B February 2025.xlsx": g2b_file("February", "05/02/2025"),
        "GSTR2B/input/PR/X February 2025.xlsx": pr_file("Feb"),
    }))
    assert fire(c) == (0, "DONE 1 month(s): Feb 2025")
    wb = openpyxl.load_workbook(io.BytesIO(c.uploads["GSTR2B/output/X Recon February 2025.xlsx"][0]))
    assert wb["GSTR2B_PrevMonth"].tables["GST2bPrev"].ref == "A1:AB2"
    assert wb["GSTR2B_PrevMonth"].cell(2, 1).value is None


def test_run_uses_neighbour_months():
    c = FakeContainer(blobs_for(**{
        "GSTR2B/input/Gstr2b/X GSTR2B February 2025.xlsx": g2b_file("February", "05/02/2025", n=2),
        "GSTR2B/input/Gstr2b/X GSTR2B March 2025.xlsx": g2b_file(),
        "GSTR2B/input/PR/X February 2025.xlsx": pr_file("Feb", n=3),
        "GSTR2B/input/PR/X March 2025.xlsx": pr_file(),
        "GSTR2B/input/PR/X April 2025.xlsx": pr_file("Apr", n=4),
    }))
    assert fire(c) == (0, "DONE 2 month(s): Feb 2025, Mar 2025")
    wb = openpyxl.load_workbook(io.BytesIO(c.uploads["GSTR2B/output/X Recon March 2025.xlsx"][0]))
    assert wb["GSTR2B_PrevMonth"].max_row == 3  # 2 Feb GSTR-2B docs
    assert wb["Books_PrevMnth_NextMnth"].max_row == 8  # 3 Feb + 4 Apr books docs
    assert wb["Recon Summary"]["C2"].value == "27AADCA0878J1ZI"


def test_run_duplicate_month_skips_only_that_month():
    c = FakeContainer(blobs_for(**{
        "GSTR2B/input/Gstr2b/X GSTR2B February 2025.xlsx": g2b_file("February", "05/02/2025"),
        "GSTR2B/input/Gstr2b/X GSTR2B March 2025.xlsx": g2b_file(),
        "GSTR2B/input/PR/X February 2025.xlsx": pr_file("Feb"),
        "GSTR2B/input/PR/X March 2025.xlsx": pr_file(),
        "GSTR2B/input/PR/X March 2025 copy.xlsx": pr_file(),
    }))
    assert fire(c) == (0, "DONE 1 month(s): Feb 2025")
    assert "Mar 2025: two purchase register files" in summary(c)


def test_run_filename_month_disagrees():
    c = FakeContainer(blobs_for(**{
        "GSTR2B/input/Gstr2b/X GSTR2B March 2025.xlsx": g2b_file("February", "05/02/2025"),
        "GSTR2B/input/PR/X February 2025.xlsx": pr_file("Feb"),
    }))
    assert fire(c) == (4, "NOTHING_TO_DO")
    text = summary(c)
    assert "X GSTR2B March 2025.xlsx" in text and "file name says Mar 2025" in text


def test_run_unreadable_file_is_recorded():
    c = FakeContainer(blobs_for(**{
        "GSTR2B/input/Gstr2b/X GSTR2B March 2025.xlsx": g2b_file(),
        "GSTR2B/input/PR/X March 2025.xlsx": b"not excel",
    }))
    assert fire(c) == (4, "NOTHING_TO_DO")
    assert "not a readable Excel file" in summary(c)


def test_run_too_many_files():
    files = {f"GSTR2B/input/PR/p{i}.xlsx": b"x" for i in range(13)}
    assert fire(FakeContainer(blobs_for(**files))) == (3, "ERROR too many files in input/PR (13 > 12)")


def test_run_ignores_nested_and_non_xlsx():
    c = FakeContainer(blobs_for(**{
        "GSTR2B/input/Gstr2b/X GSTR2B March 2025.xlsx": g2b_file(),
        "GSTR2B/input/PR/X March 2025.xlsx": pr_file(),
        "GSTR2B/input/PR/old/X March 2025.xlsx": pr_file(),
        "GSTR2B/input/PR/notes.txt": b"hi",
    }))
    assert fire(c) == (0, "DONE 1 month(s): Mar 2025")


def test_run_tolerance_from_config():
    c = FakeContainer(blobs_for(**{
        "GSTR2B/config/settings.json": json.dumps({"tax_tolerance": 0}).encode(),
        "GSTR2B/input/Gstr2b/X GSTR2B March 2025.xlsx": make_g2b_xlsx([b2b_row(number="INV/1", date="05/03/2025", cgst=1585)]),
        "GSTR2B/input/PR/X March 2025.xlsx": pr_file(),
    }))
    assert fire(c)[0] == 0
    assert "MISMATCH 1" in summary(c)


@pytest.mark.parametrize("cfg", [b"{", b'{"tax_tolerance": -1}', b'{"tax_tolerance": "5"}', b'{"other": 1}'])
def test_run_invalid_config(cfg):
    c = FakeContainer(blobs_for(**{"GSTR2B/config/settings.json": cfg}))
    assert fire(c) == (3, "ERROR invalid config settings.json")


def test_run_template_without_formulas():
    c = FakeContainer(blobs_for(**{
        TPL: make_full_template(formulas=False),
        "GSTR2B/input/Gstr2b/X GSTR2B March 2025.xlsx": g2b_file(),
        "GSTR2B/input/PR/X March 2025.xlsx": pr_file(),
    }))
    code, line = fire(c)
    assert code == 3 and line.startswith("ERROR template missing formulas")
    assert [n for n in c.uploads if n.endswith(".xlsx")] == []


def test_run_never_prints_sas_url():
    def boom(url):
        raise HttpResponseError(message=f"failed for {url}")
    code, line = run(["GSTR2B"], ENV, make_container=boom, now=NOW)
    assert code == 3 and line.startswith("ERROR HttpResponseError")
    assert "sig=" not in line and "SECRET" not in line


def test_main_prints_one_line(capsys, monkeypatch):
    monkeypatch.setattr(reconcile, "run", lambda argv, environ: (1, "INVALID_TRANSACTION"))
    monkeypatch.setattr("sys.argv", ["reconcile.py", "x"])
    with pytest.raises(SystemExit) as e:
        main()
    assert e.value.code == 1
    assert capsys.readouterr().out == "INVALID_TRANSACTION\n"


def test_mini_template_is_not_enough():
    # a template missing data sheets is a template error, not a crash
    c = FakeContainer(blobs_for(**{
        TPL: make_mini_template(),
        "GSTR2B/input/Gstr2b/X GSTR2B March 2025.xlsx": g2b_file(),
        "GSTR2B/input/PR/X March 2025.xlsx": pr_file(),
    }))
    code, line = fire(c)
    assert code == 3
    assert line.startswith(("ERROR template has no sheet", "ERROR sheet Books_CurrMnth has no column"))
