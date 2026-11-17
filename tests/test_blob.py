import json
import os
from datetime import datetime, timezone
from types import SimpleNamespace

import pytest
from azure.core.exceptions import HttpResponseError

import blob
from blob import run

SECRET = "s3cr3t~value"
ENV = {"AZURE_TENANT_ID": "tid", "AZURE_CLIENT_ID": "cid", "AZURE_CLIENT_SECRET": SECRET,
       "AZURE_STORAGE_ACCOUNT": "acct", "AZURE_STORAGE_CONTAINER": "standardize-test"}
IDS = ["T1", "GSTR2B", "TXN-001"]
RUN = "T1/GSTR2B/TXN-001"
STAMP = datetime(2025, 3, 5, 10, 0, tzinfo=timezone.utc)


class FakeContainer:
    def __init__(self, blobs, stamps=None):
        self.blobs = dict(blobs)
        self.stamps = stamps or {}
        self.uploads = {}

    def list_blobs(self, name_starts_with=None, **kw):
        prefix = name_starts_with or ""
        return iter(SimpleNamespace(name=n, size=len(self.blobs[n]),
                                    last_modified=self.stamps.get(n, STAMP))
                    for n in sorted(self.blobs) if n.startswith(prefix))

    def download_blob(self, name, **kw):
        data = self.blobs[name]
        return SimpleNamespace(readall=lambda: data)

    def upload_blob(self, name, data, overwrite=False, content_settings=None, **kw):
        assert overwrite is True
        self.uploads[name] = (data, content_settings.content_type)


def call(argv, container, env=ENV):
    return run(argv, env, make_container=lambda environ: container)


def inputs():
    return {f"{RUN}/Input/config.json": b"{}",
            f"{RUN}/Input/Outputformat.xlsx": b"tpl",
            f"{RUN}/Input/PR/PR 022025.xlsx": b"pr",
            f"{RUN}/Input/GSTR/2B Feb.xlsx": b"g2b"}


# --- argument and environment checks -------------------------------------------------

@pytest.mark.parametrize("ids", [["T1", "GSTR2B"], ["T1", "GSTR2B", "../x"],
                                 ["T 1", "GSTR2B", "X"], ["T1", "GSTR2B", ""]])
def test_bad_ids_are_rejected_without_touching_blob(ids, tmp_path):
    def boom(environ):
        raise AssertionError("must not connect")
    code, line = run(["pull", *ids, str(tmp_path)], ENV, make_container=boom)
    assert (code, line) == (1, "INVALID_TRANSACTION")


def test_unknown_command_is_rejected():
    assert call(["delete", *IDS], FakeContainer({}))[0] == 1


def test_other_agents_are_unsupported(tmp_path):
    assert call(["pull", "T1", "LEDGER", "X", str(tmp_path)], FakeContainer({})) == \
        (5, "UNSUPPORTED_AGENT LEDGER")


def test_agent_id_is_matched_ignoring_case(tmp_path):
    c = FakeContainer({"T1/gstr2b/X/Input/a.xlsx": b"a"})
    assert call(["pull", "T1", "gstr2b", "X", str(tmp_path)], c)[0] == 0


def test_missing_env_var_is_an_error(tmp_path):
    env = dict(ENV, AZURE_CLIENT_SECRET="")
    assert call(["pull", *IDS, str(tmp_path)], FakeContainer({}), env) == \
        (3, "ERROR AZURE_CLIENT_SECRET not set")


def test_azure_errors_never_echo_the_message(tmp_path):
    class Exploding(FakeContainer):
        def list_blobs(self, **kw):
            exc = HttpResponseError(message=f"secret {SECRET} https://acct.blob/x?sig=1")
            exc.status_code = 403
            exc.error_code = "AuthorizationPermissionMismatch"
            raise exc
    code, line = call(["pull", *IDS, str(tmp_path)], Exploding({}))
    assert (code, line) == (3, "ERROR HttpResponseError 403 AuthorizationPermissionMismatch")


# --- pull ----------------------------------------------------------------------------

def test_pull_mirrors_the_run_folder_and_skips_output(tmp_path):
    blobs = inputs()
    blobs[f"{RUN}/Output/old.xlsx"] = b"old"
    blobs["T1/GSTR2B/TXN-0011/Input/other.xlsx"] = b"neighbour"   # same prefix, other txn
    blobs["T1/GSTR2B/OTHER/Input/x.xlsx"] = b"other"
    code, line = call(["pull", *IDS, str(tmp_path / "work")], FakeContainer(blobs))
    assert (code, line) == (0, "PULLED 4 file(s)")
    got = sorted(p.relative_to(tmp_path / "work").as_posix()
                 for p in (tmp_path / "work").rglob("*") if p.is_file())
    assert got == ["Input/GSTR/2B Feb.xlsx", "Input/Outputformat.xlsx",
                   "Input/PR/PR 022025.xlsx", "Input/config.json"]
    assert (tmp_path / "work/Input/PR/PR 022025.xlsx").read_bytes() == b"pr"


def test_pull_keeps_the_blob_modified_time(tmp_path):
    name = f"{RUN}/Input/PR/a.xlsx"
    later = datetime(2025, 4, 1, 8, 0, tzinfo=timezone.utc)
    call(["pull", *IDS, str(tmp_path)], FakeContainer({name: b"a"}, {name: later}))
    assert os.path.getmtime(tmp_path / "Input/PR/a.xlsx") == later.timestamp()


def test_pull_skips_directory_marker_blobs(tmp_path):
    blobs = {f"{RUN}/Input": b"", f"{RUN}/Input/PR": b"", f"{RUN}/Input/PR/a.xlsx": b"a"}
    assert call(["pull", *IDS, str(tmp_path)], FakeContainer(blobs)) == (0, "PULLED 1 file(s)")


def test_pull_of_missing_folder_is_not_found(tmp_path):
    c = FakeContainer({"T1/GSTR2B/OTHER/Input/x.xlsx": b"x"})
    assert call(["pull", *IDS, str(tmp_path)], c) == (2, f"NOT_FOUND {RUN}")


def test_pull_with_only_an_output_folder_is_not_found(tmp_path):
    c = FakeContainer({f"{RUN}/Output/old.xlsx": b"x"})
    assert call(["pull", *IDS, str(tmp_path)], c) == (2, f"NOT_FOUND {RUN}")


@pytest.mark.parametrize("bad", ["Input/../../evil.py", "Input/./x", "Input//x", "Input/a\\b"])
def test_pull_refuses_unsafe_blob_names(bad, tmp_path):
    c = FakeContainer({f"{RUN}/{bad}": b"x", f"{RUN}/Input/a.xlsx": b"a"})
    code, line = call(["pull", *IDS, str(tmp_path / "w")], c)
    assert code == 3 and line.startswith("ERROR unsafe blob name")
    assert not (tmp_path / "evil.py").exists()


def test_pull_refuses_too_many_files(tmp_path, monkeypatch):
    monkeypatch.setattr(blob, "MAX_FILES", 2)
    assert call(["pull", *IDS, str(tmp_path)], FakeContainer(inputs())) == \
        (3, "ERROR too many files (4 > 2)")


def test_pull_refuses_too_much_data(tmp_path, monkeypatch):
    monkeypatch.setattr(blob, "MAX_BYTES", 5)
    code, line = call(["pull", *IDS, str(tmp_path)], FakeContainer(inputs()))
    assert (code, line) == (3, "ERROR inputs too large (10 bytes > 5)")


def test_pull_needs_a_local_dir():
    assert call(["pull", *IDS], FakeContainer(inputs()))[0] == 1


# --- put -----------------------------------------------------------------------------

def test_put_uploads_into_the_existing_output_folder(tmp_path):
    f = tmp_path / "PR_vs_GSTR2B_022025_042025.xlsx"
    f.write_bytes(b"book")
    c = FakeContainer(dict(inputs(), **{f"{RUN}/Output/old.txt": b"x"}))
    assert call(["put", *IDS, str(f)], c) == \
        (0, f"UPLOADED {RUN}/Output/PR_vs_GSTR2B_022025_042025.xlsx")
    assert c.uploads[f"{RUN}/Output/PR_vs_GSTR2B_022025_042025.xlsx"] == (b"book", blob.XLSX_TYPE)


def test_put_creates_lowercase_output_when_none_exists(tmp_path):
    f = tmp_path / "r.xlsx"
    f.write_bytes(b"book")
    c = FakeContainer(inputs())
    assert call(["put", *IDS, str(f)], c) == (0, f"UPLOADED {RUN}/output/r.xlsx")


@pytest.mark.parametrize("name", ["r.csv", ".x.xlsx", "a;b.xlsx", "run_status.json"])
def test_put_only_accepts_plain_xlsx_names(name, tmp_path):
    f = tmp_path / name
    f.write_bytes(b"x")
    c = FakeContainer(inputs())
    code, line = call(["put", *IDS, str(f)], c)
    assert code == 3 and line.startswith("ERROR")
    assert c.uploads == {}


def test_put_of_missing_local_file_is_an_error(tmp_path):
    code, line = call(["put", *IDS, str(tmp_path / "nope.xlsx")], FakeContainer(inputs()))
    assert (code, line) == (3, "ERROR local file not found: nope.xlsx")


def test_put_refuses_a_missing_run_folder(tmp_path):
    f = tmp_path / "r.xlsx"
    f.write_bytes(b"x")
    c = FakeContainer({})
    assert call(["put", *IDS, str(f)], c) == (2, f"NOT_FOUND {RUN}")
    assert c.uploads == {}


# --- status --------------------------------------------------------------------------

def test_status_writes_run_status_json():
    c = FakeContainer(inputs())
    env = dict(ENV, CLAUDE_CODE_REMOTE_SESSION_ID="cse_01abc")
    code, line = call(["status", *IDS, "failed", "ERROR no PR files", "month 03"], c, env)
    assert (code, line) == (0, f"UPLOADED {RUN}/output/run_status.json")
    data, ctype = c.uploads[f"{RUN}/output/run_status.json"]
    assert ctype == "application/json"
    doc = json.loads(data)
    assert doc["tenant_id"] == "T1" and doc["agent_id"] == "GSTR2B"
    assert doc["transaction_id"] == "TXN-001" and doc["session_id"] == "cse_01abc"
    assert doc["status"] == "failed" and doc["messages"] == ["ERROR no PR files", "month 03"]
    assert "finished_at" in doc


def test_status_only_accepts_done_or_failed():
    c = FakeContainer(inputs())
    assert call(["status", *IDS, "maybe", "x"], c)[0] == 1
    assert c.uploads == {}


def test_status_never_creates_a_missing_run_folder():
    c = FakeContainer({})
    assert call(["status", *IDS, "failed", "x"], c) == (2, f"NOT_FOUND {RUN}")
    assert c.uploads == {}


def test_main_prints_one_line_and_exits_with_code(capsys, monkeypatch):
    monkeypatch.setattr(blob, "make_container", lambda environ: FakeContainer({}))
    monkeypatch.setattr("sys.argv", ["blob.py", "pull", "bad id", "x", "y", "z"])
    for k, v in ENV.items():
        monkeypatch.setenv(k, v)
    with pytest.raises(SystemExit) as exc:
        blob.main()
    assert exc.value.code == 1
    assert capsys.readouterr().out == "INVALID_TRANSACTION\n"
