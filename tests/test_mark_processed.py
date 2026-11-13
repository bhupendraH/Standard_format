from types import SimpleNamespace

import pytest
from azure.core.exceptions import HttpResponseError, ResourceNotFoundError

import mark_processed
from mark_processed import is_valid_request_id, main, run

SAS_URL = "https://acct.blob.core.windows.net/standardize-test?sv=x&sig=SECRET"
ENV = {"AZURE_BLOB_SAS_URL": SAS_URL}


class FakeBlobClient:
    def __init__(self, metadata):  # None means the blob doesn't exist
        self._metadata = metadata

    def get_blob_properties(self):
        if self._metadata is None:
            raise ResourceNotFoundError("not found")
        return SimpleNamespace(metadata=self._metadata)


class FakeContainer:
    def __init__(self, blobs=(), dirs=()):
        self.blobs = list(blobs)  # blob names
        self.dirs = set(dirs)  # HNS directory objects (hdi_isfolder=true)
        self.uploads = []  # (name, data, overwrite, content_type)

    def list_blobs(self, name_starts_with=None, **kw):
        prefix = name_starts_with or ""
        return iter(SimpleNamespace(name=n) for n in self.blobs if n.startswith(prefix))

    def get_blob_client(self, name):
        return FakeBlobClient({"hdi_isfolder": "true"} if name in self.dirs else None)

    def upload_blob(self, name, data, overwrite=False, content_settings=None, **kw):
        if isinstance(data, str):
            data = data.encode("utf-8")
        self.uploads.append((name, data, overwrite, content_settings.content_type))


def go(argv, container, env=ENV):
    return run(argv, env, lambda url: container)


def test_valid_ids():
    assert is_valid_request_id("request_101")
    assert is_valid_request_id("request_0")


@pytest.mark.parametrize(
    "value",
    [
        "",
        "request_",
        "request_10a",
        "Request_101",
        "request_101\n",
        " request_101",
        "request_101/",
        "../request_101",
        "request_١٠١",  # Arabic-Indic digits
    ],
)
def test_invalid_ids(value):
    assert not is_valid_request_id(value)


def test_writes_marker_when_folder_has_files():
    c = FakeContainer(blobs=["request_101/placeholder.txt"])
    assert go(["request_101"], c) == (0, "WRITTEN request_101/request_101.txt")
    assert c.uploads == [
        ("request_101/request_101.txt", b"request_101 processed", True, "text/plain")
    ]


def test_writes_marker_when_hns_folder_empty():
    c = FakeContainer(dirs={"request_105"})
    assert go(["request_105"], c) == (0, "WRITTEN request_105/request_105.txt")
    assert len(c.uploads) == 1


def test_not_found_writes_nothing():
    c = FakeContainer(blobs=["request_101/placeholder.txt"])
    assert go(["request_999"], c) == (2, "NOT_FOUND request_999")
    assert c.uploads == []


def test_prefix_collision_not_found():
    c = FakeContainer(blobs=["request_101/placeholder.txt"])
    assert go(["request_10"], c) == (2, "NOT_FOUND request_10")
    assert c.uploads == []


@pytest.mark.parametrize("argv", [["abc"], [], ["request_1", "request_2"]])
def test_invalid_id_touches_nothing(argv):
    def make_container(url):
        raise AssertionError("container must not be created")

    assert run(argv, ENV, make_container) == (1, "INVALID_ID")


def test_missing_env():
    c = FakeContainer(blobs=["request_101/placeholder.txt"])
    assert go(["request_101"], c, env={}) == (3, "ERROR AZURE_BLOB_SAS_URL not set")
    assert c.uploads == []


def test_azure_error_hides_secret():
    class FailingContainer(FakeContainer):
        def list_blobs(self, **kw):
            err = HttpResponseError(message=f"AuthenticationFailed {SAS_URL}")
            err.status_code = 403
            raise err

    code, line = go(["request_101"], FailingContainer())
    assert code == 3
    assert line.startswith("ERROR HttpResponseError")
    assert "403" in line
    assert "SECRET" not in line
    assert "https://" not in line


def test_main_prints_one_line(capsys, monkeypatch):
    monkeypatch.setattr(mark_processed, "run", lambda argv, environ: (2, "NOT_FOUND request_9"))
    with pytest.raises(SystemExit) as exc:
        main()
    assert exc.value.code == 2
    assert capsys.readouterr().out == "NOT_FOUND request_9\n"
