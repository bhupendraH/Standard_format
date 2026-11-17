"""Blob I/O for one GSTR-2B reconciliation run, limited to that run's folder.

The run folder is <tenant_id>/<agent_id>/<transaction_id>/ (called RUN).
Claude does the reconciliation; this script is its only way to reach Blob.

Usage:
    python3 blob.py pull   <tenant> <agent> <txn> <local_dir>
        Download everything under RUN except its output folder into <local_dir>
        (same relative paths, blob modified time kept). <local_dir> must be
        new or empty.
    python3 blob.py put    <tenant> <agent> <txn> <local_file.xlsx>
        Upload the workbook to RUN/<output folder>/<file name>.
    python3 blob.py status <tenant> <agent> <txn> done|failed <message> [<message> ...]
        Write RUN/<output folder>/run_status.json.

<output folder> is RUN's existing output folder with its own spelling
("Output", "OUTPUT", ...), else "output". Reads stay inside RUN, writes stay
inside its output folder; nothing is ever deleted.

Environment (an Entra ID app with Storage Blob Data Contributor; never printed):
    AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET,
    AZURE_STORAGE_ACCOUNT, AZURE_STORAGE_CONTAINER
    CLAUDE_CODE_REMOTE_SESSION_ID (optional, recorded in run_status.json)

Prints exactly one line and exits with:
    0  PULLED <n> file(s) | UPLOADED <blob path>
    1  INVALID_TRANSACTION | USAGE ...
    2  NOT_FOUND <RUN>
    3  ERROR <reason>
    5  UNSUPPORTED_AGENT <agent_id>
"""
from __future__ import annotations

import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Mapping, Optional

_ID = re.compile(r"[A-Za-z0-9_-]+")
_OUT_NAME = re.compile(r"[A-Za-z0-9][A-Za-z0-9 ._()-]{0,150}\.xlsx", re.I)
_ERROR_CODE = re.compile(r"[A-Za-z]{1,64}")
SUPPORTED_AGENT = "GSTR2B"
ENV_VARS = ("AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET",
            "AZURE_STORAGE_ACCOUNT", "AZURE_STORAGE_CONTAINER")
MAX_FILES = 60
MAX_BYTES = 200 * 1024 * 1024
MAX_UPLOAD_BYTES = 50 * 1024 * 1024
XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
USAGE = "USAGE blob.py pull|put|status <tenant_id> <agent_id> <transaction_id> ..."


class _Fail(Exception):
    def __init__(self, code: int, line: str):
        super().__init__(line)
        self.code, self.line = code, line


def make_container(environ: Mapping[str, str]):
    """Container client signed in as the Entra ID app (client credentials)."""
    from azure.identity import ClientSecretCredential
    from azure.storage.blob import ContainerClient
    credential = ClientSecretCredential(
        tenant_id=environ["AZURE_TENANT_ID"],
        client_id=environ["AZURE_CLIENT_ID"],
        client_secret=environ["AZURE_CLIENT_SECRET"],
    )
    return ContainerClient(
        account_url=f"https://{environ['AZURE_STORAGE_ACCOUNT']}.blob.core.windows.net",
        container_name=environ["AZURE_STORAGE_CONTAINER"],
        credential=credential,
    )


def azure_error_line(exc: BaseException) -> str:
    """'ERROR <type> [<http status>] [<azure error code>]' - never the message."""
    reason = type(exc).__name__
    status = getattr(exc, "status_code", None)
    if status is not None:
        reason += f" {status}"
    code = getattr(exc, "error_code", None)
    code = str(getattr(code, "value", code) or "")
    if _ERROR_CODE.fullmatch(code):
        reason += f" {code}"
    return f"ERROR {reason}"


def _content_settings(content_type: str):
    from azure.storage.blob import ContentSettings
    return ContentSettings(content_type=content_type)


def _list(container, run: str) -> list:
    """Blobs under RUN/ as (relative name, blob item)."""
    return [(b.name[len(run) + 1:], b) for b in container.list_blobs(name_starts_with=f"{run}/")]


def _output_dir(rels) -> str:
    for rel in rels:
        first, sep, _ = rel.partition("/")
        if sep and first.lower() == "output":
            return first
    return "output"


def _is_output(rel: str) -> bool:
    return rel.partition("/")[0].lower() == "output"


def _safe(rel: str) -> bool:
    parts = rel.split("/")
    return all(p not in ("", ".", "..") and "\\" not in p and ":" not in p
               and not any(ord(ch) < 32 for ch in p) for p in parts)


def _pull(container, run: str, local_dir: str) -> tuple:
    listed = _list(container, run)
    rels = {rel for rel, _ in listed}
    wanted = [(rel, b) for rel, b in listed
              if not _is_output(rel)
              and not any(other.startswith(rel + "/") for other in rels)]  # folder markers
    if not wanted:
        raise _Fail(2, f"NOT_FOUND {run}")
    if not all(_safe(rel) for rel, _ in wanted):
        raise _Fail(3, "ERROR unsafe blob name in run folder")
    if len(wanted) > MAX_FILES:
        raise _Fail(3, f"ERROR too many files ({len(wanted)} > {MAX_FILES})")
    total = sum(b.size or 0 for _, b in wanted)
    if total > MAX_BYTES:
        raise _Fail(3, f"ERROR inputs too large ({total} bytes > {MAX_BYTES})")

    root = Path(local_dir)
    if root.exists() and any(root.iterdir()):
        raise _Fail(3, "ERROR local dir is not empty")
    for rel, b in wanted:
        dest = root.joinpath(*rel.split("/"))
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(container.download_blob(b.name).readall())
        if b.last_modified is not None:
            stamp = b.last_modified.timestamp()
            os.utime(dest, (stamp, stamp))
    return 0, f"PULLED {len(wanted)} file(s)"


def _existing_output_dir(container, run: str) -> str:
    rels = [rel for rel, _ in _list(container, run)]
    if not rels:
        raise _Fail(2, f"NOT_FOUND {run}")
    return _output_dir(rels)


def _put(container, run: str, local_file: str) -> tuple:
    path = Path(local_file)
    if not _OUT_NAME.fullmatch(path.name):
        raise _Fail(3, "ERROR output name not allowed (letters, digits, space . _ ( ) -; .xlsx)")
    if not path.is_file():
        raise _Fail(3, f"ERROR local file not found: {path.name}")
    data = path.read_bytes()
    if len(data) > MAX_UPLOAD_BYTES:
        raise _Fail(3, f"ERROR workbook too large ({len(data)} bytes)")
    name = f"{run}/{_existing_output_dir(container, run)}/{path.name}"
    container.upload_blob(name, data, overwrite=True,
                          content_settings=_content_settings(XLSX_TYPE))
    return 0, f"UPLOADED {name}"


def _status(container, run: str, ids, state: str, messages, environ) -> tuple:
    doc = {
        "tenant_id": ids[0], "agent_id": ids[1], "transaction_id": ids[2],
        "session_id": environ.get("CLAUDE_CODE_REMOTE_SESSION_ID", ""),
        "status": state,
        "messages": [m[:1000] for m in messages[:20]],
        "finished_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }
    name = f"{run}/{_existing_output_dir(container, run)}/run_status.json"
    container.upload_blob(name, json.dumps(doc, indent=2).encode("utf-8"), overwrite=True,
                          content_settings=_content_settings("application/json"))
    return 0, f"UPLOADED {name}"


def run(argv: list, environ: Mapping[str, str],
        make_container: Optional[Callable[[Mapping[str, str]], Any]] = None) -> tuple:
    if not argv or argv[0] not in ("pull", "put", "status"):
        return 1, USAGE
    cmd, rest = argv[0], argv[1:]
    ids, extra = rest[:3], rest[3:]
    if len(ids) != 3 or any(_ID.fullmatch(i or "") is None for i in ids):
        return 1, "INVALID_TRANSACTION"
    if (cmd in ("pull", "put") and len(extra) != 1) or \
            (cmd == "status" and (len(extra) < 2 or extra[0] not in ("done", "failed"))):
        return 1, USAGE
    if ids[1].upper() != SUPPORTED_AGENT:
        return 5, f"UNSUPPORTED_AGENT {ids[1]}"
    for name in ENV_VARS:
        if not environ.get(name):
            return 3, f"ERROR {name} not set"

    factory = make_container or globals()["make_container"]
    run_dir = "/".join(ids)
    try:
        container = factory(environ)
        if cmd == "pull":
            return _pull(container, run_dir, extra[0])
        if cmd == "put":
            return _put(container, run_dir, extra[0])
        return _status(container, run_dir, ids, extra[0], extra[1:], environ)
    except _Fail as exc:
        return exc.code, exc.line
    except Exception as exc:  # never echo str(exc): it may contain secrets or URLs
        return 3, azure_error_line(exc)


def main() -> None:
    code, line = run(sys.argv[1:], os.environ)
    print(line)
    sys.exit(code)


if __name__ == "__main__":
    main()
