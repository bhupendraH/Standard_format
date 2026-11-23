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
        1. Write RUN/<output folder>/run_status.json (all messages), unless the
           run folder is missing or the agent is unsupported.
        2. Then, whatever step 1 returned, close the backend execution: POST to
           STATUS_API_URL (update_agent_execution_log) with execution_id = <txn>,
           SUCCESS for done, FAIL for failed with the first message (max 255
           chars) as agent_fail_reason.

<output folder> is RUN's existing output folder with its own spelling
("Output", "OUTPUT", ...), else "output". Reads stay inside RUN, writes stay
inside its output folder; nothing is ever deleted.

Environment (an Entra ID app with Storage Blob Data Contributor):
    AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_STORAGE_ACCOUNT, AZURE_STORAGE_CONTAINER
    AZURE_CLIENT_SECRET (optional: in the cloud it is a network secret that the
        sandbox proxy adds to requests to login.microsoftonline.com, so the run
        never sees it; set it in the shell for local use. Never printed.)
    CLAUDE_CODE_REMOTE_SESSION_ID (optional, recorded in run_status.json)
    STATUS_API_URL (optional: the backend's update_agent_execution_log URL;
        not set = the API is not called)
    ROUTINE_UPDATE_KEY (optional: sent as the X-Routine-Key header to
        STATUS_API_URL; never printed. Omit it if a network secret adds the header.)

Prints exactly one line and exits with:
    0  PULLED <n> file(s) | UPLOADED <blob path> | UPLOADED <path>; NOTIFIED
    1  INVALID_TRANSACTION | USAGE ...
    2  NOT_FOUND <RUN>
    3  ERROR <reason>
    4  UPLOADED <path>; ERROR status API failed (<reason>)
    5  UNSUPPORTED_AGENT <agent_id>
`status` lines are "<blob part>; <api part>", where the api part is NOTIFIED,
"STATUS_API_URL not set, API not called" or "ERROR status API failed (...)".
Its exit code is the blob part's (2/3/5) if that failed, else 4 if the API
failed, else 0.
"""
from __future__ import annotations

import json
import os
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Mapping, Optional

import requests

_ID = re.compile(r"[A-Za-z0-9_-]+")
_OUT_NAME = re.compile(r"[A-Za-z0-9][A-Za-z0-9 ._()-]{0,150}\.xlsx", re.I)
_ERROR_CODE = re.compile(r"[A-Za-z0-9_]{1,64}")
SUPPORTED_AGENT = "GSTR2B"
ENV_VARS = ("AZURE_TENANT_ID", "AZURE_CLIENT_ID",
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


class TokenError(Exception):
    """Token request failed. Carries only the HTTP status and Entra error code."""

    def __init__(self, status_code: int, error_code: str):
        super().__init__(f"token request failed: {status_code} {error_code}")
        self.status_code, self.error_code = status_code, error_code


class AppTokenCredential:
    """Client-credentials token for the Entra ID app, usable by the Azure SDK.

    The client secret is sent only if AZURE_CLIENT_SECRET is set (local use).
    In the cloud it is a network secret: the sandbox proxy adds it to requests
    to login.microsoftonline.com, so we send no secret and no Authorization header.
    """

    def __init__(self, environ: Mapping[str, str]):
        self._tenant = environ["AZURE_TENANT_ID"]
        self._client = environ["AZURE_CLIENT_ID"]
        self._secret = environ.get("AZURE_CLIENT_SECRET") or ""
        self._cache = {}

    def get_token(self, *scopes, **kwargs):
        from azure.core.credentials import AccessToken
        scope = " ".join(scopes)
        cached = self._cache.get(scope)
        if cached and cached.expires_on - 300 > time.time():
            return cached
        form = {"grant_type": "client_credentials", "client_id": self._client, "scope": scope}
        if self._secret:
            form["client_secret"] = self._secret
        resp = requests.post(
            f"https://login.microsoftonline.com/{self._tenant}/oauth2/v2.0/token",
            data=form, headers={"Accept": "application/json"}, timeout=60)
        try:
            body = resp.json()
        except ValueError:
            body = {}
        if resp.status_code != 200 or not body.get("access_token"):
            codes = body.get("error_codes") or []
            code = f"AADSTS{codes[0]}" if codes else str(body.get("error") or "")
            raise TokenError(resp.status_code, code)
        token = AccessToken(body["access_token"], int(time.time()) + int(body.get("expires_in", 3599)))
        self._cache[scope] = token
        return token


def make_container(environ: Mapping[str, str]):
    """Container client signed in as the Entra ID app (client credentials)."""
    from azure.storage.blob import ContainerClient
    credential = AppTokenCredential(environ)
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


def execution_body(execution_id: str, state: str, messages) -> dict:
    """Body for the backend's update_agent_execution_log API."""
    if state == "done":
        return {"execution_id": execution_id, "agent_execution_status": "SUCCESS"}
    reason = (messages[0] if messages else "failed")[:255]
    return {"execution_id": execution_id, "agent_execution_status": "FAIL",
            "agent_fail_reason": reason}


def _body_status_code(resp):
    """The 'status_code' inside a JSON response body: an int, None if there is no
    such field (or no JSON body), or 'invalid' if it is present but not a number."""
    try:
        body = resp.json()
    except Exception:
        return None
    if not isinstance(body, dict) or "status_code" not in body:
        return None
    try:
        return int(body["status_code"])
    except (TypeError, ValueError):
        return "invalid"


def notify_api(url: str, body: dict, headers: dict, post: Callable = None,
               sleep: Callable = time.sleep) -> str:
    """POST the result to the backend. Returns '' on success, else a short reason.

    Success = HTTP 2xx and, if the JSON body has a status_code, that it is 200.
    Retries connection errors and 5xx (3 attempts). Never echoes response bodies,
    headers or exception messages.
    """
    post = post or requests.post
    reason = ""
    for attempt in range(3):
        if attempt:
            sleep(2 * attempt)
        try:
            # No redirects: a redirect would resend as GET and carry the key elsewhere.
            resp = post(url, json=body, headers=headers, timeout=30, allow_redirects=False)
        except Exception as exc:
            reason = type(exc).__name__
            continue
        if 200 <= resp.status_code < 300:
            inner = _body_status_code(resp)
            if inner is None or inner == 200:
                return ""
            return f"API status_code {inner}"
        reason = f"HTTP {resp.status_code}"
        if resp.status_code < 500:
            break
    return reason


def _write_status_blob(ids, state: str, messages, environ, factory) -> tuple:
    """Blob part of `status`: write RUN/<output>/run_status.json. Never raises."""
    if ids[1].upper() != SUPPORTED_AGENT:
        return 5, f"UNSUPPORTED_AGENT {ids[1]}"
    for name in ENV_VARS:
        if not environ.get(name):
            return 3, f"ERROR {name} not set"
    run_dir = "/".join(ids)
    doc = {
        "tenant_id": ids[0], "agent_id": ids[1], "transaction_id": ids[2],
        "session_id": environ.get("CLAUDE_CODE_REMOTE_SESSION_ID", ""),
        "status": state,
        "messages": [m[:1000] for m in messages[:20]],
        "finished_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }
    try:
        container = factory(environ)
        name = f"{run_dir}/{_existing_output_dir(container, run_dir)}/run_status.json"
        container.upload_blob(name, json.dumps(doc, indent=2).encode("utf-8"), overwrite=True,
                              content_settings=_content_settings("application/json"))
        return 0, f"UPLOADED {name}"
    except _Fail as exc:
        return exc.code, exc.line
    except Exception as exc:  # never echo str(exc)
        return 3, azure_error_line(exc)


def _status(ids, state: str, messages, environ, factory) -> tuple:
    """Write run_status.json, then close the backend execution (execution_id = txn).

    The API part runs whatever the Blob part returned, so the execution is closed
    even when the run folder is missing or Blob fails.
    """
    code, line = _write_status_blob(ids, state, messages, environ, factory)
    url = environ.get("STATUS_API_URL", "")
    if not url:
        return code, f"{line}; STATUS_API_URL not set, API not called"
    headers = {"Accept": "application/json"}
    key = environ.get("ROUTINE_UPDATE_KEY", "")
    if key:
        headers["X-Routine-Key"] = key
    reason = notify_api(url, execution_body(ids[2], state, messages), headers)
    if reason:
        return (code or 4), f"{line}; ERROR status API failed ({reason})"
    return code, f"{line}; NOTIFIED"


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
    factory = make_container or globals()["make_container"]
    if cmd == "status":
        return _status(ids, extra[0], extra[1:], environ, factory)
    if ids[1].upper() != SUPPORTED_AGENT:
        return 5, f"UNSUPPORTED_AGENT {ids[1]}"
    for name in ENV_VARS:
        if not environ.get(name):
            return 3, f"ERROR {name} not set"

    run_dir = "/".join(ids)
    try:
        container = factory(environ)
        if cmd == "pull":
            return _pull(container, run_dir, extra[0])
        return _put(container, run_dir, extra[0])
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
