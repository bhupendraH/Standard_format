"""Write a marker file into a request's folder in Azure Blob Storage.

Usage:
    python mark_processed.py <request_id>

<request_id> must be "request_" followed by ASCII digits, e.g. request_101.
If the folder <request_id>/ exists in the container, writes
<request_id>/<request_id>.txt containing "<request_id> processed".

Environment:
    AZURE_BLOB_SAS_URL  container SAS URL (Read, Write, List). Never printed.

Prints exactly one line and exits with:
    0  WRITTEN <id>/<id>.txt
    1  INVALID_ID
    2  NOT_FOUND <id>
    3  ERROR <reason>
"""
from __future__ import annotations

import os
import re
import sys
from typing import Any, Callable, Mapping

from azure.core.exceptions import ResourceNotFoundError
from azure.storage.blob import ContainerClient, ContentSettings

SAS_ENV = "AZURE_BLOB_SAS_URL"
_ID_PATTERN = re.compile(r"request_[0-9]+")


def is_valid_request_id(value: str) -> bool:
    return _ID_PATTERN.fullmatch(value) is not None


def folder_exists(container, request_id: str) -> bool:
    """True if any blob sits under <id>/, or <id> is a directory object (HNS)."""
    if next(iter(container.list_blobs(name_starts_with=f"{request_id}/")), None) is not None:
        return True
    try:
        props = container.get_blob_client(request_id).get_blob_properties()
    except ResourceNotFoundError:
        return False
    return (props.metadata or {}).get("hdi_isfolder") == "true"


def write_marker(container, request_id: str) -> str:
    name = f"{request_id}/{request_id}.txt"
    container.upload_blob(
        name,
        f"{request_id} processed".encode("utf-8"),
        overwrite=True,
        content_settings=ContentSettings(content_type="text/plain"),
    )
    return name


def run(
    argv: list[str],
    environ: Mapping[str, str],
    make_container: Callable[[str], Any] = ContainerClient.from_container_url,
) -> tuple[int, str]:
    if len(argv) != 1 or not is_valid_request_id(argv[0]):
        return 1, "INVALID_ID"
    request_id = argv[0]

    sas_url = environ.get(SAS_ENV)
    if not sas_url:
        return 3, f"ERROR {SAS_ENV} not set"

    try:
        container = make_container(sas_url)
        if not folder_exists(container, request_id):
            return 2, f"NOT_FOUND {request_id}"
        return 0, f"WRITTEN {write_marker(container, request_id)}"
    except Exception as exc:  # never echo str(exc): it may contain the SAS URL
        reason = type(exc).__name__
        status = getattr(exc, "status_code", None)
        if status is not None:
            reason += f" {status}"
        return 3, f"ERROR {reason}"


def main() -> None:
    code, line = run(sys.argv[1:], os.environ)
    print(line)
    sys.exit(code)


if __name__ == "__main__":
    main()
