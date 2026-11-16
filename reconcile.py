"""Reconcile books (purchase register) against GSTR-2B for a transaction folder in Blob.

Usage:
    python reconcile.py <transaction_id>        e.g. GSTR2B

Reads  <txn>/config/Output_template.xlsx (+ optional <txn>/config/*.json),
       <txn>/input/Gstr2b/*.xlsx and <txn>/input/PR/*.xlsx.
Writes <txn>/output/<GSTR-2B name with "GSTR2B" -> "Recon"> per month that has
       both files, and <txn>/output/run_summary.txt.

Environment:
    AZURE_BLOB_SAS_URL  container SAS URL (Read, Write, List). Never printed.

Prints exactly one line and exits with:
    0  DONE <n> month(s): <Mon YYYY>, ...
    1  INVALID_TRANSACTION
    2  NOT_FOUND <txn>
    3  ERROR <reason>
    4  NOTHING_TO_DO
"""
from __future__ import annotations

import json
import os
import re
import sys
from datetime import datetime
from typing import Any, Callable, Mapping

from azure.storage.blob import ContainerClient, ContentSettings

from gstrecon.model import InputError
from gstrecon.normalise import IST
from gstrecon.readers import month_from_filename, read_gstr2b, read_pr
from gstrecon.sheets import MonthInputs, build_month
from gstrecon.xlsx_writer import TemplateError, fill_template
from mark_processed import SAS_ENV, azure_error_line

_TXN = re.compile(r"[A-Za-z0-9_-]+(/[A-Za-z0-9_-]+){0,2}")
MAX_FILES = 12
XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_ABBR = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
SIDE_NAME = {"Gstr2b": "GSTR-2B", "PR": "purchase register"}


def label(month) -> str:
    return f"{_ABBR[month[1] - 1]} {month[0]}"


def _shift(month, delta):
    y, m = month
    m += delta
    return (y + (m - 1) // 12, (m - 1) % 12 + 1)


def _children(names, prefix, ext):
    return sorted(n for n in names if n.startswith(prefix) and "/" not in n[len(prefix):]
                  and n.lower().endswith(ext))


def _basename(name):
    return name.rsplit("/", 1)[-1]


def output_name(g2b_blob: str) -> str:
    base = _basename(g2b_blob)
    new, n = re.subn(r"gstr-?2b", "Recon", base, count=1, flags=re.I)
    return new if n else f"Recon {base}"


def _download(container, name) -> bytes:
    return container.download_blob(name).readall()


def _upload(container, name, data: bytes, content_type: str):
    container.upload_blob(name, data, overwrite=True,
                          content_settings=ContentSettings(content_type=content_type))


def _tolerance(container, names, txn):
    configs = _children(names, f"{txn}/config/", ".json")
    if not configs:
        return 5.0
    cfg_name = configs[0]
    try:
        cfg = json.loads(_download(container, cfg_name).decode("utf-8"))
        tol = cfg["tax_tolerance"]
        if set(cfg) != {"tax_tolerance"} or isinstance(tol, bool) \
                or not isinstance(tol, (int, float)) or tol < 0:
            raise ValueError
    except (ValueError, KeyError, TypeError, UnicodeDecodeError):
        raise _RunError(f"ERROR invalid config {_basename(cfg_name)}") from None
    return float(tol)


class _RunError(Exception):
    pass


def _load_side(container, files, side, reader, lines):
    """Read every file of one side. Returns {month: [(blob, parsed)]}."""
    by_month = {}
    for name in files:
        short = name.split("/input/", 1)[-1]
        try:
            parsed = reader(_download(container, name))
        except InputError as exc:
            lines["unusable"].append(f"{short}: {exc}")
            continue
        named = month_from_filename(name)
        if named is not None and named != parsed.month:
            lines["unusable"].append(
                f"{short}: file name says {label(named)} but contents are {label(parsed.month)}")
            continue
        by_month.setdefault(parsed.month, []).append((name, parsed))
    return by_month


def _reconcile(container, txn: str, now: datetime):
    names = [b.name for b in container.list_blobs(name_starts_with=f"{txn}/")]
    template_name = f"{txn}/config/Output_template.xlsx"
    if template_name not in names:
        return 2, f"NOT_FOUND {txn}"

    files = {}
    for side in ("Gstr2b", "PR"):
        files[side] = _children(names, f"{txn}/input/{side}/", ".xlsx")
        if len(files[side]) > MAX_FILES:
            return 3, f"ERROR too many files in input/{side} ({len(files[side])} > {MAX_FILES})"
    try:
        tolerance = _tolerance(container, names, txn)
    except _RunError as exc:
        return 3, str(exc)

    lines = {"unusable": [], "skipped": [], "produced": [], "ignored": []}
    g2b = _load_side(container, files["Gstr2b"], "Gstr2b", read_gstr2b, lines)
    pr = _load_side(container, files["PR"], "PR", read_pr, lines)

    def single(by_month, month):
        entries = by_month.get(month) or []
        return entries[0][1] if len(entries) == 1 else None

    template = _download(container, template_name)
    outputs, done, result = [], [], None
    for month in sorted(set(g2b) | set(pr)):
        dupes = [SIDE_NAME[s] for s, d in (("Gstr2b", g2b), ("PR", pr)) if len(d.get(month, ())) > 1]
        if dupes:
            lines["skipped"].append(f"{label(month)}: two {' and '.join(dupes)} files for this month")
            continue
        if month not in g2b:
            lines["skipped"].append(f"{label(month)}: purchase register only, no GSTR-2B")
            continue
        if month not in pr:
            lines["skipped"].append(f"{label(month)}: GSTR-2B only, no purchase register")
            continue
        g2b_blob, g2b_file = g2b[month][0]
        inputs = MonthInputs(
            pr=pr[month][0][1], g2b=g2b_file,
            pr_prev=single(pr, _shift(month, -1)), pr_next=single(pr, _shift(month, 1)),
            g2b_prev=single(g2b, _shift(month, -1)))
        data = build_month(inputs, now.date(), tolerance)
        try:
            xlsx = fill_template(template, data.sheets, data.header)
        except TemplateError as exc:
            result = (3, f"ERROR {exc}")
            break
        out_name = output_name(g2b_blob)
        outputs.append((f"{txn}/output/{out_name}", xlsx))
        done.append(label(month))
        counts = ", ".join(f"{k} {v}" for k, v in data.counts.items())
        lines["produced"].append(f"{label(month)} -> {out_name}\n    {counts}")
        lines["produced"] += [f"    warning: {w}" for w in data.warnings]
        lines["ignored"] += [f"{label(month)} sheet {s}: {n} row(s) not reconciled"
                             for s, n in g2b_file.ignored.items()]

    if result is None:
        for name, xlsx in outputs:
            _upload(container, name, xlsx, XLSX_TYPE)
        result = (0, f"DONE {len(done)} month(s): {', '.join(done)}") if done else (4, "NOTHING_TO_DO")

    summary = [f"GSTR-2B reconciliation run {now.astimezone(IST):%Y-%m-%d %H:%M} IST",
               f"Transaction: {txn}", f"Tax tolerance: Rs {tolerance:.2f}", ""]
    for title, key in (("Produced", "produced"), ("Skipped months", "skipped"),
                       ("Unusable files", "unusable"), ("GSTR-2B sheets not reconciled", "ignored")):
        summary.append(f"{title}:")
        summary += [f"  {x}" for x in lines[key]] or ["  none"]
        summary.append("")
    summary.append(f"Result: {result[1]}")
    _upload(container, f"{txn}/output/run_summary.txt", ("\n".join(summary) + "\n").encode("utf-8"),
            "text/plain; charset=utf-8")
    return result


def run(
    argv: list,
    environ: Mapping[str, str],
    make_container: Callable[[str], Any] = ContainerClient.from_container_url,
    now: datetime = None,
) -> tuple:
    if len(argv) != 1 or _TXN.fullmatch(argv[0] or "") is None:
        return 1, "INVALID_TRANSACTION"
    sas_url = environ.get(SAS_ENV)
    if not sas_url:
        return 3, f"ERROR {SAS_ENV} not set"
    try:
        container = make_container(sas_url)
        return _reconcile(container, argv[0], now or datetime.now(IST))
    except Exception as exc:  # never echo str(exc): it may contain the SAS URL
        return 3, azure_error_line(exc)


def main() -> None:
    code, line = run(sys.argv[1:], os.environ)
    print(line)
    sys.exit(code)


if __name__ == "__main__":
    main()
