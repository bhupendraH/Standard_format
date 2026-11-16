"""Small conversions shared by the readers and the sheet builder."""
from __future__ import annotations

import re
from datetime import date, datetime, timedelta, timezone
from typing import Optional, Union

from .model import InputError

EMPTY_DATE = date(1900, 1, 1)
IST = timezone(timedelta(hours=5, minutes=30))
_EXCEL_EPOCH = date(1899, 12, 30)
_GSTIN = re.compile(r"[0-9]{2}[A-Z0-9]{13}")
_DATE_FORMATS = ("%d-%b-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d-%b-%y")

STATE_CODES = {
    "JAMMU AND KASHMIR": 1, "HIMACHAL PRADESH": 2, "PUNJAB": 3, "CHANDIGARH": 4,
    "UTTARAKHAND": 5, "HARYANA": 6, "DELHI": 7, "RAJASTHAN": 8, "UTTAR PRADESH": 9,
    "BIHAR": 10, "SIKKIM": 11, "ARUNACHAL PRADESH": 12, "NAGALAND": 13, "MANIPUR": 14,
    "MIZORAM": 15, "TRIPURA": 16, "MEGHALAYA": 17, "ASSAM": 18, "WEST BENGAL": 19,
    "JHARKHAND": 20, "ODISHA": 21, "CHHATTISGARH": 22, "MADHYA PRADESH": 23,
    "GUJARAT": 24, "DADRA AND NAGAR HAVELI AND DAMAN AND DIU": 26, "MAHARASHTRA": 27,
    "KARNATAKA": 29, "GOA": 30, "LAKSHADWEEP": 31, "KERALA": 32, "TAMIL NADU": 33,
    "PUDUCHERRY": 34, "ANDAMAN AND NICOBAR ISLANDS": 35, "TELANGANA": 36,
    "ANDHRA PRADESH": 37, "LADAKH": 38, "OTHER TERRITORY": 97,
    # common alternative spellings
    "ORISSA": 21, "PONDICHERRY": 34, "NEW DELHI": 7, "DAMAN AND DIU": 26,
    "DADRA AND NAGAR HAVELI": 26, "JAMMU & KASHMIR": 1, "ANDAMAN & NICOBAR ISLANDS": 35,
}


def normalise_number(raw) -> Union[str, int]:
    """Upper-case, keep A-Z/0-9 only, drop leading zeros of each digit run."""
    s = re.sub(r"[^A-Z0-9]", "", str(raw or "").upper())
    s = re.sub(r"(?<![0-9])0+(?=[0-9])", "", s)
    return int(s) if s.isdigit() else s


def is_gstin(value) -> bool:
    return bool(value) and _GSTIN.fullmatch(str(value).strip().upper()) is not None


def parse_date(value) -> Optional[date]:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if not text:
        return None
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise InputError(f"unreadable date: {text[:40]}")


def to_float(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        raise InputError(f"unreadable number: {text[:40]}") from None


def state_code(name, gstin=None) -> Optional[int]:
    key = re.sub(r"\s+", " ", str(name or "").strip().upper())
    if key in STATE_CODES:
        return STATE_CODES[key]
    if is_gstin(gstin):
        return int(str(gstin).strip()[:2])
    return None


def tax_period(year: int, month: int) -> int:
    return int(f"{month}{year}")


def excel_serial(d) -> int:
    if isinstance(d, datetime):
        d = d.date()
    return (d - _EXCEL_EPOCH).days
