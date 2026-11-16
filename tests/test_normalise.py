from datetime import date, datetime

import pytest

from gstrecon.model import InputError
from gstrecon.normalise import (
    EMPTY_DATE,
    excel_serial,
    is_gstin,
    normalise_number,
    parse_date,
    state_code,
    tax_period,
    to_float,
)


@pytest.mark.parametrize(
    "raw,expected",
    [
        ("DN/003/25-26", "DN32526"),
        ("NI-0412391", "NI412391"),
        ("12/2024-25", 12202425),
        ("12/2024-25..", 12202425),
        ("jmjacnh2425feb06", "JMJACNH2425FEB6"),
        ("GTICD1168657", "GTICD1168657"),
        ("3470", 3470),
        ("INV-001", "INV1"),
        ("", ""),
    ],
)
def test_normalise_number(raw, expected):
    assert normalise_number(raw) == expected


def test_normalise_number_accepts_numbers():
    assert normalise_number(3470) == 3470


def test_is_gstin():
    assert is_gstin("27AADCA0878J1ZI")
    assert is_gstin(" 27aadca0878j1zi ")
    assert not is_gstin("UNREGISTERED")
    assert not is_gstin("")
    assert not is_gstin(None)


def test_parse_date_forms():
    expected = date(2025, 2, 28)
    assert parse_date("28-Feb-2025") == expected
    assert parse_date("28/02/2025") == expected
    assert parse_date("2025-02-28") == expected
    assert parse_date(datetime(2025, 2, 28, 0, 0)) == expected
    assert parse_date(expected) == expected
    assert parse_date(None) is None
    assert parse_date("  ") is None


def test_parse_date_rejects_garbage():
    with pytest.raises(InputError):
        parse_date("someday")


def test_to_float():
    assert to_float(None) == 0.0
    assert to_float("") == 0.0
    assert to_float("1,234.50") == 1234.5
    assert to_float(7) == 7.0


def test_state_code():
    assert state_code("Maharashtra") == 27
    assert state_code("DELHI") == 7
    assert state_code("", "24AAJCK5894N1ZQ") == 24
    assert state_code(None, "UNREGISTERED") is None


def test_tax_period():
    assert tax_period(2025, 3) == 32025
    assert tax_period(2024, 12) == 122024


def test_excel_serial():
    assert excel_serial(EMPTY_DATE) == 2
    assert excel_serial(date(2025, 3, 6)) == 45722
    assert excel_serial(datetime(2025, 3, 6, 12, 0)) == 45722
