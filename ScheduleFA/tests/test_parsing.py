"""Tests for low-level parsing helpers."""
import datetime

import pytest

from dividend_contribution.parsing import parse_amount, parse_date, norm


@pytest.mark.parametrize(
    "value, expected",
    [
        (1234.5, 1234.5),
        (10, 10.0),
        ("\u20b91,234.50", 1234.5),
        ("$2,000", 2000.0),
        ("250.75", 250.75),
        ("-7.25", -7.25),
        ("12%", 12.0),
        ("\u20b9499.99", 499.99),
        (None, None),
        ("", None),
        ("-", None),
        ("n/a", None),
        ("abc", None),
        (True, None),
    ],
)
def test_parse_amount(value, expected):
    assert parse_amount(value) == expected


@pytest.mark.parametrize(
    "value, expected",
    [
        (datetime.datetime(2025, 7, 19), datetime.date(2025, 7, 19)),
        (datetime.date(2025, 7, 19), datetime.date(2025, 7, 19)),
        ("2025-07-19", datetime.date(2025, 7, 19)),
        ("19/07/2025", datetime.date(2025, 7, 19)),
        ("04/03/2025", datetime.date(2025, 3, 4)),
        ("4/3/2025", datetime.date(2025, 3, 4)),
        ("19-Jul-2025", datetime.date(2025, 7, 19)),
        (None, None),
        ("", None),
        ("not a date", None),
    ],
)
def test_parse_date(value, expected):
    assert parse_date(value) == expected


def test_norm_strips_and_lowercases():
    assert norm("  USD -> INR (TT Buy SBI) ") == "usd->inr(ttbuysbi)"
    assert norm("Lot number") == "lotnumber"
    assert norm(None) == ""
