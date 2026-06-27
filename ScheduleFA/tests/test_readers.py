"""Tests for reading the sample Excel workbooks."""
import datetime

import pytest

from dividend_contribution.readers import (
    ReadError,
    read_dividends,
    read_transactions,
)


def test_read_dividends_computes_inr(dividends_path):
    rows = read_dividends(dividends_path)
    assert len(rows) == 4  # the "Total" row is excluded
    first = rows[0]
    assert first[0] == datetime.date(2025, 2, 15)
    assert first[2] == 100.0           # USD
    assert first[3] == 80.0            # rate (parsed from "₹80.00")
    assert first[1] == 8000.0          # INR = USD * rate
    assert round(sum(r[1] for r in rows), 2) == 35000.0


def test_read_transactions_parses_ledger(stocks_path):
    txns, lot_order = read_transactions(stocks_path)
    assert lot_order == ["1", "2"]
    assert txns[0] == (datetime.date(2025, 1, 1), "buy", "1", 10.0)
    assert txns[2] == (datetime.date(2025, 8, 1), "sell", "1", 5.0)


def test_read_transactions_alias_headers(stocks_renamed_path):
    # Headers are "Txn Date / Action / Lot / Shares" -> resolved via aliases.
    txns, lot_order = read_transactions(stocks_renamed_path)
    assert lot_order == ["1", "2"]
    assert len(txns) == 3


def test_read_dividends_missing_file_raises(tmp_path):
    with pytest.raises(ReadError):
        read_dividends(str(tmp_path / "does_not_exist.xlsx"))


def test_read_transactions_bad_sheet_raises(stocks_path):
    with pytest.raises(ReadError):
        read_transactions(stocks_path, sheet="NoSuchSheet")
