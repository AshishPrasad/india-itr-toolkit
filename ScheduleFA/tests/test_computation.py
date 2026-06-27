"""Tests for the orchestration layer (computation.compute -> Result)."""
import datetime

from dividend_contribution.computation import Result, compute

D = datetime.date

# Lot 1: 10 shares all year; Lot 2: 10 shares bought before the 2nd dividend;
# 5 of lot 1 sold before the 3rd dividend.
TXNS = [
    (D(2025, 1, 1), "buy", "1", 10.0),
    (D(2025, 5, 1), "buy", "2", 10.0),
    (D(2025, 8, 1), "sell", "1", 5.0),
]
LOT_ORDER = ["1", "2"]
DIVIDENDS = [
    (D(2025, 2, 15), 8000.0, 100.0, 80.0),
    (D(2025, 5, 20), 8500.0, 100.0, 85.0),
    (D(2025, 8, 20), 9000.0, 100.0, 90.0),
    (D(2025, 11, 20), 9500.0, 100.0, 95.0),
]


def test_compute_returns_result_with_counts():
    result = compute(DIVIDENDS, TXNS, LOT_ORDER)
    assert isinstance(result, Result)
    assert result.n_dividends == 4
    assert result.n_transactions == 3
    assert result.lot_order == LOT_ORDER
    assert result.net_held == {"1": 5.0, "2": 10.0}


def test_compute_reconciles_total_allocated_unallocated():
    result = compute(DIVIDENDS, TXNS, LOT_ORDER)
    assert result.total_dividend == 35000.0
    assert result.unallocated == 0.0
    assert round(result.allocated + result.unallocated, 2) == round(
        result.total_dividend, 2
    )


def test_compute_contrib_rounded_foots_to_allocated():
    result = compute(DIVIDENDS, TXNS, LOT_ORDER)
    assert round(sum(result.contrib_rounded.values()), 2) == round(
        result.allocated, 2
    )
    # Every lot present in the rounded mapping.
    assert set(result.contrib_rounded) == set(LOT_ORDER)


def test_compute_unallocated_when_dividend_predates_holdings():
    # A dividend before any purchase cannot be allocated to any lot.
    divs = [(D(2024, 1, 1), 500.0, 5.0, 100.0)] + DIVIDENDS
    result = compute(divs, TXNS, LOT_ORDER)
    assert result.unallocated == 500.0
    assert result.allocated == 35000.0
    assert round(result.allocated + result.unallocated, 2) == round(
        result.total_dividend, 2
    )
    # The rounded per-lot contributions foot to the allocated amount, not total.
    assert round(sum(result.contrib_rounded.values()), 2) == 35000.0


def test_compute_surfaces_oversell_warnings():
    bad = [
        (D(2025, 1, 1), "buy", "9", 3.0),
        (D(2025, 2, 1), "sell", "9", 5.0),
    ]
    result = compute(DIVIDENDS, bad, ["9"])
    assert result.warnings
    assert any("negative" in w.lower() for w in result.warnings)


def test_compute_clean_ledger_has_no_warnings():
    assert compute(DIVIDENDS, TXNS, LOT_ORDER).warnings == []
