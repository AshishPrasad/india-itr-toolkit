"""Tests for the core allocation logic."""
import datetime

from dividend_contribution.allocation import (
    allocate,
    held_qty_at,
    net_held_all,
    round_to_total,
    validate,
)

D = datetime.date

# Buy lot 1 (10) on Jan 1, buy lot 2 (10) on May 1, sell 5 of lot 1 on Aug 1.
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


def test_held_qty_at_respects_dates():
    assert held_qty_at(TXNS, "1", D(2025, 2, 15)) == 10.0   # before the sell
    assert held_qty_at(TXNS, "2", D(2025, 2, 15)) == 0.0    # not yet bought
    assert held_qty_at(TXNS, "1", D(2025, 8, 20)) == 5.0    # after the sell
    assert held_qty_at(TXNS, "2", D(2025, 8, 20)) == 10.0


def test_net_held_all_final_positions():
    assert net_held_all(TXNS, LOT_ORDER) == {"1": 5.0, "2": 10.0}


def test_allocate_full_example():
    contrib, unallocated, breakdown = allocate(DIVIDENDS, TXNS, LOT_ORDER)
    assert unallocated == 0.0
    assert contrib["1"] == 8000 + 4250 + 3000 + 9500 * 5 / 15
    assert contrib["2"] == 0 + 4250 + 6000 + 9500 * 10 / 15
    assert round(contrib["1"] + contrib["2"], 6) == 35000.0
    # breakdown: per-date total held quantity
    assert [b[2] for b in breakdown] == [10.0, 20.0, 15.0, 15.0]


def test_allocate_unallocated_when_no_shares_held():
    # Dividend before any purchase cannot be allocated.
    divs = [(D(2024, 1, 1), 500.0, 5.0, 100.0)] + DIVIDENDS
    contrib, unallocated, _ = allocate(divs, TXNS, LOT_ORDER)
    assert unallocated == 500.0
    assert round(sum(contrib.values()), 6) == 35000.0


def test_validate_flags_oversell():
    bad = [
        (D(2025, 1, 1), "buy", "9", 3.0),
        (D(2025, 2, 1), "sell", "9", 5.0),
    ]
    warnings = validate(bad, ["9"])
    assert len(warnings) == 1
    assert "negative" in warnings[0].lower()


def test_validate_clean_ledger_has_no_warnings():
    assert validate(TXNS, LOT_ORDER) == []


def test_round_to_total_foots_exactly():
    values = {"1": 18416.6667, "2": 16583.3333}
    rounded = round_to_total(values, 35000.00)
    assert round(sum(rounded.values()), 2) == 35000.00
    assert rounded["1"] == 18416.67
    assert rounded["2"] == 16583.33


def test_round_to_total_handles_negative_deficit():
    # Values sum to 2.01 but the target is 1.99, so the floored paise (200)
    # exceed the target (199): the negative-deficit branch must remove a paise
    # from the lot with the SMALLEST remainder ("a", .2 < "b", .8).
    values = {"a": 1.002, "b": 1.008}
    rounded = round_to_total(values, 1.99)
    assert round(sum(rounded.values()), 2) == 1.99
    assert rounded["a"] == 0.99
    assert rounded["b"] == 1.00
