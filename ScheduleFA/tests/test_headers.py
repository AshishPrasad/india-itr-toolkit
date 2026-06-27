"""Tests for header detection and column resolution."""
import openpyxl

from dividend_contribution.headers import (
    find_header_row,
    is_blank_row,
    is_total_row,
    resolve_column,
)


def _ws(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row, start=1):
            if val is not None:
                ws.cell(r, c, val)
    return ws


DIV_SPECS = {
    "date": (None, "Date", ["date"]),
    "usd": (None, "Total Dividend (USD)", ["dividend (usd)", "usd amount"]),
    "rate": (None, "USD -> INR (TT Buy SBI)", ["usd -> inr", "rate"]),
}


def test_find_header_row_skips_title_rows():
    ws = _ws([
        ["Some Title", None, None],
        [None, None, None],
        ["Date", "Total Dividend (USD)", "USD -> INR (TT Buy SBI)"],
        ["2020-01-15", 100, "50.00"],
    ])
    hdr, mapping = find_header_row(ws, DIV_SPECS)
    assert hdr == 3
    assert mapping == {"date": 1, "usd": 2, "rate": 3}


def test_resolve_prefers_exact_over_alias():
    # "USD" appears in both the amount and the rate header; exact match wins.
    header_map = {
        "totaldividend(usd)": 2,
        "usd->inr(ttbuysbi)": 3,
    }
    assert resolve_column(header_map, DIV_SPECS["usd"]) == 2
    assert resolve_column(header_map, DIV_SPECS["rate"]) == 3


def test_resolve_specific_alias_beats_broad_on_earlier_column():
    # "Trade Date" (col 1) contains the broad "date" token, but the specific
    # "acquisition date" alias must win and resolve to "Acquisition Date" (col 2).
    header_map = {"tradedate": 1, "acquisitiondate": 2}
    spec = (None, "Date", ["acquisition date", "date"])
    assert resolve_column(header_map, spec) == 2


def test_resolve_broad_alias_still_matches_when_only_option():
    # With no more-specific column, the broad token is an acceptable fallback.
    header_map = {"tradedate": 1}
    spec = (None, "Date", ["acquisition date", "date"])
    assert resolve_column(header_map, spec) == 1


def test_explicit_override_takes_priority():
    header_map = {"mydate": 5, "date": 1}
    spec = ("My Date", "Date", ["date"])
    assert resolve_column(header_map, spec) == 5


def test_find_header_row_reports_partial_when_missing():
    ws = _ws([["Date", "Quantity"]])
    hdr, mapping = find_header_row(ws, DIV_SPECS)
    assert hdr is None
    assert mapping == {"date": 1}


def test_is_blank_and_total_rows():
    ws = _ws([
        ["Date", "Amount"],
        [None, None],
        ["Total", 999],
    ])
    assert is_blank_row(ws, 2)
    assert not is_blank_row(ws, 1)
    assert is_total_row(ws, 3, 1)
    assert not is_total_row(ws, 1, 1)
