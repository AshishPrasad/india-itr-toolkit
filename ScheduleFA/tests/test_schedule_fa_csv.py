"""Tests for the schedule_fa_csv package (reader, writer, CLI)."""
import csv
import datetime
import re

import openpyxl
import pytest

from schedule_fa_csv.reader import Holding, ReadError, extract_holdings, read_holdings
from schedule_fa_csv.writer import HEADERS, _sanitize, _round_series, write_csv
from schedule_fa_csv.cli import main


def _ws(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row, start=1):
            if val is not None:
                ws.cell(r, c, val)
    return ws


FULL_HEADERS = [
    "Date", "Lot number", "Quantity", "Initial value", "Peak value in INR",
    "Closing value in INR", "Dividend Contribution", "Sale proceeds",
    "Country/Region name", "Country Name and Code", "Name of entity",
    "Address of entity", "ZIP Code", "Nature of entity",
]
ROW_1 = [
    datetime.datetime(2023, 5, 10), 1, 10, 100000, 150000, 130000, 5000, 0,
    "United States of America", "United States of America - 2", "Example Corp",
    "1 Example Way", "00000", "Listed company",
]


# --------------------------------------------------------------------------- #
# Reader
# --------------------------------------------------------------------------- #
def test_read_holdings_sample(fa_input_path):
    holdings = read_holdings(fa_input_path)
    assert len(holdings) == 2
    h1, h2 = holdings
    assert h1.date_acquired == datetime.date(2023, 5, 10)
    assert h1.initial_value == 100000.0
    assert h1.peak_value == 150000.0
    assert h1.closing_value == 130000.0
    assert h1.gross_paid == 5000.0
    assert h1.gross_proceeds == 0.0          # not sold
    assert h1.zip_code == "00000"            # leading zeros preserved
    assert h1.entity_name == "Example Corp"
    assert h2.gross_proceeds == 75000.0      # sold lot


def test_read_holdings_proceeds_optional_defaults_zero():
    # No "Sale proceeds" column at all -> proceeds default to 0.
    rows = [
        [h for h in FULL_HEADERS if h != "Sale proceeds"],
        [v for h, v in zip(FULL_HEADERS, ROW_1) if h != "Sale proceeds"],
    ]
    holdings = extract_holdings(_ws(rows))
    assert len(holdings) == 1
    assert holdings[0].gross_proceeds == 0.0


def test_read_holdings_alias_headers():
    headers = [
        "Date of acquiring the interest", "Lot", "Qty",
        "Initial value of the investment", "Peak value of investment",
        "Closing balance", "Total gross amount paid", "Sale proceeds (INR)",
        "Country/Region", "Country code", "Entity name", "Address",
        "PIN code", "Nature",
    ]
    row = list(ROW_1)
    row[7] = 12345  # non-zero so the proceeds alias is actually exercised
    holdings = extract_holdings(_ws([headers, row]))
    assert len(holdings) == 1
    assert holdings[0].country_region == "United States of America"
    # Proves the "Sale proceeds (INR)" alias resolved (default would be 0.0).
    assert holdings[0].gross_proceeds == 12345.0


def test_read_holdings_specific_alias_wins_over_broad():
    # "Trade Date"/"Transaction cost" precede the real columns; broad aliases
    # ("date", "cost") must not hijack the acquisition date / initial value.
    headers = [
        "Trade Date", "Acquisition Date", "Lot number", "Transaction cost",
        "Cost basis", "Peak value in INR", "Closing value in INR",
        "Dividend Contribution", "Country/Region name", "Country Name and Code",
        "Name of entity", "Address of entity", "ZIP Code", "Nature of entity",
    ]
    row = [
        datetime.datetime(2020, 1, 1), datetime.datetime(2023, 5, 10), "L1",
        49.99, 100000, 150000, 130000, 5000, "USA", "USA - 2", "E", "A",
        "00000", "Listed",
    ]
    h = extract_holdings(_ws([headers, row]))[0]
    assert h.date_acquired == datetime.date(2023, 5, 10)  # not the Trade Date
    assert h.initial_value == 100000.0                    # not the 49.99 fee


def test_read_holdings_ambiguous_mapping_raises():
    # One column would satisfy two logical fields -> collision error.
    headers = [
        "Date", "Lot number", "Value", "Peak value in INR",
        "Closing value in INR", "Dividend Contribution",
        "Country/Region name", "Country Name and Code", "Name of entity",
        "Address of entity", "ZIP Code", "Nature of entity",
    ]
    row = [
        ROW_1[0], "L1", 100000, 150000, 130000, 5000,
        "USA", "USA - 2", "E", "A", "00000", "Listed",
    ]
    # Both initial_value and (overridden) closing map to the single "Value" col.
    with pytest.raises(ReadError):
        extract_holdings(
            _ws([headers, row]),
            cols={"initial_value": "Value", "closing_value": "Value"},
        )


def test_read_holdings_numeric_zip_has_no_decimal():
    row = list(ROW_1)
    row[12] = 98052  # numeric ZIP
    holdings = extract_holdings(_ws([FULL_HEADERS, row]))
    assert holdings[0].zip_code == "98052"


def test_read_holdings_stops_at_total_row():
    total = [None] * len(FULL_HEADERS)
    total[0] = "Total"
    holdings = extract_holdings(_ws([FULL_HEADERS, ROW_1, total, ROW_1]))
    assert len(holdings) == 1


def test_read_holdings_missing_headers_raises():
    with pytest.raises(ReadError):
        extract_holdings(_ws([["Date", "Lot number"], [ROW_1[0], 1]]))


def test_read_holdings_no_data_rows_raises():
    # Headers present but no data rows -> ReadError.
    with pytest.raises(ReadError):
        extract_holdings(_ws([FULL_HEADERS]))


def test_read_holdings_skips_rows_without_a_date():
    no_date = list(ROW_1)
    no_date[0] = None  # missing acquisition date -> row skipped
    holdings = extract_holdings(_ws([FULL_HEADERS, no_date, ROW_1]))
    assert len(holdings) == 1
    assert holdings[0].date_acquired == datetime.date(2023, 5, 10)


def test_read_holdings_currency_text_parsed():
    row = list(ROW_1)
    row[3] = "\u20b91,00,000.00"  # initial value as lakh-grouped rupee text
    holdings = extract_holdings(_ws([FULL_HEADERS, row]))
    assert holdings[0].initial_value == 100000.0


# --------------------------------------------------------------------------- #
# Writer
# --------------------------------------------------------------------------- #
def test_sanitize_blocks_formula_text():
    assert _sanitize("=cmd()") == "'=cmd()"
    assert _sanitize("@x") == "'@x"
    assert _sanitize("Example Corp") == "Example Corp"


def test_round_series_foots_to_rounded_total():
    # 33.34 + 33.33 + 33.33 = 100.00 -> rounds to 100 (the lost rupee restored).
    result = _round_series([33.34, 33.33, 33.33])
    assert sum(result) == 100
    assert all(isinstance(x, int) for x in result)


def test_round_series_distributes_to_largest_remainders():
    # 10.4 + 10.4 + 10.4 = 31.2 -> target 31; one rupee added back.
    result = _round_series([10.4, 10.4, 10.4])
    assert sum(result) == 31
    assert sorted(result) == [10, 10, 11]


def test_round_series_handles_empty_and_whole_numbers():
    assert _round_series([]) == []
    assert _round_series([100000.0, 130000.0]) == [100000, 130000]


def test_write_csv_headers_and_values(fa_input_path, tmp_path):
    holdings = read_holdings(fa_input_path)
    out = tmp_path / "fa.csv"
    write_csv(str(out), holdings)

    with open(str(out), newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))
    assert rows[0] == HEADERS                 # exact header order
    assert rows[1][6] == "2023-05-10"         # default YYYY-MM-DD date
    assert rows[1][7] == "100000"             # initial value (whole rupees)
    assert rows[1][9] == "130000"             # closing balance
    assert rows[1][11] == "0"                 # proceeds (not sold)
    assert rows[2][11] == "75000"             # proceeds (sold)


def test_write_csv_custom_date_format(fa_input_path, tmp_path):
    holdings = read_holdings(fa_input_path)
    out = tmp_path / "fa.csv"
    write_csv(str(out), holdings, date_format="%d-%m-%Y")
    with open(str(out), newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))
    assert rows[1][6] == "10-05-2023"


def test_write_csv_dividend_column_foots_to_total(tmp_path):
    # Three lots whose contributions sum to 100.00 (33.34 + 33.33 + 33.33).
    def _holding(paid):
        return Holding(
            lot="x", country_region="C", country_code="C", entity_name="E",
            address="A", zip_code="0", nature="N",
            date_acquired=datetime.date(2025, 1, 1),
            initial_value=0.0, peak_value=0.0, closing_value=0.0,
            gross_paid=paid, gross_proceeds=0.0,
        )
    holdings = [_holding(33.34), _holding(33.33), _holding(33.33)]
    out = tmp_path / "fa.csv"
    write_csv(str(out), holdings)
    with open(str(out), newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))[1:]
    assert sum(int(r[10]) for r in rows) == 100         # column foots exactly
    assert all(re.fullmatch(r"-?\d+", r[10]) for r in rows)  # whole rupees


def test_write_csv_neutralises_formula_entity(tmp_path):
    holdings = extract_holdings(_ws([
        FULL_HEADERS,
        [*ROW_1[:10], "=HYPERLINK(1)", *ROW_1[11:]],
    ]))
    out = tmp_path / "fa.csv"
    write_csv(str(out), holdings)
    with open(str(out), newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))
    assert rows[1][2] == "'=HYPERLINK(1)"


# --------------------------------------------------------------------------- #
# CLI end-to-end
# --------------------------------------------------------------------------- #
def test_cli_end_to_end(fa_input_path, tmp_path, capsys):
    out = tmp_path / "schedule_fa.csv"
    rc = main(["--input", fa_input_path, "--output", str(out)])
    assert rc == 0
    assert out.exists()
    with open(str(out), newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))
    assert rows[0] == HEADERS
    assert len(rows) == 3                      # header + 2 holdings
    assert "Holdings read" in capsys.readouterr().out


def test_cli_missing_input_returns_error(tmp_path, capsys):
    out = tmp_path / "schedule_fa.csv"
    rc = main(["--input", str(tmp_path / "nope.xlsx"), "--output", str(out)])
    assert rc == 2
    assert "ERROR" in capsys.readouterr().err


def test_cli_column_override_maps_custom_header(tmp_path):
    # Rename the peak column to something the aliases won't match.
    headers = list(FULL_HEADERS)
    headers[FULL_HEADERS.index("Peak value in INR")] = "Maximum INR"
    inp = tmp_path / "in.xlsx"
    _ws([headers, ROW_1]).parent.save(str(inp))
    out = tmp_path / "out.csv"

    # Without the override, the peak column cannot be located -> error.
    assert main(["--input", str(inp), "--output", str(out)]) == 2

    # With --peak-col the custom header is resolved and mapped through.
    rc = main([
        "--input", str(inp), "--output", str(out), "--peak-col", "Maximum INR",
    ])
    assert rc == 0
    with open(str(out), newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))
    assert rows[1][8] == "150000"  # peak value carried into the output

