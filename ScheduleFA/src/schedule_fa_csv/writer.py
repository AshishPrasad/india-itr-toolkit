"""Write Schedule FA holdings to the Table A3 CSV format."""
import csv
from decimal import ROUND_FLOOR, ROUND_HALF_UP, Decimal

# Exact output column headers, in order (Schedule FA "Table A3").
HEADERS = [
    "Country/Region name",
    "Country Name and Code",
    "Name of entity",
    "Address of entity",
    "ZIP Code",
    "Nature of entity",
    "Date of acquiring the interest",
    "Initial value of the investment",
    "Peak value of investment during the Period",
    "Closing balance",
    "Total gross amount paid/credited with respect to the holding during the period",
    "Total gross proceeds from sale or redemption of investment during the period",
]

# Leading characters a spreadsheet may interpret as the start of a formula.
_FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")

_DEFAULT_DATE_FORMAT = "%d-%m-%y"


def _sanitize(value):
    """Neutralise CSV/spreadsheet formula injection in text fields."""
    if isinstance(value, str) and value.startswith(_FORMULA_PREFIXES):
        return "'" + value
    return value


def _round_series(values):
    """Round floats to whole rupees so they sum to the rounded grand total.

    Uses largest-remainder rounding: each value is floored, then the rupees lost
    to flooring are added back to the rows with the largest fractional parts, so
    ``sum(result)`` equals the column total rounded half-up. Returns a list of
    ints aligned with ``values``.
    """
    if not values:
        return []
    decs = [Decimal(str(v)) for v in values]
    floors = [int(d.to_integral_value(rounding=ROUND_FLOOR)) for d in decs]
    remainders = [decs[i] - Decimal(floors[i]) for i in range(len(decs))]
    target = int(sum(decs).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    result = list(floors)
    deficit = target - sum(floors)  # always >= 0 for floor-based remainders
    if deficit > 0:
        order = sorted(range(len(decs)), key=lambda i: remainders[i], reverse=True)
        for i in order[:deficit]:
            result[i] += 1
    return result


def write_csv(path, holdings, date_format=_DEFAULT_DATE_FORMAT):
    """Write ``holdings`` to ``path`` as the Schedule FA Table A3 CSV.

    Amounts are written as whole rupees using largest-remainder rounding, so each
    numeric column sums exactly to its rounded total (no per-row drift). Values
    carry no currency symbol or thousands separators, so the file is safe to
    parse as CSV. The file is UTF-8 with a BOM for clean opening in Excel.
    Returns ``path``.
    """
    initial = _round_series([h.initial_value for h in holdings])
    peak = _round_series([h.peak_value for h in holdings])
    closing = _round_series([h.closing_value for h in holdings])
    paid = _round_series([h.gross_paid for h in holdings])
    proceeds = _round_series([h.gross_proceeds for h in holdings])

    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow(HEADERS)
        for i, holding in enumerate(holdings):
            acquired = (
                holding.date_acquired.strftime(date_format)
                if holding.date_acquired
                else ""
            )
            writer.writerow([
                _sanitize(holding.country_region),
                _sanitize(holding.country_code),
                _sanitize(holding.entity_name),
                _sanitize(holding.address),
                _sanitize(holding.zip_code),
                _sanitize(holding.nature),
                acquired,
                str(initial[i]),
                str(peak[i]),
                str(closing[i]),
                str(paid[i]),
                str(proceeds[i]),
            ])
    return path
