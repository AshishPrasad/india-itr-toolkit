"""Command-line interface for the Schedule FA CSV generator."""
import argparse
import sys

from . import __version__
from .reader import ReadError, read_holdings
from .writer import write_csv


def build_parser():
    parser = argparse.ArgumentParser(
        prog="schedule-fa-csv",
        description=(
            "Generate the Schedule FA Table A3 (financial interest in an "
            "entity) CSV from a per-lot Excel summary."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--version", action="version", version="%(prog)s " + __version__
    )
    parser.add_argument("--input", required=True, help="Path to the input .xlsx")
    parser.add_argument("--output", required=True, help="Path for the output .csv")
    parser.add_argument("--sheet", help="Sheet name in the input file (default: first)")
    parser.add_argument(
        "--date-format",
        default="%d-%m-%y",
        help="strftime format for the acquisition date column",
    )

    cols = parser.add_argument_group("column overrides (exact header names)")
    cols.add_argument("--date-col", help="Acquisition date column")
    cols.add_argument("--lot-col", help="Lot number column")
    cols.add_argument("--initial-col", help="Initial value column")
    cols.add_argument("--peak-col", help="Peak value (INR) column")
    cols.add_argument("--closing-col", help="Closing value (INR) column")
    cols.add_argument("--dividend-col", help="Dividend contribution column")
    cols.add_argument("--proceeds-col", help="Sale proceeds column")
    cols.add_argument("--country-region-col", help="Country/Region name column")
    cols.add_argument("--country-code-col", help="Country Name and Code column")
    cols.add_argument("--entity-name-col", help="Name of entity column")
    cols.add_argument("--address-col", help="Address of entity column")
    cols.add_argument("--zip-col", help="ZIP Code column")
    cols.add_argument("--nature-col", help="Nature of entity column")
    return parser


def _columns(args):
    return {
        "date": args.date_col,
        "lot": args.lot_col,
        "initial_value": args.initial_col,
        "peak_value": args.peak_col,
        "closing_value": args.closing_col,
        "gross_paid": args.dividend_col,
        "gross_proceeds": args.proceeds_col,
        "country_region": args.country_region_col,
        "country_code": args.country_code_col,
        "entity_name": args.entity_name_col,
        "address": args.address_col,
        "zip_code": args.zip_col,
        "nature": args.nature_col,
    }


def main(argv=None):
    args = build_parser().parse_args(argv)
    try:
        holdings = read_holdings(args.input, args.sheet, _columns(args))
    except ReadError as exc:
        print("ERROR: %s" % exc, file=sys.stderr)
        return 2

    write_csv(args.output, holdings, date_format=args.date_format)

    total_paid = sum(h.gross_paid for h in holdings)
    total_proceeds = sum(h.gross_proceeds for h in holdings)
    print("Holdings read        : %d" % len(holdings))
    print("Total gross paid     : \u20b9{:,.2f}".format(total_paid))
    print("Total gross proceeds : \u20b9{:,.2f}".format(total_proceeds))
    print("Output written to    : %s" % args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
