"""Command-line interface for the dividend-contribution tool."""
import argparse
import sys

from . import __version__
from .readers import ReadError, read_dividends, read_transactions
from .computation import compute
from .writer import write_output


def build_parser():
    parser = argparse.ArgumentParser(
        prog="dividend-contribution",
        description=(
            "Compute how much of the total dividend each stock lot contributed, "
            "based on the quantity of each lot held on every dividend payout date."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--version", action="version", version="%(prog)s " + __version__)
    parser.add_argument("--dividends", required=True, help="Path to the dividends .xlsx")
    parser.add_argument("--stocks", required=True, help="Path to the stock-ledger .xlsx")
    parser.add_argument("--output", required=True, help="Path for the output .xlsx")
    parser.add_argument("--div-sheet", help="Sheet name in the dividends file (default: first)")
    parser.add_argument("--stk-sheet", help="Sheet name in the stocks file (default: first)")

    div = parser.add_argument_group("dividend column overrides")
    div.add_argument("--div-date-col", help="Dividend date column header")
    div.add_argument("--div-usd-col", help="Dividend USD-amount column header")
    div.add_argument("--div-rate-col", help="USD->INR rate column header")

    stk = parser.add_argument_group("stock column overrides")
    stk.add_argument("--stk-date-col", help="Transaction date column header")
    stk.add_argument("--stk-type-col", help="Buy/Sell column header")
    stk.add_argument("--stk-lot-col", help="Lot number column header")
    stk.add_argument("--stk-qty-col", help="Quantity column header")
    return parser


def _print_summary(result, output_path):
    print("Dividend events read : %d" % result.n_dividends)
    print(
        "Stock transactions   : %d  across %d lots"
        % (result.n_transactions, len(result.lot_order))
    )
    print("Total dividend (INR) : \u20b9{:,.2f}".format(result.total_dividend))
    if result.unallocated > 1e-6:
        print(
            "  Unallocated (no shares held on a payout date): \u20b9{:,.2f}".format(
                result.unallocated
            )
        )
    print("Allocated to lots    : \u20b9{:,.2f}".format(result.allocated))
    print("Output written to    : %s" % output_path)
    if result.warnings:
        print("\nWARNINGS:")
        for warning in result.warnings:
            print("  - %s" % warning)


def main(argv=None):
    args = build_parser().parse_args(argv)

    div_cols = {
        "date": args.div_date_col,
        "usd": args.div_usd_col,
        "rate": args.div_rate_col,
    }
    stk_cols = {
        "date": args.stk_date_col,
        "type": args.stk_type_col,
        "lot": args.stk_lot_col,
        "qty": args.stk_qty_col,
    }

    try:
        dividends = read_dividends(args.dividends, args.div_sheet, div_cols)
        txns, lot_order = read_transactions(args.stocks, args.stk_sheet, stk_cols)
    except ReadError as exc:
        print("ERROR: %s" % exc, file=sys.stderr)
        return 2

    result = compute(dividends, txns, lot_order)
    write_output(args.output, result)
    _print_summary(result, args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
