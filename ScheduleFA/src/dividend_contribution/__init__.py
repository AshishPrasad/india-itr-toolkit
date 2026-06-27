"""Compute per-lot dividend contribution from a dividend file and a stock ledger.

Public API:
    read_dividends, read_transactions  -- load input Excel files
    allocate, compute                  -- core allocation logic
    write_output                       -- write the result Excel file
"""
from .parsing import parse_amount, parse_date, norm
from .readers import read_dividends, read_transactions
from .allocation import (
    held_qty_at,
    net_held_all,
    validate,
    allocate,
    round_to_total,
)
from .computation import compute, Result
from .writer import write_output

__version__ = "1.0.0"

__all__ = [
    "parse_amount",
    "parse_date",
    "norm",
    "read_dividends",
    "read_transactions",
    "held_qty_at",
    "net_held_all",
    "validate",
    "allocate",
    "round_to_total",
    "compute",
    "Result",
    "write_output",
    "__version__",
]
