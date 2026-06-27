"""Generate the Schedule FA "Table A3" CSV from a per-lot Excel summary.

Public API:
    read_holdings  -- load the input Excel into Holding records
    write_csv      -- write the 12-column Schedule FA CSV
    Holding        -- one foreign-equity holding (lot)
    HEADERS        -- the exact output column headers, in order
"""
from .reader import Holding, ReadError, read_holdings
from .writer import HEADERS, write_csv

__version__ = "1.0.0"

__all__ = [
    "Holding",
    "ReadError",
    "read_holdings",
    "HEADERS",
    "write_csv",
    "__version__",
]
