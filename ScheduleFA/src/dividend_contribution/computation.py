"""High-level orchestration: turn parsed inputs into a final result."""
from dataclasses import dataclass, field
from typing import Dict, List, Tuple

from .allocation import allocate, net_held_all, round_to_total, validate


@dataclass
class Result:
    """Outcome of a dividend-contribution computation."""

    lot_order: List[str]
    contrib: Dict[str, float]            # unrounded per-lot INR
    contrib_rounded: Dict[str, float]    # 2dp per-lot INR, sums to ``allocated``
    net_held: Dict[str, float]           # final net held quantity per lot
    total_dividend: float
    allocated: float
    unallocated: float
    warnings: List[str] = field(default_factory=list)
    breakdown: List[Tuple] = field(default_factory=list)
    n_dividends: int = 0
    n_transactions: int = 0


def compute(dividends, txns, lot_order):
    """Run validation + allocation + rounding and return a :class:`Result`."""
    warnings = validate(txns, lot_order)
    contrib, unallocated, breakdown = allocate(dividends, txns, lot_order)
    total_dividend = sum(rec[1] for rec in dividends)
    allocated = total_dividend - unallocated
    contrib_rounded = round_to_total(contrib, round(allocated, 2))
    nh = net_held_all(txns, lot_order)
    return Result(
        lot_order=lot_order,
        contrib=contrib,
        contrib_rounded=contrib_rounded,
        net_held=nh,
        total_dividend=total_dividend,
        allocated=allocated,
        unallocated=unallocated,
        warnings=warnings,
        breakdown=breakdown,
        n_dividends=len(dividends),
        n_transactions=len(txns),
    )
