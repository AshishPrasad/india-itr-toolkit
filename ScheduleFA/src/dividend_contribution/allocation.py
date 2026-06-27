"""Core allocation logic (pure Python, no Excel dependencies)."""
import datetime

_EPS = 1e-9


def held_qty_at(txns, lot, on_date):
    """Net held quantity of ``lot`` on/before ``on_date``.

    Buys add, sells subtract, for transactions dated on or before ``on_date``.
    """
    q = 0.0
    for d, ttype, l, qty in txns:
        if l != lot or d > on_date:
            continue
        q += qty if ttype == "buy" else -qty
    return q


def net_held_all(txns, lot_order):
    """Final net held quantity for every lot (as of the latest possible date)."""
    return {lot: held_qty_at(txns, lot, datetime.date.max) for lot in lot_order}


def validate(txns, lot_order):
    """Return human-readable warnings for lots that are over-sold."""
    warnings = []
    for lot in lot_order:
        events = sorted(
            (d, qty if t == "buy" else -qty)
            for d, t, l, qty in txns
            if l == lot
        )
        running = 0.0
        for d, delta in events:
            running += delta
            if running < -_EPS:
                warnings.append(
                    "Lot %s: net holding went negative (%.4f) on %s "
                    "(sold more than held)." % (lot, running, d)
                )
                break
    return warnings


def allocate(dividends, txns, lot_order):
    """Allocate each dividend across lots by net held quantity on its date.

    Returns ``(contrib, unallocated, breakdown)`` where:
        contrib     -- {lot: total INR contribution}
        unallocated -- INR that could not be allocated (no shares held)
        breakdown   -- [(date, inr, total_held_qty, n_lots), ...]
    """
    contrib = {lot: 0.0 for lot in lot_order}
    unallocated = 0.0
    breakdown = []
    for record in sorted(dividends, key=lambda x: x[0]):
        d, inr = record[0], record[1]
        held = {}
        for lot in lot_order:
            q = held_qty_at(txns, lot, d)
            if q > _EPS:
                held[lot] = q
        total_held = sum(held.values())
        if total_held <= 0:
            unallocated += inr
            breakdown.append((d, inr, 0.0, 0))
            continue
        for lot, q in held.items():
            contrib[lot] += inr * q / total_held
        breakdown.append((d, inr, total_held, len(held)))
    return contrib, unallocated, breakdown


def round_to_total(values, target):
    """Round values to 2 dp using largest-remainder so they sum to ``target``."""
    target_paise = int(round(target * 100))
    floors, remainders = {}, {}
    for key, val in values.items():
        paise = val * 100
        floor = int(paise // 1)
        floors[key] = floor
        remainders[key] = paise - floor
    deficit = target_paise - sum(floors.values())
    if deficit > 0:
        order = sorted(remainders, key=lambda k: remainders[k], reverse=True)
        for key in order[:deficit]:
            floors[key] += 1
    elif deficit < 0:
        order = sorted(remainders, key=lambda k: remainders[k])
        for key in order[: -deficit]:
            floors[key] -= 1
    return {key: floors[key] / 100.0 for key in values}
