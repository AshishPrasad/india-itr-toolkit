"""Locate header rows and resolve logical fields to worksheet columns.

These helpers operate on any object exposing the openpyxl worksheet interface
(``cell(row, col).value``, ``max_row``, ``max_column``), so they can be unit
tested with real in-memory workbooks.
"""
from .parsing import norm

# A ColumnSpec describes how to find one logical field in a header row:
#   explicit      -- exact header name supplied by the user (highest priority)
#   default_exact -- the default exact header name to try next
#   aliases       -- patterns matched against header text (specific ones win)
ColumnSpec = tuple  # (explicit, default_exact, [aliases])


def resolve_column(header_map, spec):
    """Resolve a single field to a column index using a header lookup map.

    ``header_map`` maps normalised header text -> column index. Resolution order:

    1. Explicit user override (exact match) — highest priority.
    2. Exact normalised equality against ``default_exact`` or any alias. Patterns
       are tried in declaration order (most specific first), so an exact
       ``"acquisition date"`` column wins over a broad ``"date"`` alias.
    3. Substring fallback for headers that carry extra qualifiers, trying the
       **longest (most specific) pattern first** so a broad single-word alias
       (e.g. ``"date"``, ``"cost"``, ``"lot"``) never pre-empts a more specific
       one on an earlier column.

    Returns the column index, or None if not found.
    """
    explicit, default_exact, aliases = spec
    if explicit:
        return header_map.get(norm(explicit))

    candidates = [norm(default_exact)] + [norm(a) for a in aliases]

    # Pass 1: exact normalised equality, specific patterns first.
    for pattern in candidates:
        if pattern in header_map:
            return header_map[pattern]

    # Pass 2: substring fallback, longest (most specific) pattern first.
    for pattern in sorted(candidates, key=len, reverse=True):
        for header_text, col in header_map.items():
            if pattern in header_text:
                return col
    return None


def _header_map(ws, row):
    mapping = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(row, c).value
        if val is not None:
            mapping[norm(val)] = c
    return mapping


def find_header_row(ws, specs, max_scan=50):
    """Find the first row that resolves every field in ``specs``.

    ``specs`` maps field name -> ColumnSpec.
    Returns ``(header_row_index, {field: column})`` when all fields resolve,
    otherwise ``(None, best_partial_mapping)`` for diagnostics.
    """
    best_row, best_map, best_hit = None, {}, -1
    scan_to = min(ws.max_row, max_scan)
    for r in range(1, scan_to + 1):
        header_map = _header_map(ws, r)
        mapping = {}
        for field, spec in specs.items():
            col = resolve_column(header_map, spec)
            if col is not None:
                mapping[field] = col
        if len(mapping) == len(specs):
            return r, mapping
        if len(mapping) > best_hit:
            best_hit, best_row, best_map = len(mapping), r, mapping
    return None, best_map


def is_blank_row(ws, row):
    """True if every cell in the row is empty."""
    return all(ws.cell(row, c).value is None for c in range(1, ws.max_column + 1))


def is_total_row(ws, row, key_col):
    """True if the key column looks like a totals/footer label."""
    val = ws.cell(row, key_col).value
    return isinstance(val, str) and any(
        k in val.lower() for k in ("total", "grand", "sum")
    )
