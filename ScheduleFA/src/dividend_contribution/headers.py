"""Locate header rows and resolve logical fields to worksheet columns.

These helpers operate on any object exposing the openpyxl worksheet interface
(``cell(row, col).value``, ``max_row``, ``max_column``), so they can be unit
tested with real in-memory workbooks.
"""
from .parsing import norm

# A ColumnSpec describes how to find one logical field in a header row:
#   explicit      -- exact header name supplied by the user (highest priority)
#   default_exact -- the default exact header name to try next
#   aliases       -- substrings to match as a last resort
ColumnSpec = tuple  # (explicit, default_exact, [aliases])


def resolve_column(header_map, spec):
    """Resolve a single field to a column index using a header lookup map.

    ``header_map`` maps normalised header text -> column index.
    Returns the column index, or None if not found.
    """
    explicit, default_exact, aliases = spec
    if explicit:
        return header_map.get(norm(explicit))
    if norm(default_exact) in header_map:
        return header_map[norm(default_exact)]
    for header_text, col in header_map.items():
        if any(norm(a) in header_text for a in aliases):
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
