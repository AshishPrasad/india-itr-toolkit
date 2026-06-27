"""Low-level value parsing helpers (no external dependencies)."""
import datetime

# Characters stripped from numeric/currency text such as "₹1,234.50" or "12%".
_STRIP_CHARS = ("\u20b9", "$", "\u20ac", "\u00a3", ",", "%")
_EMPTY_TOKENS = ("", "-", "\u2014", "\u2013", "n/a", "na", "none")

_DATE_FORMATS = (
    "%Y-%m-%d",
    "%d-%m-%Y",
    "%d/%m/%Y",
    "%m/%d/%Y",
    "%Y/%m/%d",
    "%d-%b-%Y",
    "%d-%B-%Y",
    "%b %d, %Y",
    "%d %b %Y",
    "%d %B %Y",
)


def parse_amount(value):
    """Parse a number that may be plain or a currency/percent string.

    Returns a float, or None if the value is empty/non-numeric.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    for ch in _STRIP_CHARS:
        text = text.replace(ch, "")
    text = text.strip()
    if text.lower() in _EMPTY_TOKENS:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def parse_date(value):
    """Parse a date from an Excel datetime or a string in common formats.

    Returns a ``datetime.date`` or None.
    """
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    text = str(value).strip()
    if not text:
        return None
    for fmt in _DATE_FORMATS:
        try:
            return datetime.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def norm(text):
    """Normalise a header label: lowercase with all whitespace removed."""
    if text is None:
        return ""
    return "".join(str(text).lower().split())
