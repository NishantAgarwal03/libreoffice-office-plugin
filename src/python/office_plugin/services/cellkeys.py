"""Single definition of "are these two cell values the same?" and "is this cell blank?".

Purpose: normalise Calc cell values (text or numbers from getDataArray) into hashable keys so
         Compare Lists, Highlight and other features match values the same way.
Inputs:  raw cell values (str / float / int / None) and a CleanOptions.
Outputs: hashable keys; booleans.
Must not: import uno (pure Python, unit-tested without LibreOffice).
"""

import re
import unicodedata
from collections import namedtuple

CleanOptions = namedtuple("CleanOptions", "trim ignore_case numeric strip_nonprint")
CleanOptions.__new__.__defaults__ = (True, True, True, True)

# no cleaning at all: exact comparison (a number never equals text)
EXACT = CleanOptions(False, False, False, False)

_WS = re.compile(r"\s+")                      # \s includes NBSP and other unicode spaces in str patterns
_NUMBER = re.compile(r"^[+-]?(\d+(\.\d*)?|\.\d+)$")


def _strip_nonprintable(text):
    # control (Cc) and format (Cf: zero-width space, BOM, soft hyphen...) characters, but keep whitespace
    return "".join(ch for ch in text
                   if ch.isspace() or unicodedata.category(ch) not in ("Cc", "Cf"))


def clean_text(text, opts):
    if opts.strip_nonprint:
        text = _strip_nonprintable(text)
    if opts.trim:
        text = _WS.sub(" ", text).strip()
    if opts.ignore_case:
        text = text.casefold()
    return text


def _canonical_number(value):
    value = float(value)
    return int(value) if value.is_integer() else value


def normalize(value, opts=CleanOptions()):
    """-> hashable key; equal keys mean 'same value' under `opts`."""
    if value is None:
        return ("s", "")
    if isinstance(value, bool):
        return ("n", int(value))
    if isinstance(value, (int, float)):
        return ("n", _canonical_number(value))
    text = clean_text(str(value), opts)
    if opts.numeric and _NUMBER.match(text):
        return ("n", _canonical_number(text))
    return ("s", text)


def row_key(row, cols, opts=CleanOptions()):
    """Key of the chosen columns of one row (missing trailing cells count as blank)."""
    return tuple(normalize(row[c] if c < len(row) else "", opts) for c in cols)


def is_blank(value, whitespace=False):
    """Empty cell. With `whitespace`, text made only of spaces / invisible characters is blank too."""
    if value is None or value == "":
        return True
    if whitespace and isinstance(value, str):
        return clean_text(value, CleanOptions(True, False, False, True)) == ""
    return False


def header_guess(first_row):
    """True when a row looks like a header: at least one text cell and no numbers."""
    cells = [v for v in first_row if v != ""]
    return bool(cells) and all(isinstance(v, str) for v in cells)


def key_is_blank(key):
    """True when every part of a row_key is blank text."""
    return all(part == ("s", "") for part in key)
