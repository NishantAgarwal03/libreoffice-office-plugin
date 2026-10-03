"""Numbers written with unit letters/words -> numbers (pure Python, no uno).

Purpose: read '1 lakh', '1L', '2 cr', '2cr', '1.5k', '3M', 'Rs. 2 crore', '12,34,567',
         '1 crore 25 lakh', '(2.5L)' and similar into exact Decimals.
Inputs:  text. Outputs: Decimal; ValueError with a readable message.
Must not: import uno.
"""

import re
from decimal import Decimal

_THOUSAND, _LAKH, _CRORE = Decimal(10) ** 3, Decimal(10) ** 5, Decimal(10) ** 7
_MILLION, _BILLION, _TRILLION = Decimal(10) ** 6, Decimal(10) ** 9, Decimal(10) ** 12

SUFFIXES = {
    "k": _THOUSAND, "thousand": _THOUSAND, "thou": _THOUSAND, "th": _THOUSAND, "grand": _THOUSAND,
    "l": _LAKH, "lk": _LAKH, "lac": _LAKH, "lacs": _LAKH, "lakh": _LAKH, "lakhs": _LAKH,
    "cr": _CRORE, "crs": _CRORE, "crore": _CRORE, "crores": _CRORE,
    "m": _MILLION, "mn": _MILLION, "mil": _MILLION, "million": _MILLION, "millions": _MILLION,
    "b": _BILLION, "bn": _BILLION, "bil": _BILLION, "billion": _BILLION, "billions": _BILLION,
    "arab": _BILLION, "arabs": _BILLION, "arb": _BILLION,
    "kharab": Decimal(10) ** 11, "kharabs": Decimal(10) ** 11,
    "t": _TRILLION, "tn": _TRILLION, "trillion": _TRILLION, "trillions": _TRILLION,
    "neel": Decimal(10) ** 13,
}

_PREFIX = re.compile(r"^(?:rs\.?|inr|usd|eur|gbp|rupees?|dollars?|euros?|pounds?|[$€£₹])\s*", re.IGNORECASE)
_NUMBER = r"(?:\d{1,3}(?:,\d{2,3})+|\d+)(?:\.\d+)?|\.\d+"
_PART = re.compile(r"\s*(%s)\s*([A-Za-z]*)\.?\s*" % _NUMBER)


class _Span:
    """Minimal stand-in for a match object that only knows where it ends."""

    def __init__(self, end):
        self._end = end

    def end(self):
        return self._end


def parse_units(text):
    """-> Decimal. '2 cr' -> 20000000, '1.5k' -> 1500, '1 crore 25 lakh' -> 12500000, '(2L)' -> -200000."""
    if text is None or not str(text).strip():
        raise ValueError("empty text")
    s = str(text).strip()
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative, s = True, s[1:-1].strip()
    if s[:1] in "+-−–":
        negative = negative or s[0] in "-−–"
        s = s[1:].strip()
    s = _PREFIX.sub("", s).strip()
    s = re.sub(r"(?i)\b(?:and|only)\b", " ", s).strip()
    if not s:
        raise ValueError("no number found")
    total, pos = Decimal(0), 0
    while pos < len(s):
        m = _PART.match(s, pos)
        if not m:
            raise ValueError("cannot read %r" % s[pos:].strip())
        number = Decimal(m.group(1).replace(",", ""))
        suffix = m.group(2).lower()
        if suffix:
            if suffix not in SUFFIXES:
                raise ValueError("unknown unit %r" % m.group(2))
            number *= SUFFIXES[suffix]
        while suffix:                                         # "1 lakh crore": a second unit multiplies
            extra = re.match(r"([A-Za-z]+)\.?\s*", s[m.end():])
            if not extra or extra.group(1).lower() not in SUFFIXES:
                break
            number *= SUFFIXES[extra.group(1).lower()]
            m = _Span(m.end() + extra.end())
        total += number
        pos = m.end()
        if pos < len(s) and s[pos] in ",&+":                  # '2 lakh, 50 thousand'
            pos += 1
        if not suffix and s[pos:].strip():                   # '1,5L': a bare number cannot start a compound
            raise ValueError("cannot read %r" % s[pos:].strip())
    return -total if negative else total
