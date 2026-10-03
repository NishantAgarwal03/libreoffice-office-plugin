"""Cell functions NUMTOWORDS / WORDSTONUM / UNITSTONUM (pure Python, no uno).

Purpose: the logic behind the Calc formulas installed by 'Number Functions' (a small Basic
         module in My Macros calls FunctionJob, which calls call()).
Inputs:  the function name and up to four argument values as Calc passes them (float, str, None).
Outputs: str or float; ValueError for anything invalid (Calc shows #VALUE!).
Must not: import uno.
"""

from decimal import Decimal, InvalidOperation

from office_plugin.services import numwords as N
from office_plugin.services.numunits import parse_units

FUNCTIONS = ("NUMTOWORDS", "WORDSTONUM", "UNITSTONUM")
ERROR_TEXT = "#VALUE!"            # what a cell shows for bad input (Basic cannot raise a real cell error)

_SYSTEMS = {"indian": N.INDIAN, "in": N.INDIAN, "lakh": N.INDIAN, "lakhs": N.INDIAN, "crore": N.INDIAN,
            "international": N.INTERNATIONAL, "intl": N.INTERNATIONAL, "int": N.INTERNATIONAL,
            "western": N.INTERNATIONAL, "million": N.INTERNATIONAL}
_CURRENCIES = {"inr": "Rupees / Paise", "rs": "Rupees / Paise", "rupee": "Rupees / Paise",
               "rupees": "Rupees / Paise", "₹": "Rupees / Paise",
               "usd": "Dollars / Cents", "dollar": "Dollars / Cents", "dollars": "Dollars / Cents", "$": "Dollars / Cents",
               "eur": "Euros / Cents", "euro": "Euros / Cents", "euros": "Euros / Cents", "€": "Euros / Cents",
               "gbp": "Pounds / Pence", "pound": "Pounds / Pence", "pounds": "Pounds / Pence", "£": "Pounds / Pence"}
_CASES = {"sentence": N.SENTENCE, "title": N.TITLE, "upper": N.UPPER, "lower": N.LOWER}


def _blank(value):
    return value is None or (isinstance(value, str) and not value.strip())


def _number(value):
    if isinstance(value, bool) or _blank(value):
        raise ValueError("a number is required")
    if isinstance(value, (int, float)):
        return value
    try:
        return float(Decimal(str(value).strip().replace(",", "")))
    except InvalidOperation:
        raise ValueError("%r is not a number" % value)


def _text(value):
    if _blank(value):
        raise ValueError("text is required")
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value)


def _choice(value, table, default, what):
    if _blank(value) or (isinstance(value, (int, float)) and value == 0):    # a skipped ";;" argument may arrive as 0
        return default
    try:
        return table[str(value).strip().lower()]
    except KeyError:
        raise ValueError("unknown %s %r (use %s)" % (what, value, ", ".join(sorted(set(table)))))


def call(fn, a=None, b=None, c=None, d=None):
    """NUMTOWORDS(number; [system]; [currency]; [letters]) / WORDSTONUM(text) / UNITSTONUM(text)."""
    name = str(fn).upper()
    if name == "NUMTOWORDS":
        system = _choice(b, _SYSTEMS, N.INDIAN, "numbering system")
        currency = _choice(c, _CURRENCIES, None, "currency")
        letters = _choice(d, _CASES, N.SENTENCE, "letter style")
        return N.to_words(_number(a), system, currency, letters, only=bool(currency))
    if name == "WORDSTONUM":
        return float(N.parse_words(_text(a)))
    if name == "UNITSTONUM":
        return float(parse_units(_text(a)))
    raise ValueError("unknown function %r" % fn)
