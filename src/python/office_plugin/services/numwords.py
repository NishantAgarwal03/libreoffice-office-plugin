"""Numbers to words and words to numbers (pure Python, no uno).

Purpose: Indian (thousand, lakh, crore, arab, kharab, neel) and international (thousand, million,
         billion, trillion) number words, with negatives, decimals and currency wording, and the
         reverse parser that understands both systems at once.
Inputs:  numbers (int / float / Decimal) or text; options.
Outputs: text, or Decimal; ValueError with a readable message.
Must not: import uno.
"""

import re
from collections import namedtuple
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation

INDIAN, INTERNATIONAL = "indian", "international"
MAX_VALUE = Decimal(10) ** 15 - 1

_ONES = ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven",
         "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"]
_TENS = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"]
_INDIAN_SCALES = ["thousand", "lakh", "crore", "arab", "kharab", "neel"]        # 1e3, 1e5, 1e7, 1e9, 1e11, 1e13
_INTL_SCALES = ["thousand", "million", "billion", "trillion"]                    # 1e3, 1e6, 1e9, 1e12

Currency = namedtuple("Currency", "major_singular major_plural minor_singular minor_plural")
CURRENCIES = {
    "Rupees / Paise": Currency("Rupee", "Rupees", "Paisa", "Paise"),
    "Dollars / Cents": Currency("Dollar", "Dollars", "Cent", "Cents"),
    "Euros / Cents": Currency("Euro", "Euros", "Cent", "Cents"),
    "Pounds / Pence": Currency("Pound", "Pounds", "Penny", "Pence"),
}
SENTENCE, TITLE, UPPER, LOWER = "Sentence case", "Title Case", "UPPER CASE", "lower case"


# ======================================================================================
# number -> words
# ======================================================================================
def _below_thousand(n, use_and):
    parts = []
    hundreds, rest = divmod(n, 100)
    if hundreds:
        parts.append(_ONES[hundreds] + " hundred")
    if rest:
        if hundreds and use_and:
            parts.append("and")
        parts.append(_ONES[rest] if rest < 20 else _TENS[rest // 10] + ("-" + _ONES[rest % 10] if rest % 10 else ""))
    return " ".join(parts)


def integer_to_words(n, system=INDIAN, use_and=False):
    """Non-negative integer below 10^15 -> lower-case words."""
    n = int(n)
    if n < 0 or n > MAX_VALUE:
        raise ValueError("number out of range (0 to 999 trillion)")
    if n == 0:
        return "zero"
    parts = []
    low = n % 1000
    n //= 1000
    if system == INDIAN:
        for scale in _INDIAN_SCALES:
            chunk, n = n % 100, n // 100
            if chunk:
                parts.append(_below_thousand(chunk, False) + " " + scale)
            if not n:
                break
    else:
        for scale in _INTL_SCALES:
            chunk, n = n % 1000, n // 1000
            if chunk:
                parts.append(_below_thousand(chunk, use_and) + " " + scale)
            if not n:
                break
    parts.reverse()
    if low:
        text = _below_thousand(low, use_and)
        if use_and and parts and low < 100:                   # "one thousand and five"
            text = "and " + text
        parts.append(text)
    return " ".join(parts)


def _to_decimal(value):
    if isinstance(value, Decimal):
        return value
    if isinstance(value, bool):
        raise ValueError("not a number")
    # 15 significant digits is what Calc shows: 1.1+2.2 is stored as 3.3000000000000003 but reads 3.3
    return Decimal(format(float(value), ".15g")) if isinstance(value, float) else Decimal(int(value))


def _apply_case(text, case):
    if case == UPPER:
        return text.upper()
    if case == LOWER:
        return text.lower()
    if case == TITLE:
        return " ".join("-".join(p.capitalize() for p in w.split("-")) for w in text.split(" "))
    return text[:1].upper() + text[1:]


def to_words(value, system=INDIAN, currency=None, case=SENTENCE, use_and=False, only=False, max_decimals=10):
    """Words for a number.

    currency: a key of CURRENCIES or None. With a currency the value is rounded to 2 decimals and
    written as 'Rupees one lakh and Paise fifty Only' (prefix style); without it decimals read
    'twelve point four five'."""
    number = _to_decimal(value)
    if not number.is_finite():
        raise ValueError("not a finite number")
    negative = number < 0
    number = abs(number)
    if currency:
        names = CURRENCIES[currency]
        number = number.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        major, minor = int(number), int((number - int(number)) * 100)
        words = "%s %s" % (names.major_plural, integer_to_words(major, system, use_and))
        if minor:
            words += " and %s %s" % (names.minor_plural, integer_to_words(minor, system, use_and))
    else:
        integer = int(number)
        words = integer_to_words(integer, system, use_and)
        fraction = format(number - integer, "f").split(".")[1].rstrip("0") if number != integer else ""
        if fraction:
            if len(fraction) > max_decimals:
                raise ValueError("more than %d decimal places" % max_decimals)
            words += " point " + " ".join(_ONES[int(d)] for d in fraction)
    if negative and number != 0:
        words = "minus " + words
    if only and (currency or only == "always"):
        words += " only"
    return _apply_case(words, case)


# ======================================================================================
# words -> number
# ======================================================================================
_UNITS = {w: i for i, w in enumerate(_ONES)}
_UNITS.update({"nought": 0, "nil": 0, "oh": 0, "o": 0})
_TENS_MAP = {w: i * 10 for i, w in enumerate(_TENS) if w}
_TENS_MAP.update({"fourty": 40})
_SCALES = {"thousand": 10 ** 3, "lakh": 10 ** 5, "lakhs": 10 ** 5, "lac": 10 ** 5, "lacs": 10 ** 5, "lack": 10 ** 5,
           "crore": 10 ** 7, "crores": 10 ** 7, "million": 10 ** 6, "millions": 10 ** 6, "arab": 10 ** 9,
           "arabs": 10 ** 9, "billion": 10 ** 9, "billions": 10 ** 9, "kharab": 10 ** 11, "kharabs": 10 ** 11,
           "trillion": 10 ** 12, "trillions": 10 ** 12, "neel": 10 ** 13, "neels": 10 ** 13}
_MINOR = {"paise", "paisa", "paisas", "cent", "cents", "pence", "penny", "pennies"}
_MAJOR = {"rupee", "rupees", "rs", "inr", "dollar", "dollars", "usd", "euro", "euros", "eur", "pound", "pounds", "gbp"}
_FILLER = {"and", "only", "of", "a", "the", "total", "amount", "rs", "inr", "usd", "eur", "gbp"}


def _tokens(text):
    text = text.lower().replace("&", " and ").replace("-", " ").replace(",", " ")
    text = re.sub(r"[.]", " . ", text)
    return [t for t in text.split() if t]


def _parse_integer(tokens):
    """Number words (no decimals / currency) -> int. Unknown words raise ValueError."""
    total = current = last_chunk = 0
    after_scale = False
    seen = False
    for tok in tokens:
        if tok in _FILLER or tok in _MAJOR:
            continue
        if tok in _UNITS:
            current += _UNITS[tok]
            after_scale = False
        elif tok in _TENS_MAP:
            current += _TENS_MAP[tok]
            after_scale = False
        elif tok in ("hundred", "hundreds"):
            current = (current or 1) * 100
            after_scale = False
        elif tok in _SCALES:
            if after_scale and not current:                  # "one lakh crore": scales multiply
                total -= last_chunk
                last_chunk *= _SCALES[tok]
            else:
                last_chunk = (current or 1) * _SCALES[tok]
            total += last_chunk
            current = 0
            after_scale = True
        else:
            raise ValueError("unrecognised word %r" % tok)
        seen = True
    if not seen:
        raise ValueError("no number words found")
    return total + current


def _parse_decimal_part(tokens):
    digits = ""
    for tok in tokens:
        if tok in _UNITS and _UNITS[tok] < 10:
            digits += str(_UNITS[tok])
        elif tok in _FILLER:
            continue
        else:
            raise ValueError("unrecognised word %r after 'point'" % tok)
    if not digits:
        raise ValueError("nothing after 'point'")
    return Decimal("0." + digits)


def _parse_plain(tokens):
    if "point" in tokens:
        i = tokens.index("point")
        whole = _parse_integer(tokens[:i]) if any(t not in _FILLER for t in tokens[:i]) else 0
        return Decimal(whole) + _parse_decimal_part(tokens[i + 1:])
    return Decimal(_parse_integer(tokens))


def parse_words(text):
    """'one lakh twenty-three thousand', 'Rupees twelve and Paise fifty Only', 'minus two crore'
    -> Decimal. Indian and international words are both understood. Raises ValueError."""
    if text is None or not str(text).strip():
        raise ValueError("empty text")
    tokens = _tokens(str(text))
    negative = False
    while tokens and tokens[0] in ("minus", "negative", "neg"):
        negative = True
        tokens = tokens[1:]
    tokens = [t for t in tokens if t != "."]
    minor_at = [i for i, t in enumerate(tokens) if t in _MINOR]
    if minor_at:
        i = minor_at[0]
        before, after = tokens[:i], tokens[i + 1:]
        after_numeric = [t for t in after if t not in _FILLER]
        if after_numeric:                                   # prefix style: "... and paise fifty"
            major_tokens, minor_tokens = before, after
        else:                                               # suffix style: "twelve rupees and fifty paise"
            cut = max([k for k, t in enumerate(before) if t in _MAJOR or t == "and"] + [-1])
            major_tokens, minor_tokens = before[:cut + 1], before[cut + 1:]
        major = _parse_plain(major_tokens) if any(t not in _FILLER and t not in _MAJOR for t in major_tokens) else Decimal(0)
        minor = _parse_integer(minor_tokens)
        if minor >= 100:
            raise ValueError("the paise/cents part must be below 100")
        value = major + Decimal(minor) / 100
    else:
        value = _parse_plain(tokens)
    return -value if negative else value
