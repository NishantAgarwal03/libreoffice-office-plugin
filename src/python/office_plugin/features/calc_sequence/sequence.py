"""Number series from start / step / end / count (pure Python, no uno).

Purpose: compute the values of a series from any sufficient combination of start, step, end and
         count, using decimal arithmetic so 0.1 steps stay exact.
Inputs:  numeric text from the dialog (blank = not given) and a default count (selection size).
Outputs: list of int/float values; ValueError with a readable message on bad input.
Must not: import uno.
"""

import re
from decimal import ROUND_FLOOR, Decimal, InvalidOperation

MAX_COUNT = 1_000_000
MAX_MAGNITUDE = Decimal("1e15")
_PLAIN = re.compile(r"^[+-]?(\d+(\.\d*)?|\.\d+)$")        # no exponents, no NaN/Infinity


def parse_number(text, label):
    """'' -> None; '2.5' -> Decimal('2.5'); anything else -> ValueError naming the field."""
    text = (text or "").strip().replace(",", ".") if text is not None else ""
    if text == "":
        return None
    if not _PLAIN.match(text):
        raise ValueError("%s must be a plain number (got %r)" % (label, text))
    try:
        value = Decimal(text)
    except InvalidOperation:
        raise ValueError("%s must be a number (got %r)" % (label, text))
    if abs(value) > MAX_MAGNITUDE:
        raise ValueError("%s is too large" % label)
    return value


def _as_number(value):
    return int(value) if value == value.to_integral_value() else float(value)


def build_series(start, step=None, end=None, count=None, default_count=None):
    """Values of the series as ints/floats.

    Decimal (or None) arguments:
      start + step + count   -> `count` values
      start + step + end     -> values up to and including `end` (when the step lands on it)
      start + end + count    -> step = (end - start) / (count - 1)
      start + step only      -> `default_count` values (the selection size)
      all of step, end, count -> `count` values, stopped early at `end`
    """
    if start is None:
        raise ValueError("Start is required")
    if count is not None and (count != count.to_integral_value() or count < 1):
        raise ValueError("Count must be a whole number of at least 1")
    if step is not None and step == 0:
        raise ValueError("Step cannot be 0")

    if step is None:
        if end is None or count is None:
            step = Decimal(1)                       # plain 'start, count/selection' fills 1, 2, 3...
            if end is not None:
                step = Decimal(1) if end >= start else Decimal(-1)
        elif count == 1:
            return [_as_number(start)]
        else:
            step = (end - start) / (count - 1)
            if step == 0:
                raise ValueError("Start and End are equal: nothing to step through")

    limit = None                                    # number of values allowed by `end`
    if end is not None:
        distance = (end - start) / step
        if distance < 0:
            raise ValueError("Step %s moves away from End %s; use the opposite sign" % (step, end))
        limit = int(distance.to_integral_value(rounding=ROUND_FLOOR)) + 1

    if count is not None:
        n = count
        if limit is not None and step is not None:
            n = min(int(count), limit)
    elif limit is not None:
        n = limit
    elif default_count:
        n = default_count
    else:
        raise ValueError("Give a Count or an End (or select the cells to fill)")
    n = int(n)
    if n > MAX_COUNT:
        raise ValueError("That would create %d values (limit %d)" % (n, MAX_COUNT))
    return [_as_number(start + step * i) for i in range(n)]


def preview(values, shown=4):
    """'1, 3, 5, 7 ... 99  (50 values)'."""
    if len(values) <= shown + 1:
        return ", ".join(str(v) for v in values)
    return "%s ... %s  (%d values)" % (", ".join(str(v) for v in values[:shown]), values[-1], len(values))
