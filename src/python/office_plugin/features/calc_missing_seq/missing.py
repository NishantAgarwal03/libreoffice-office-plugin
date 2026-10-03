"""Find gaps in a numbered list (pure Python, no uno).

Purpose: parse sequence values (plain numbers, or text such as INV-0012 / GRN/045/26), walk them in
         sheet order and report which numbers are missing, without ever building huge lists.
Inputs:  the column's cell values (float / str / ''), expected step, direction.
Outputs: Gap records (where to insert, how many, how to format them) and counters.
Must not: import uno.
"""

import re
from collections import namedtuple

_TEXT = re.compile(r"^(.*?)(\d+)(\D*)$", re.DOTALL)

# series identity: numbers of one series share prefix/suffix/kind; width = zero padding (0 = none)
Series = namedtuple("Series", "prefix suffix numeric width")
Item = namedtuple("Item", "row number series")
# `before_row`: data-row index of the item AFTER the gap (insert the placeholders just above it);
# missing numbers are first, first+step_signed, ... (count of them)
Gap = namedtuple("Gap", "before_row first step_signed count series")
Result = namedtuple("Result", "gaps total_missing duplicates out_of_order off_step skipped")


def parse(value, row):
    """-> Item or None (blank / no number in it / non-integer value)."""
    if isinstance(value, bool) or value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        if float(value) != int(value):
            return None
        return Item(row, int(value), Series("", "", True, 0))
    match = _TEXT.match(str(value).strip())
    if not match:
        return None
    prefix, digits, suffix = match.groups()
    width = len(digits) if len(digits) > 1 and digits[0] == "0" else 0
    return Item(row, int(digits), Series(prefix, suffix, False, width))


def format_number(series, number):
    """Number -> what to write in the cell: a float for numeric series, text for prefixed ones."""
    if series.numeric:
        return float(number)
    return "%s%s%s" % (series.prefix, str(number).zfill(series.width), series.suffix)


def gap_numbers(gap):
    return [gap.first + gap.step_signed * k for k in range(gap.count)]


def find_gaps(values, step=1, descending=False):
    """Walk `values` (sheet order) and find missing numbers.

    A gap is where the next value is more than one `step` beyond the running extreme. Duplicates
    and out-of-order values never move the running extreme and are only counted. A change of
    prefix/suffix starts a new series (no gap across it).
    """
    if step <= 0 or step != int(step):
        raise ValueError("Step must be a whole number of at least 1")
    step = int(step)
    sign = -1 if descending else 1
    gaps, total, dup, ooo, off, skipped = [], 0, 0, 0, 0, 0
    prev = None
    for row, value in enumerate(values):
        item = parse(value, row)
        if item is None:
            skipped += 1 if (value is not None and value != "") else 0
            continue
        if prev is None or prev.series[:3] != item.series[:3]:
            prev = item
            continue
        diff = (item.number - prev.number) * sign
        if diff == 0:
            dup += 1
        elif diff < 0:
            ooo += 1
        else:
            if diff % step:
                off += 1
            count = (diff - 1) // step               # whole steps strictly between prev and item
            if count > 0:
                width = max(prev.series.width, item.series.width)
                gaps.append(Gap(row, prev.number + sign * step, sign * step, count,
                                prev.series._replace(width=width)))
                total += count
            prev = item
    return Result(gaps, total, dup, ooo, off, skipped)


def label(gap):
    """'1004-1007' or '1012' (formatted like the cells)."""
    first = format_number(gap.series, gap.first)
    if gap.count == 1:
        return _show(first)
    last = format_number(gap.series, gap.first + gap.step_signed * (gap.count - 1))
    return "%s-%s" % (_show(first), _show(last))


def _show(value):
    return str(int(value)) if isinstance(value, float) else str(value)


def describe(result, limit=8):
    """'1004-1007, 1012, ... (3 more gaps)'."""
    parts = [label(g) for g in result.gaps[:limit]]
    text = ", ".join(parts)
    if len(result.gaps) > limit:
        text += ", ... (%d more gap(s))" % (len(result.gaps) - limit)
    return text
