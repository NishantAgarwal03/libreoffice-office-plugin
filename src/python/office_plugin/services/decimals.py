"""Fixed-decimals logic: ROUND wrapping of formulas, value rounding, number-format rewriting.

Purpose: give every number the same number of decimals while keeping formulas as formulas
         (wrapped in ROUND/ROUNDUP/ROUNDDOWN, replacing an existing outer wrapper instead of
         nesting) and values as rounded values, and make cells display exactly that many decimals.
Inputs:  formula text in API grammar ('=A1*B1', ';' separators), numbers, number-format codes.
Outputs: new formula text, rounded numbers, new format codes.
Must not: import uno.
"""

import re
from decimal import ROUND_DOWN, ROUND_HALF_UP, ROUND_UP, Decimal

ROUND, ROUND_UP_MODE, ROUND_DOWN_MODE = "Round (half up)", "Round up (away from zero)", "Round down (toward zero)"
MODES = (ROUND, ROUND_UP_MODE, ROUND_DOWN_MODE)
_FUNCTION = {ROUND: "ROUND", ROUND_UP_MODE: "ROUNDUP", ROUND_DOWN_MODE: "ROUNDDOWN"}
_QUANTIZE = {ROUND: ROUND_HALF_UP, ROUND_UP_MODE: ROUND_UP, ROUND_DOWN_MODE: ROUND_DOWN}
MAX_DECIMALS = 15


# ======================================================================================
# values
# ======================================================================================
def round_value(value, places, mode=ROUND):
    """Round like Calc: 2.675 -> 2.68 (via the shortest decimal representation, not the binary one)."""
    if not 0 <= places <= MAX_DECIMALS:
        raise ValueError("decimal places must be 0 to %d" % MAX_DECIMALS)
    dec = Decimal(repr(float(value)))
    return float(dec.quantize(Decimal(1).scaleb(-places), rounding=_QUANTIZE[mode]))


# ======================================================================================
# formulas
# ======================================================================================
def _scan(text):
    """Yield (index, char, depth, in_text) with quoted strings / sheet names marked as text."""
    depth, quote = 0, None
    for i, ch in enumerate(text):
        if quote:
            yield i, ch, depth, True
            if ch == quote:
                quote = None
            continue
        if ch in "\"'":
            quote = ch
            yield i, ch, depth, True
            continue
        if ch == "(":
            yield i, ch, depth, False
            depth += 1
        elif ch == ")":
            depth -= 1
            yield i, ch, depth, False
        else:
            yield i, ch, depth, False


def _split_args(inner):
    """Top-level ';' split of function arguments (quotes and nested parentheses respected)."""
    args, start = [], 0
    for i, ch, depth, in_text in _scan(inner):
        if ch == ";" and depth == 0 and not in_text:
            args.append(inner[start:i])
            start = i + 1
    args.append(inner[start:])
    return args


def unwrap(formula):
    """If the whole formula is ROUND/ROUNDUP/ROUNDDOWN(expr; integer) return (expr, function_name),
    else (None, None). `formula` starts with '='."""
    body = formula[1:] if formula.startswith("=") else formula
    m = re.match(r"(ROUNDDOWN|ROUNDUP|ROUND)\(", body, re.IGNORECASE)
    if not m or not body.endswith(")"):
        return None, None
    open_index = m.end() - 1
    for i, ch, depth, in_text in _scan(body):
        if i >= open_index and ch == ")" and not in_text and depth == 0:
            if i != len(body) - 1:
                return None, None                 # the first parenthesis closes before the end
            break
    inner = body[open_index + 1:-1]
    args = _split_args(inner)
    if len(args) != 2 or not re.fullmatch(r"\s*-?\d+\s*", args[1]):
        return None, None
    return args[0], m.group(1).upper()


def wrap_round(formula, places, mode=ROUND):
    """'=A1*B1' -> '=ROUND(A1*B1;2)'; an existing outer ROUND/ROUNDUP/ROUNDDOWN is replaced, not nested.
    Returns (new_formula, status) with status 'wrapped' | 'updated' | 'same'."""
    function = _FUNCTION[mode]
    inner, existing = unwrap(formula)
    body = formula[1:] if inner is None else inner
    new = "=%s(%s;%d)" % (function, body, places)
    if inner is None:
        return new, "wrapped"
    return new, ("same" if new.replace(" ", "") == formula.replace(" ", "") else "updated")


def remove_round(formula):
    """'=ROUND(A1*B1;2)' -> '=A1*B1'; None when there is no outer wrapper."""
    inner, _ = unwrap(formula)
    return None if inner is None else "=" + inner


# ======================================================================================
# number formats
# ======================================================================================
_SECTION_SPLIT = re.compile(r";(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)")        # ';' outside double quotes
_DIGITS = re.compile(r"([#0?][#0?,]*)(?:\.([0#?]*))?")


def _rewrite_section(section, places):
    """Change the decimals of the first digit placeholder in one format section."""
    if re.search(r"[eE][+-]", section):
        return section                                           # scientific: leave alone
    out, i, done = [], 0, False
    while i < len(section):
        ch = section[i]
        if ch == '"':                                            # literal text
            j = section.find('"', i + 1)
            j = len(section) - 1 if j < 0 else j
            out.append(section[i:j + 1])
            i = j + 1
        elif ch == "[":                                          # [Red], [$₹-4009] ...
            j = section.find("]", i)
            j = len(section) - 1 if j < 0 else j
            out.append(section[i:j + 1])
            i = j + 1
        elif ch == "\\" and i + 1 < len(section):
            out.append(section[i:i + 2])
            i += 2
        elif not done and ch in "#0?":
            m = _DIGITS.match(section, i)
            integer_part = m.group(1)
            out.append(integer_part + ("." + "0" * places if places else ""))
            i, done = m.end(), True
        else:
            out.append(ch)
            i += 1
    return "".join(out)


def format_with_decimals(code, places):
    """'#,##0.00' + 3 -> '#,##0.000'; 'General' -> '0.00'; '0.0%' + 2 -> '0.00%'; currency/colour kept."""
    if not 0 <= places <= MAX_DECIMALS:
        raise ValueError("decimal places must be 0 to %d" % MAX_DECIMALS)
    if code.strip().lower() == "general" or code.strip() == "":
        return "0" + ("." + "0" * places if places else "")
    return ";".join(_rewrite_section(s, places) for s in _SECTION_SPLIT.split(code))


def is_percent(code):
    """True when the format shows the value as a percentage (an unquoted % in the code)."""
    stripped = re.sub(r'"[^"]*"|\[[^\]]*\]|\.', "", code)
    return "%" in stripped
