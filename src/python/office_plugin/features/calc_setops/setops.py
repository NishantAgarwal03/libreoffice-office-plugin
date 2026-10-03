"""Set operations on two lists of rows (pure Python, no uno).

Purpose: compute union / intersection / differences of two row lists, matching rows on chosen
         key columns with the shared "clean" normalisation.
Inputs:  data rows of list A and list B (header rows already removed), key column indexes,
         CleanOptions.
Outputs: result as [(source, index)] ('A' or 'B' plus the row index in that list) and counts.
Must not: import uno.
"""

from office_plugin.services.cellkeys import CleanOptions, key_is_blank, row_key

UNION = "Union"
INTERSECTION = "Intersection"
A_MINUS_B = "A minus B"
B_MINUS_A = "B minus A"
SYM_DIFF = "Symmetric difference"
OPERATIONS = (UNION, INTERSECTION, A_MINUS_B, B_MINUS_A, SYM_DIFF)

SHEET_NAMES = {UNION: "Union", INTERSECTION: "Intersection", A_MINUS_B: "A minus B",
               B_MINUS_A: "B minus A", SYM_DIFF: "Sym diff"}


def _distinct(rows, key_cols, opts):
    """-> (ordered [(index, key)] of first occurrences, skipped_blank_count). Blank keys skipped."""
    seen, out, blank = set(), [], 0
    for index, row in enumerate(rows):
        key = row_key(row, key_cols, opts)
        if key_is_blank(key):
            blank += 1
            continue
        if key not in seen:
            seen.add(key)
            out.append((index, key))
    return out, blank


def compute(op, rows_a, rows_b, key_cols, opts=CleanOptions()):
    """-> (result, counts). Rows are distinct by key (first occurrence kept); a row whose key
    columns are all blank never takes part."""
    if op not in OPERATIONS:
        raise ValueError("unknown operation %r" % op)
    if not key_cols:
        raise ValueError("no key columns")
    a, blank_a = _distinct(rows_a, key_cols, opts)
    b, blank_b = _distinct(rows_b, key_cols, opts)
    keys_a = {k for _, k in a}
    keys_b = {k for _, k in b}

    only_a = [("A", i) for i, k in a if k not in keys_b]
    only_b = [("B", i) for i, k in b if k not in keys_a]
    both_a = [("A", i) for i, k in a if k in keys_b]
    if op == UNION:
        result = [("A", i) for i, _ in a] + only_b
    elif op == INTERSECTION:
        result = both_a
    elif op == A_MINUS_B:
        result = only_a
    elif op == B_MINUS_A:
        result = only_b
    else:
        result = only_a + only_b

    counts = {"a_rows": len(rows_a), "b_rows": len(rows_b),
              "a_distinct": len(a), "b_distinct": len(b),
              "common": len(keys_a & keys_b), "result": len(result),
              "skipped_blank": blank_a + blank_b}
    return result, counts
