"""Which rows / cells to highlight (pure Python, no uno).

Purpose: compute the row indexes (or cells) to colour for each highlight mode.
Inputs:  data rows of list A (and B), key column indexes, CleanOptions.
Outputs: sets of 0-based row indexes, or {(row, col)} for cell-by-cell differences.
Must not: import uno.
"""

from collections import Counter

from office_plugin.services.palette import PALETTE  # noqa: F401  (re-exported: colours the plugin paints)
from office_plugin.services.cellkeys import CleanOptions, key_is_blank, normalize, row_key

DUP_ALL = "Duplicates (every occurrence)"
DUP_REPEATS = "Duplicates (2nd and later only)"
UNIQUE = "Unique values (appear once)"
IN_BOTH = "Rows of A that are also in B"
ONLY_IN_A = "Rows of A that are not in B"
CELL_DIFF = "Cell-by-cell differences (A vs B)"
MODES = (DUP_ALL, DUP_REPEATS, UNIQUE, IN_BOTH, ONLY_IN_A, CELL_DIFF)
NEEDS_B = {IN_BOTH, ONLY_IN_A, CELL_DIFF}
SUGGESTED_COLOUR = {DUP_ALL: "Yellow", DUP_REPEATS: "Orange", UNIQUE: "Blue",
                    IN_BOTH: "Green", ONLY_IN_A: "Red", CELL_DIFF: "Red"}



def _keys(rows, key_cols, opts):
    out = []
    for row in rows:
        key = row_key(row, key_cols, opts)
        out.append(None if key_is_blank(key) else key)     # rows with a blank key never match
    return out


def duplicates(rows, key_cols, opts=CleanOptions(), repeats_only=False):
    keys = _keys(rows, key_cols, opts)
    counts = Counter(k for k in keys if k is not None)
    seen, marked = set(), set()
    for i, k in enumerate(keys):
        if k is None or counts[k] < 2:
            continue
        if not repeats_only or k in seen:
            marked.add(i)
        seen.add(k)
    return marked


def unique(rows, key_cols, opts=CleanOptions()):
    keys = _keys(rows, key_cols, opts)
    counts = Counter(k for k in keys if k is not None)
    return {i for i, k in enumerate(keys) if k is not None and counts[k] == 1}


def membership(rows_a, rows_b, key_cols, opts=CleanOptions(), present=True):
    """-> (indexes in A, indexes in B): rows whose key IS (present) / is NOT in the other list."""
    ka, kb = _keys(rows_a, key_cols, opts), _keys(rows_b, key_cols, opts)
    set_a, set_b = {k for k in ka if k is not None}, {k for k in kb if k is not None}

    def pick(keys, other):
        return {i for i, k in enumerate(keys) if k is not None and ((k in other) == present)}

    return pick(ka, set_b), pick(kb, set_a)


def cell_diffs(grid_a, grid_b, opts=CleanOptions()):
    """-> {(row, col)} where A and B (same position) hold different values."""
    if len(grid_a) != len(grid_b) or any(len(a) != len(b) for a, b in zip(grid_a, grid_b)):
        raise ValueError("the two ranges must have the same size")
    return {(r, c) for r, (ra, rb) in enumerate(zip(grid_a, grid_b))
            for c, (a, b) in enumerate(zip(ra, rb)) if normalize(a, opts) != normalize(b, opts)}
