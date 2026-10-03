"""Duplicate-row detection (pure Python, no uno)."""

from office_plugin.services.cellkeys import CleanOptions, key_is_blank, row_key


def col_name(index):
    """0 -> 'A', 25 -> 'Z', 26 -> 'AA'."""
    name = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def find_duplicate_rows(rows, key_cols, has_header=True, opts=CleanOptions()):
    """Return 0-based indices (into `rows`) of rows to delete.

    The first occurrence of each key is kept (Excel behaviour). `rows` is a sequence of
    row sequences (e.g. Calc getDataArray()). The header row, if any, is never reported.
    Values match as defined by `opts` (see services/cellkeys). Empty keys (every selected
    column blank) are not considered duplicates of each other.
    """
    if not key_cols:
        return []
    start = 1 if has_header else 0
    seen = set()
    dupes = []
    for index in range(start, len(rows)):
        key = row_key(rows[index], key_cols, opts)
        if key_is_blank(key):
            continue
        if key in seen:
            dupes.append(index)
        else:
            seen.add(key)
    return dupes


def group_runs(indices):
    """[2,3,4,9,11,12] -> [(2,4),(9,9),(11,12)] (inclusive runs, ascending)."""
    runs = []
    for i in sorted(indices):
        if runs and i == runs[-1][1] + 1:
            runs[-1] = (runs[-1][0], i)
        else:
            runs.append((i, i))
    return runs
