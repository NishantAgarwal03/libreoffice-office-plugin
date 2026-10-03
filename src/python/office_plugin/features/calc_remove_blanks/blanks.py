"""Blank row / column / cell detection (pure Python, no uno).

Purpose: given a mask of blank cells (list of rows of bool), find the blank rows, blank columns,
         and the runs of blank cells to delete.
Inputs:  mask[r][c] is True when the cell is blank.
Outputs: index lists / runs relative to the mask (0-based).
Must not: import uno.
"""


def blank_rows(mask):
    return [r for r, row in enumerate(mask) if row and all(row)]


def blank_cols(mask):
    width = max((len(r) for r in mask), default=0)
    return [c for c in range(width)
            if mask and all(c >= len(row) or row[c] for row in mask)]


def runs(indices):
    """[2,3,4,9] -> [(2,4),(9,9)] (inclusive ascending runs)."""
    out = []
    for i in sorted(indices):
        if out and i == out[-1][1] + 1:
            out[-1] = (out[-1][0], i)
        else:
            out.append((i, i))
    return out


def blank_cell_runs(mask, axis):
    """Runs of blank cells for shift-up ('up': per column) or shift-left ('left': per row).

    Returns [(line, start, end)] where `line` is the column (up) or row (left) index and
    start..end are inclusive row (up) or column (left) indexes.
    """
    result = []
    if axis == "up":
        width = max((len(r) for r in mask), default=0)
        for c in range(width):
            blanks = [r for r, row in enumerate(mask) if c >= len(row) or row[c]]
            result += [(c, a, b) for a, b in runs(blanks)]
    elif axis == "left":
        for r, row in enumerate(mask):
            result += [(r, a, b) for a, b in runs([c for c, blank in enumerate(row) if blank])]
    else:
        raise ValueError("axis must be 'up' or 'left'")
    return result


def count_blank_cells(mask):
    return sum(1 for row in mask for blank in row if blank)
