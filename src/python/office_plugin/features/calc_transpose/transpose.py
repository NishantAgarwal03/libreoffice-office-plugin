"""Transpose geometry (pure Python, no uno).

Purpose: rectangle maths for transposing a table: target shape, the cells a transpose would
         overwrite outside the source, bounds checks, and the swapped grid for the fallback.
Inputs:  Rect(row1, col1, row2, col2) 0-based inclusive rectangles, grids as lists of rows.
Outputs: rectangles / lists of rectangles / a transposed grid.
Must not: import uno.
"""

from collections import namedtuple

Rect = namedtuple("Rect", "row1 col1 row2 col2")

IN_PLACE = "In place (replace the table)"
NEW_SHEET = "New sheet"
AT_CELL = "At cell..."
DESTINATIONS = (IN_PLACE, NEW_SHEET, AT_CELL)


def rows_of(rect):
    return rect.row2 - rect.row1 + 1


def cols_of(rect):
    return rect.col2 - rect.col1 + 1


def cell_count(rect):
    return rows_of(rect) * cols_of(rect)


def transposed_rect(src, anchor_row, anchor_col):
    """Where the transposed table lands when its top-left cell is (anchor_row, anchor_col)."""
    return Rect(anchor_row, anchor_col, anchor_row + cols_of(src) - 1, anchor_col + rows_of(src) - 1)


def intersects(a, b):
    return not (a.row2 < b.row1 or b.row2 < a.row1 or a.col2 < b.col1 or b.col2 < a.col1)


def rect_difference(a, b):
    """Cells of `a` that are not in `b`, as up to four non-overlapping rectangles."""
    if not intersects(a, b):
        return [a]
    parts = []
    if a.row1 < b.row1:                                                       # strip above b
        parts.append(Rect(a.row1, a.col1, b.row1 - 1, a.col2))
    if b.row2 < a.row2:                                                       # strip below b
        parts.append(Rect(b.row2 + 1, a.col1, a.row2, a.col2))
    top, bottom = max(a.row1, b.row1), min(a.row2, b.row2)
    if a.col1 < b.col1:                                                       # left of b (middle band)
        parts.append(Rect(top, a.col1, bottom, b.col1 - 1))
    if b.col2 < a.col2:                                                       # right of b
        parts.append(Rect(top, b.col2 + 1, bottom, a.col2))
    return parts


def fits(rect, max_row, max_col):
    """True when the rectangle lies inside a sheet of max_row+1 rows and max_col+1 columns."""
    return rect.row1 >= 0 and rect.col1 >= 0 and rect.row2 <= max_row and rect.col2 <= max_col


def swap_grid(grid):
    """Transposed copy of a rectangular grid (list of row tuples)."""
    return [tuple(col) for col in zip(*grid)] if grid else []


def describe(src):
    """'10 rows x 4 columns -> 4 rows x 10 columns' for the dialog preview."""
    return "%d row(s) x %d column(s)  ->  %d row(s) x %d column(s)" % (
        rows_of(src), cols_of(src), cols_of(src), rows_of(src))
