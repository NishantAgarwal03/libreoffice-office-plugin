"""Fill-blanks planning (pure Python, no uno).

Purpose: decide which blank cells take which value, given a grid of cell values.
Inputs:  grid (list of row tuples), direction 'down' | 'right', a predicate telling blanks.
Outputs: runs as (row1, col1, row2, col2, value): rectangles (1 wide or 1 tall) to fill.
Must not: import uno.
"""


def plan_fill(grid, direction="down", is_blank=lambda v: v == "" or v is None):
    """Each blank cell takes the nearest non-blank value above it (down) or to its left (right),
    inside the grid. Blanks with nothing before them stay blank. Contiguous cells taking the
    same source value are returned as one run."""
    runs = []
    rows = len(grid)
    cols = max((len(r) for r in grid), default=0)

    def cell(r, c):
        row = grid[r]
        return row[c] if c < len(row) else ""

    if direction == "down":
        for c in range(cols):
            last, start = None, None
            for r in range(rows + 1):
                filling = r < rows and last is not None and is_blank(cell(r, c))
                if filling and start is None:
                    start = r
                if not filling and start is not None:
                    runs.append((start, c, r - 1, c, last))
                    start = None
                if r < rows and not is_blank(cell(r, c)):
                    last = cell(r, c)
    elif direction == "right":
        for r in range(rows):
            last, start = None, None
            for c in range(cols + 1):
                filling = c < cols and last is not None and is_blank(cell(r, c))
                if filling and start is None:
                    start = c
                if not filling and start is not None:
                    runs.append((r, start, r, c - 1, last))
                    start = None
                if c < cols and not is_blank(cell(r, c)):
                    last = cell(r, c)
    else:
        raise ValueError("direction must be 'down' or 'right'")
    return runs


def count_cells(runs):
    return sum((r2 - r1 + 1) * (c2 - c1 + 1) for r1, c1, r2, c2, _ in runs)
