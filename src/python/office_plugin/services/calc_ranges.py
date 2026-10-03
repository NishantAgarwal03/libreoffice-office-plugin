"""Calc range helpers shared by the Calc data tools.

Purpose: turn selections / typed addresses into ranges, read them as grids, group row runs,
         and wrap edits in a single undo step.
Inputs:  UNO controller / model / ranges (the pure helpers take plain values).
Outputs: ranges, grids (list of tuples), address text.
Must not: import uno at module level (pure helpers stay testable without LibreOffice).
"""

import contextlib
import re

_SHEET_REF = re.compile(r"^\$?(?:'((?:[^']|'')+)'|([^.']+))\.(.+)$")


# ---- pure helpers -----------------------------------------------------------
def col_name(index):
    """0 -> 'A', 25 -> 'Z', 26 -> 'AA'."""
    name = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def column_labels(first_row, start_col, has_header):
    """Labels for a column picker: 'B - Item' (header text) or 'Column B'."""
    return [("%s - %s" % (col_name(start_col + i), v) if has_header and v != ""
             else "Column %s" % col_name(start_col + i)) for i, v in enumerate(first_row)]


def group_runs(indices):
    """[2,3,4,9,11,12] -> [(2,4),(9,9),(11,12)] (inclusive ascending runs)."""
    runs = []
    for i in sorted(indices):
        if runs and i == runs[-1][1] + 1:
            runs[-1] = (runs[-1][0], i)
        else:
            runs.append((i, i))
    return runs


def split_sheet_ref(text):
    """"'My Sheet'.A1:B2" / "$Sheet2.$A$1" -> (sheet_name | None, "A1:B2")."""
    text = text.strip()
    m = _SHEET_REF.match(text)
    if not m:
        return None, text.replace("$", "")
    name = (m.group(1) or m.group(2)).replace("''", "'")
    return name, m.group(3).replace("$", "")


def unique_name(existing, base):
    """base, base_2, base_3 ... not in `existing`."""
    taken = set(existing)
    if base not in taken:
        return base
    n = 2
    while "%s_%d" % (base, n) in taken:
        n += 1
    return "%s_%d" % (base, n)


# ---- UNO-bound helpers ------------------------------------------------------
def selected_ranges(controller, clip=True):
    """Ranges of the current selection (a Ctrl+multi-selection gives several). [] if none.

    An entire selected column / row / the whole sheet is clipped to the used area (see clip_to_used):
    it would otherwise make every tool read a million rows. Ranges the user drew keep their exact size.
    Pass clip=False for tools that WRITE into the selection (e.g. Sequence Numbers)."""
    sel = controller.getSelection()
    try:
        if sel.supportsService("com.sun.star.sheet.SheetCellRanges"):
            found = [sel.getByIndex(i) for i in range(sel.getCount())]
        elif sel.supportsService("com.sun.star.sheet.SheetCellRange"):
            found = [sel]
        else:
            return []
    except Exception:
        return []
    return [clip_to_used(r) for r in found] if clip else found


def clip_to_used(rng):
    """Limit an ENTIRE-column / entire-row / whole-sheet selection to the sheet's used area.

    Only a selection that reaches the sheet's last row (or last column) is clipped in that direction,
    so a deliberately drawn range such as A1:A10 keeps its exact size (Fill Blanks must still fill
    it down). At least the first row / column of the selection is kept."""
    a = rng.getRangeAddress()
    sheet = rng.getSpreadsheet()
    end_row, end_col = a.EndRow, a.EndColumn
    whole_rows = a.EndRow >= sheet.getRows().getCount() - 1
    whole_cols = a.EndColumn >= sheet.getColumns().getCount() - 1
    if whole_rows or whole_cols:
        used = used_area(sheet).getRangeAddress()
        if whole_rows:
            end_row = max(min(a.EndRow, used.EndRow), a.StartRow)
        if whole_cols:
            end_col = max(min(a.EndColumn, used.EndColumn), a.StartColumn)
    if (end_row, end_col) == (a.EndRow, a.EndColumn):
        return rng
    return sheet.getCellRangeByPosition(a.StartColumn, a.StartRow, end_col, end_row)


def expand_if_single(model, rng):
    """A lone cell becomes its current region (like Excel); other ranges are returned as is."""
    addr = rng.getRangeAddress()
    if addr.StartColumn == addr.EndColumn and addr.StartRow == addr.EndRow:
        sheet = model.getSheets().getByIndex(addr.Sheet)
        cursor = sheet.createCursorByRange(rng)
        cursor.collapseToCurrentRegion()
        return cursor
    return rng


def region_from_selection(controller):
    ranges = selected_ranges(controller)
    return expand_if_single(controller.getModel(), ranges[0]) if ranges else None


def range_from_text(model, controller, text, clip=True):
    """Parse 'Sheet2.A1:C9', "'My Sheet'.A1:B2", '$A$1:$B$9' (active sheet) -> range.
    Raises ValueError with a readable message."""
    text = (text or "").strip()
    if not text:
        raise ValueError("empty range")
    sheet_name, ref = split_sheet_ref(text)
    sheets = model.getSheets()
    try:
        sheet = sheets.getByName(sheet_name) if sheet_name else controller.getActiveSheet()
    except Exception:
        raise ValueError("no sheet named %r" % sheet_name)
    try:
        rng = sheet.getCellRangeByName(ref)
        rng.getRangeAddress()
    except Exception:
        raise ValueError("%r is not a valid range" % text)
    return clip_to_used(rng) if clip else rng


def address_text(rng):
    """'$Sheet1.$A$1:$C$9' (re-parsable by range_from_text)."""
    try:
        return rng.getPropertyValue("AbsoluteName")
    except Exception:
        addr = rng.getRangeAddress()
        return "%s%d:%s%d" % (col_name(addr.StartColumn), addr.StartRow + 1,
                              col_name(addr.EndColumn), addr.EndRow + 1)


def used_area(sheet):
    """Range A1:<last used cell> of a sheet (what a lone selected cell means for tools that must
    see blank rows, which a 'current region' would stop at)."""
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    end = cursor.getRangeAddress()
    return sheet.getCellRangeByPosition(0, 0, end.EndColumn, end.EndRow)


def formula_grid(rng):
    """Raw cell contents as text ('' for empty cells, '=A1+1' for formulas)."""
    return [tuple(row) for row in rng.getFormulaArray()]


def cell_range_address(sheet_index, col1, row1, col2, row2):
    import uno
    addr = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
    addr.Sheet, addr.StartColumn, addr.StartRow = sheet_index, col1, row1
    addr.EndColumn, addr.EndRow = col2, row2
    return addr


def data_grid(rng):
    """Range contents as a list of row tuples (numbers as float, text as str)."""
    return [tuple(row) for row in rng.getDataArray()]


def sub_range(sheet, addr, row1, col1, row2, col2):
    """Cell range inside `addr` given 0-based offsets relative to its top-left."""
    return sheet.getCellRangeByPosition(addr.StartColumn + col1, addr.StartRow + row1,
                                        addr.StartColumn + col2, addr.StartRow + row2)


def set_cell_value(sheet, col, row, value):
    """Write one value. Unlike XCellRangeData.setDataArray (NOT undoable in LibreOffice), the
    cell setters are recorded in the undo stack, so use this for edits to existing data."""
    cell = sheet.getCellByPosition(col, row)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        cell.setValue(float(value))
    else:
        cell.setString("" if value is None else str(value))


def fill_cells(sheet, col1, row1, col2, row2, value):
    """Write `value` into every cell of the (absolute) rectangle, undoably."""
    for row in range(row1, row2 + 1):
        for col in range(col1, col2 + 1):
            set_cell_value(sheet, col, row, value)


@contextlib.contextmanager
def undo_context(model, title):
    """Everything inside is ONE undo step."""
    manager = model.getUndoManager()
    manager.enterUndoContext(title)
    try:
        yield
    finally:
        manager.leaveUndoContext()


def add_sheet_after(model, base_name, after_index):
    """Insert a new sheet after `after_index` with a unique name; returns the sheet."""
    sheets = model.getSheets()
    name = unique_name(sheets.getElementNames(), base_name[:31])
    sheets.insertNewByName(name, after_index + 1)
    return sheets.getByName(name)
