"""Convert the cells of a selection with a function and deliver the results.

Purpose: shared engine for 'convert each cell' tools (number to words, words to number, units to
         number): plan the conversion, report cells that could not be converted, ask before
         overwriting, write undoably next to the selection or in place.
Inputs:  UNO controller/ranges, an `accepts(data, formula)` predicate and a `convert(data)` function
         (raises ValueError with a readable message).
Outputs: results written to the sheet; a summary message.
Must not: import uno at module level (plan() is pure and unit-tested).
"""

from collections import namedtuple

from office_plugin.services import calc_ranges as R

RIGHT = "Next column to the right"
IN_PLACE = "In place (replace the cells)"
DELIVERIES = (RIGHT, IN_PLACE)

Item = namedtuple("Item", "row col result error")      # 0-based offsets inside the range
Plan = namedtuple("Plan", "items skipped")             # skipped = cells that were not eligible


def plan(data_grid, formula_grid, accepts, convert, skip=frozenset()):
    """Pure: -> Plan(items, skipped). Empty cells are ignored silently; ineligible cells counted.
    `skip` = {(row, col)} offsets that must not be converted (counted as skipped)."""
    items, skipped = [], 0
    for r, row in enumerate(data_grid):
        for c, data in enumerate(row):
            if (r, c) in skip and not (data in ("", None)):
                skipped += 1
                continue
            formula = formula_grid[r][c] if r < len(formula_grid) and c < len(formula_grid[r]) else ""
            if data in ("", None) and formula in ("", None):
                continue
            if not accepts(data, formula):
                skipped += 1
                continue
            try:
                items.append(Item(r, c, convert(data), None))
            except (ValueError, ArithmeticError) as exc:
                items.append(Item(r, c, None, str(exc)))
    return Plan(items, skipped)


def _address(addr, item, delivery):
    col = addr.StartColumn + item.col + ((addr.EndColumn - addr.StartColumn + 1) if delivery == RIGHT else 0)
    return col, addr.StartRow + item.row


# number-format kinds a number-to-words style conversion must not touch (com.sun.star.util.NumberFormat)
_DATE, _TIME, _LOGICAL, _TEXT = 2, 4, 1024, 256


def special_format_cells(model, rng):
    """{(row, col)} offsets of date/time/boolean/text-formatted cells in the range."""
    addr = rng.getRangeAddress()
    formats, kinds, found = model.getNumberFormats(), {}, set()
    for flags in (1, 2, 16):                                   # value, date/time, formula cells
        hits = rng.queryContentCells(flags)
        if hits.getCount() == 0:
            continue
        cells = hits.getCells().createEnumeration()
        while cells.hasMoreElements():
            cell = cells.nextElement()
            key = cell.getPropertyValue("NumberFormat")
            if key not in kinds:
                kinds[key] = formats.getByKey(key).getPropertyValue("Type")
            if kinds[key] & (_DATE | _TIME | _LOGICAL | _TEXT):
                where = cell.getCellAddress()
                found.add((where.Row - addr.StartRow, where.Column - addr.StartColumn))
    return found


def run(ui, frame, controller, title, accepts, convert, delivery, nothing_text, done_noun,
        skip_special_formats=False):
    """Convert every selected range. Returns the number of converted cells (0 when cancelled).
    skip_special_formats: leave date / time / boolean / text-formatted cells alone."""
    model = controller.getModel()
    ranges = R.selected_ranges(controller)
    if not ranges:
        ui.message(title, "Select the cells to convert first.", frame, "warning")
        return 0

    jobs = []                                              # (sheet, addr, plan)
    for rng in ranges:
        addr = rng.getRangeAddress()
        sheet = model.getSheets().getByIndex(addr.Sheet)
        skip = special_format_cells(model, rng) if skip_special_formats else frozenset()
        jobs.append((sheet, addr, plan(R.data_grid(rng), R.formula_grid(rng), accepts, convert, skip)))
    converted = [(sheet, addr, it) for sheet, addr, p in jobs for it in p.items if it.error is None]
    errors = [(sheet, addr, it) for sheet, addr, p in jobs for it in p.items if it.error is not None]
    skipped = sum(p.skipped for _, _, p in jobs)

    if not converted:
        text = nothing_text
        if errors:
            text += "\n" + _error_lines(errors)
        ui.message(title, text, frame, "warning" if errors else "info")
        return 0

    # ---- what will be overwritten --------------------------------------------------------------
    if delivery == IN_PLACE:
        question = "Replace %d cell(s) in place?" % len(converted)
    else:
        clash = 0
        for sheet, addr, it in converted:
            col, row = _address(addr, it, delivery)
            if col >= sheet.getColumns().getCount():
                ui.message(title, "There is no room to the right of the selection.", frame, "warning")
                return 0
            if sheet.getCellByPosition(col, row).getFormula() != "":
                clash += 1
        question = ("%d non-empty cell(s) to the right will be overwritten.\nContinue?" % clash) if clash else None
    if question and not ui.confirm(title, question, frame):
        return 0

    with R.undo_context(model, title):
        for sheet, addr, it in converted:
            col, row = _address(addr, it, delivery)
            R.set_cell_value(sheet, col, row, it.result)
    summary = "%d %s." % (len(converted), done_noun)
    if skipped:
        summary += "\n%d cell(s) skipped (not applicable)." % skipped
    if errors:
        summary += "\n%d could not be converted:\n%s" % (len(errors), _error_lines(errors))
    ui.message(title, summary, frame, "warning" if errors else "info")
    return len(converted)


def _error_lines(errors, limit=8):
    lines = []
    for sheet, addr, it in errors[:limit]:
        lines.append("  %s%d: %s" % (R.col_name(addr.StartColumn + it.col), addr.StartRow + it.row + 1, it.error))
    if len(errors) > limit:
        lines.append("  ... and %d more" % (len(errors) - limit))
    return "\n".join(lines)
