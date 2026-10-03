"""Fill Blanks: copy the value above (or left) into empty cells of the selection.

Values only: a formula cell is copied as its displayed value. Only blank cells are written, so
formats and every other cell stay as they are. One undo step (cell setters are undoable;
setDataArray is not, so it is never used on existing data).
"""

from office_plugin.core.feature import Feature
from office_plugin.services import calc_ranges as R
from office_plugin.services.cellkeys import is_blank

from .fill import count_cells, plan_fill

_TITLE = "Fill Blanks"
_MAX_MERGE_SCAN = 20000     # blank cells examined for merged areas before we give up


class CalcFillBlanksFeature(Feature):
    def execute(self, command, frame, args):
        docs = self.services.get("doc_utils")
        ui = self.services.get("ui")
        controller = docs.controller_of(frame) or docs.active_controller()
        if controller is None:
            return
        model = controller.getModel()
        ranges = [R.expand_if_single(model, r) for r in R.selected_ranges(controller)]
        if not ranges:
            ui.message(_TITLE, "Select a range of cells first.", frame, "warning")
            return
        answer = ui.form(_TITLE, [
            {"type": "label", "text": "Fill empty cells in the selection with the value above\n"
                                      "(or to the left). Only blank cells are changed."},
            {"type": "dropdown", "name": "direction", "label": "Direction",
             "items": ["Down", "Right"], "value": "Down"},
            {"type": "checkbox", "name": "whitespace", "value": False,
             "label": "Cells containing only spaces count as blank"},
            {"type": "checkbox", "name": "unmerge", "value": True,
             "label": "Unmerge merged cells and fill them with their value"},
        ], frame)
        if answer is None:
            return
        direction = "down" if answer["direction"] == "Down" else "right"

        total = merged = 0
        with R.undo_context(model, _TITLE):
            for rng in ranges:
                filled, unmerged = self._fill_range(model, rng, direction, answer)
                total += filled
                merged += unmerged
        if total == 0:
            ui.message(_TITLE, "No blank cells with a value %s to fill."
                       % ("above" if direction == "down" else "to the left"), frame)
            return
        text = "Filled %d cell(s)." % total
        if merged:
            text += " Unmerged %d merged area(s)." % merged
        ui.message(_TITLE, text, frame)

    # ---- one range ---------------------------------------------------------------
    def _fill_range(self, model, rng, direction, opts):
        addr = rng.getRangeAddress()
        sheet = model.getSheets().getByIndex(addr.Sheet)
        ws = opts["whitespace"]

        def blank(value):
            return is_blank(value, whitespace=ws)

        filled = merged = 0
        if opts["unmerge"]:
            filled, merged = self._unmerge_and_fill(sheet, rng, addr, blank)
        runs = plan_fill(R.data_grid(rng), direction, blank)
        for r1, c1, r2, c2, value in runs:      # cell setters, not setDataArray: those are undoable
            R.fill_cells(sheet, addr.StartColumn + c1, addr.StartRow + r1,
                         addr.StartColumn + c2, addr.StartRow + r2, value)
        return filled + count_cells(runs), merged

    def _unmerge_and_fill(self, sheet, rng, addr, blank):
        """Real merged areas: unmerge each one that touches a blank cell of the selection and
        fill the part inside the selection with the area's top-left value."""
        grid = R.data_grid(rng)
        blanks = [(r, c) for r, row in enumerate(grid) for c, v in enumerate(row) if blank(v)]
        if len(blanks) > _MAX_MERGE_SCAN:
            self.log.info("calc_fill_blanks: %d blank cells, skipping merged-area scan", len(blanks))
            return 0, 0
        done, filled, merged = [], 0, 0
        for r, c in blanks:
            row, col = addr.StartRow + r, addr.StartColumn + c
            if any(a[0] <= row <= a[2] and a[1] <= col <= a[3] for a in done):
                continue
            cell = sheet.getCellByPosition(col, row)
            cursor = sheet.createCursorByRange(cell)
            cursor.collapseToMergedArea()
            area = cursor.getRangeAddress()
            if area.StartRow == area.EndRow and area.StartColumn == area.EndColumn:
                continue                                    # not merged
            done.append((area.StartRow, area.StartColumn, area.EndRow, area.EndColumn))
            top_left = sheet.getCellByPosition(area.StartColumn, area.StartRow)
            value = top_left.getDataArray()[0][0]
            area_range = sheet.getCellRangeByPosition(area.StartColumn, area.StartRow,
                                                      area.EndColumn, area.EndRow)
            area_range.merge(False)
            merged += 1
            if blank(value):
                continue
            # the part of the area inside the selection
            r1, r2 = max(area.StartRow, addr.StartRow), min(area.EndRow, addr.EndRow)
            c1, c2 = max(area.StartColumn, addr.StartColumn), min(area.EndColumn, addr.EndColumn)
            if r1 > r2 or c1 > c2:
                continue
            R.fill_cells(sheet, c1, r1, c2, r2, value)
            inside_top_left = r1 == area.StartRow and c1 == area.StartColumn
            filled += (r2 - r1 + 1) * (c2 - c1 + 1) - (1 if inside_top_left else 0)
        return filled, merged
