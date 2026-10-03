"""Remove Blanks: delete blank rows, columns or cells of the selection.

A cell is blank when its raw content is empty (a formula that returns "" is NOT blank, so no
formula is ever deleted by mistake); optionally whitespace-only text counts too.
Deleting shifts the cells after the selection (same columns for rows/cells-up, same rows for
columns/cells-left), like Excel's Delete ▸ Shift cells. One undo step.
"""

from office_plugin.core.feature import Feature
from office_plugin.services import calc_ranges as R
from office_plugin.services.cellkeys import is_blank

from . import blanks as B

_TITLE = "Remove Blanks"
_MODES = ["Blank rows", "Blank columns", "Blank cells"]


class CalcRemoveBlanksFeature(Feature):
    def execute(self, command, frame, args):
        from com.sun.star.sheet.CellDeleteMode import LEFT, UP

        docs = self.services.get("doc_utils")
        ui = self.services.get("ui")
        controller = docs.controller_of(frame) or docs.active_controller()
        if controller is None:
            return
        model = controller.getModel()
        ranges = R.selected_ranges(controller)
        if len(ranges) != 1:
            ui.message(_TITLE, "Select a single range of cells first.", frame, "warning")
            return
        rng = ranges[0]
        addr = rng.getRangeAddress()
        sheet = model.getSheets().getByIndex(addr.Sheet)
        if addr.StartColumn == addr.EndColumn and addr.StartRow == addr.EndRow:
            rng = R.used_area(sheet)                    # lone cell -> everything used on the sheet
            addr = rng.getRangeAddress()

        answer = ui.form(_TITLE, [
            {"type": "label", "text": "Remove blank rows, columns or cells inside\n%s" % R.address_text(rng)},
            {"type": "dropdown", "name": "mode", "label": "Remove", "items": _MODES, "value": _MODES[0]},
            {"type": "dropdown", "name": "shift", "label": "For blank cells, shift remaining cells",
             "items": ["Up", "Left"], "value": "Up"},
            {"type": "checkbox", "name": "whitespace", "value": False,
             "label": "Cells containing only spaces count as blank"},
            {"type": "checkbox", "name": "whole", "value": False,
             "label": "Delete whole sheet rows/columns (also removes data outside the selection)"},
        ], frame)
        if answer is None:
            return

        ws = answer["whitespace"]
        mask = [[is_blank(text, whitespace=ws) for text in row] for row in R.formula_grid(rng)]
        mode = answer["mode"]
        if mode == _MODES[0]:
            kind, items = "row", B.blank_rows(mask)
        elif mode == _MODES[1]:
            kind, items = "column", B.blank_cols(mask)
        else:
            axis = "up" if answer["shift"] == "Up" else "left"
            kind, items = "cell", B.blank_cell_runs(mask, axis)
        count = B.count_blank_cells(mask) if kind == "cell" else len(items)
        if count == 0:
            ui.message(_TITLE, "No blank %ss found in %s." % (kind, R.address_text(rng)), frame)
            return

        whole = answer["whole"] and kind != "cell"
        if kind == "cell":
            effect = "The remaining cells shift %s." % answer["shift"].lower()
        elif whole:
            effect = "Whole sheet %ss are deleted, including data outside the selection." % kind
        else:
            effect = "Cells after the selection move %s." % ("up" if kind == "row" else "left")
        if not ui.confirm(_TITLE, "Remove %d blank %s(s) from %s?\n%s" % (
                count, kind, R.address_text(rng), effect), frame):
            return

        with R.undo_context(model, _TITLE):
            if kind == "row":
                for a, b in reversed(B.runs(items)):                 # bottom-up keeps indexes valid
                    if whole:
                        sheet.getRows().removeByIndex(addr.StartRow + a, b - a + 1)
                    else:
                        sheet.removeRange(R.cell_range_address(
                            addr.Sheet, addr.StartColumn, addr.StartRow + a,
                            addr.EndColumn, addr.StartRow + b), UP)
            elif kind == "column":
                for a, b in reversed(B.runs(items)):                 # right-to-left
                    if whole:
                        sheet.getColumns().removeByIndex(addr.StartColumn + a, b - a + 1)
                    else:
                        sheet.removeRange(R.cell_range_address(
                            addr.Sheet, addr.StartColumn + a, addr.StartRow,
                            addr.StartColumn + b, addr.EndRow), LEFT)
            elif answer["shift"] == "Up":
                for col, a, b in reversed(items):                    # per column, bottom-up
                    sheet.removeRange(R.cell_range_address(
                        addr.Sheet, addr.StartColumn + col, addr.StartRow + a,
                        addr.StartColumn + col, addr.StartRow + b), UP)
            else:
                for row, a, b in reversed(items):                    # per row, right-to-left
                    sheet.removeRange(R.cell_range_address(
                        addr.Sheet, addr.StartColumn + a, addr.StartRow + row,
                        addr.StartColumn + b, addr.StartRow + row), LEFT)
        ui.message(_TITLE, "Removed %d blank %s(s)." % (count, kind), frame)
