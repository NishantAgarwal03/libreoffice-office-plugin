"""Transpose Table: rows become columns and columns become rows.

Engine (default): LibreOffice's own Copy + Paste Special > Transpose (`.uno:Copy`,
`.uno:InsertContents` with Transpose=true), which adjusts formula references and carries over
formats and merged cells. It uses the system clipboard (its content is replaced).
Fallback engine (setting "engine": "cells", or automatically if the clipboard route fails):
values and number formats only, written cell by cell.
Everything is one undo step.
"""

from office_plugin.core.feature import Feature
from office_plugin.services import calc_ranges as R

from . import transpose as T

_TITLE = "Transpose Table"
_MAX_CELLS = 5_000_000
_ALL = 1023            # CellFlags: every kind of content and formatting


class CalcTransposeFeature(Feature):
    def execute(self, command, frame, args):
        docs = self.services.get("doc_utils")
        ui = self.services.get("ui")
        controller = docs.controller_of(frame) or docs.active_controller()
        if controller is None:
            return
        model = controller.getModel()
        ranges = R.selected_ranges(controller)
        if len(ranges) != 1:
            ui.message(_TITLE, "Select a single table first.", frame, "warning")
            return
        rng = R.expand_if_single(model, ranges[0])
        a = rng.getRangeAddress()
        sheet = model.getSheets().getByIndex(a.Sheet)
        src = T.Rect(a.StartRow, a.StartColumn, a.EndRow, a.EndColumn)
        if T.cell_count(src) > _MAX_CELLS:
            ui.message(_TITLE, "The table is too large to transpose (%d cells)." % T.cell_count(src),
                       frame, "warning")
            return
        max_row, max_col = sheet.getRows().getCount() - 1, sheet.getColumns().getCount() - 1

        default_target = sheet.getCellByPosition(min(src.col2 + 2, max_col), src.row1).getPropertyValue("AbsoluteName")
        answer = ui.form(_TITLE, [
            {"type": "label", "text": "%s\n%s" % (R.address_text(rng), T.describe(src))},
            {"type": "dropdown", "name": "dest", "label": "Write the transposed table",
             "items": list(T.DESTINATIONS), "value": T.IN_PLACE},
            {"type": "text", "name": "target", "label": "Cell for 'At cell...'", "value": default_target},
            {"type": "checkbox", "name": "values", "value": False,
             "label": "Convert formulas to values"},
        ], frame)
        if answer is None:
            return
        self._transpose(frame, controller, model, rng, sheet, src, answer, max_row, max_col)

    # ------------------------------------------------------------------------------
    def _transpose(self, frame, controller, model, rng, sheet, src, answer, max_row, max_col):
        ui = self.services.get("ui")
        dest_kind = answer["dest"]
        target_sheet, anchor = sheet, (src.row1, src.col1)
        if dest_kind == T.AT_CELL:
            try:
                tr = R.range_from_text(model, controller, answer["target"], clip=False).getRangeAddress()
            except ValueError as exc:
                ui.message(_TITLE, "Cell: %s" % exc, frame, "warning")
                return
            target_sheet = model.getSheets().getByIndex(tr.Sheet)
            anchor = (tr.StartRow, tr.StartColumn)
        elif dest_kind == T.NEW_SHEET:
            anchor = (0, 0)
        dest = T.transposed_rect(src, *anchor)
        same_sheet = target_sheet.getName() == sheet.getName() and dest_kind != T.NEW_SHEET
        t_max_row, t_max_col = target_sheet.getRows().getCount() - 1, target_sheet.getColumns().getCount() - 1
        if not T.fits(dest, t_max_row, t_max_col):
            ui.message(_TITLE, "The transposed table (%d rows x %d columns) does not fit on the sheet."
                       % (T.rows_of(dest), T.cols_of(dest)), frame, "warning")
            return
        if dest_kind == T.AT_CELL and same_sheet and T.intersects(src, dest):
            ui.message(_TITLE, "The target overlaps the table. Choose another cell or use 'In place'.",
                       frame, "warning")
            return

        # cells outside the table that would be overwritten
        overwritten = 0
        if dest_kind != T.NEW_SHEET:
            for part in (T.rect_difference(dest, src) if same_sheet else [dest]):
                block = target_sheet.getCellRangeByPosition(part.col1, part.row1, part.col2, part.row2)
                overwritten += sum(1 for row in block.getFormulaArray() for text in row if text != "")
        if overwritten and not ui.confirm(
                _TITLE, "%d non-empty cell(s) outside the table will be overwritten.\nContinue?" % overwritten, frame):
            return

        clear_source = dest_kind == T.IN_PLACE
        src_text = R.address_text(rng)       # before editing: live range objects follow the changes
        with R.undo_context(model, _TITLE):
            if dest_kind == T.NEW_SHEET:
                target_sheet = R.add_sheet_after(model, "Transposed", rng.getRangeAddress().Sheet)
            engine = self.option("engine", "clipboard")
            done = False
            if engine == "clipboard":
                try:
                    self._clipboard(controller, rng, target_sheet, anchor, clear_source)
                    done = True
                except Exception:
                    self.log.exception("transpose: clipboard route failed, using the cell fallback")
            if not done:
                self._cells(rng, target_sheet, anchor, src, clear_source)
            block = target_sheet.getCellRangeByPosition(dest.col1, dest.row1, dest.col2, dest.row2)
            converted = self._freeze_formulas(target_sheet, dest, block) if answer["values"] else 0
        controller.setActiveSheet(target_sheet)
        controller.select(block)
        text = "Transposed %s to %s." % (src_text, R.address_text(block))
        if converted:
            text += " %d formula(s) converted to values." % converted
        ui.message(_TITLE, text, frame)

    # ---- engines -------------------------------------------------------------------
    def _dispatch(self, controller, command, **props):
        import uno
        helper = self.services.get("doc_utils").create("com.sun.star.frame.DispatchHelper")
        args = tuple(uno.createUnoStruct("com.sun.star.beans.PropertyValue", k, 0, v, 0)
                     for k, v in props.items())
        helper.executeDispatch(controller.getFrame(), command, "", 0, args)

    def _clipboard(self, controller, rng, target_sheet, anchor, clear_source):
        controller.select(rng)
        self._dispatch(controller, ".uno:Copy")
        if clear_source:
            rng.clearContents(_ALL)         # the clipboard keeps the copy; avoids stale cells
        controller.setActiveSheet(target_sheet)
        controller.select(target_sheet.getCellByPosition(anchor[1], anchor[0]))
        self._dispatch(controller, ".uno:InsertContents", Flags="A", Transpose=True,
                       FormulaCommand=0, SkipEmptyCells=False, AsLink=False, MoveMode=0)

    def _cells(self, rng, target_sheet, anchor, src, clear_source):
        """Fallback: values + number formats, cell by cell (formulas become values)."""
        grid = R.data_grid(rng)
        formats = [[rng.getCellByPosition(c, r).getPropertyValue("NumberFormat")
                    for c in range(T.cols_of(src))] for r in range(T.rows_of(src))]
        if clear_source:
            rng.clearContents(_ALL)
        for r, row in enumerate(grid):
            for c, value in enumerate(row):
                col, row_idx = anchor[1] + r, anchor[0] + c          # (r, c) -> (c, r)
                if value != "":
                    R.set_cell_value(target_sheet, col, row_idx, value)
                if formats[r][c]:
                    target_sheet.getCellByPosition(col, row_idx).setPropertyValue("NumberFormat", formats[r][c])

    @staticmethod
    def _freeze_formulas(target_sheet, dest, block):
        formulas, data = block.getFormulaArray(), block.getDataArray()
        converted = 0
        for r, row in enumerate(formulas):
            for c, text in enumerate(row):
                if text.startswith("="):
                    R.set_cell_value(target_sheet, dest.col1 + c, dest.row1 + r, data[r][c])
                    converted += 1
        return converted
