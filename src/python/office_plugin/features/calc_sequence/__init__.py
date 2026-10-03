"""Sequence Numbers: fill a series (start / step / end / count) into the selection.

Numbers start at the selection's top-left cell and run down or across. The series may be longer
than the selection; non-empty cells beyond the selection are only overwritten after a confirm.
Cell setters are used (undoable), so Ctrl+Z reverts everything in one step.
"""

from office_plugin.core.feature import Feature
from office_plugin.services import calc_ranges as R

from . import sequence as S

_TITLE = "Sequence Numbers"
_DOWN, _ACROSS = "Down", "Across"


class CalcSequenceFeature(Feature):
    def execute(self, command, frame, args):
        docs = self.services.get("doc_utils")
        ui = self.services.get("ui")
        controller = docs.controller_of(frame) or docs.active_controller()
        if controller is None:
            return
        model = controller.getModel()
        ranges = R.selected_ranges(controller, clip=False)      # writes INTO the selection: keep its exact size
        if len(ranges) != 1:
            ui.message(_TITLE, "Select the cells to number first.", frame, "warning")
            return
        rng = ranges[0]
        a = rng.getRangeAddress()
        n_rows, n_cols = a.EndRow - a.StartRow + 1, a.EndColumn - a.StartColumn + 1

        prev = {"start": "1", "step": "1", "end": "", "count": "",
                "direction": _DOWN if n_rows >= n_cols else _ACROSS}
        while True:
            answer = ui.form(_TITLE, [
                {"type": "label", "text": "Numbers start at %s and run down or across.\n"
                                          "Fill any combination of Step, End and Count; leave the\n"
                                          "rest empty. With only Start and Step the selection is filled."
                                          % R.address_text(rng).split(":")[0]},
                {"type": "text", "name": "start", "label": "Start", "value": prev["start"]},
                {"type": "text", "name": "step", "label": "Step (negative = descending)", "value": prev["step"]},
                {"type": "text", "name": "end", "label": "End (optional)", "value": prev["end"]},
                {"type": "text", "name": "count", "label": "Count (optional)", "value": prev["count"]},
                {"type": "dropdown", "name": "direction", "label": "Direction",
                 "items": [_DOWN, _ACROSS], "value": prev["direction"]},
            ], frame)
            if answer is None:
                return
            prev = dict(answer)
            horizontal = answer["direction"] == _ACROSS
            try:
                values = S.build_series(
                    S.parse_number(answer["start"], "Start"), S.parse_number(answer["step"], "Step"),
                    S.parse_number(answer["end"], "End"), S.parse_number(answer["count"], "Count"),
                    default_count=n_cols if horizontal else n_rows)
            except ValueError as exc:
                ui.message(_TITLE, str(exc), frame, "warning")
                continue
            break
        self._write(frame, controller, model, rng, a, values, horizontal, n_rows, n_cols)

    def _write(self, frame, controller, model, rng, a, values, horizontal, n_rows, n_cols):
        ui = self.services.get("ui")
        sheet = model.getSheets().getByIndex(a.Sheet)
        n = len(values)
        col0, row0 = a.StartColumn, a.StartRow
        last_col, last_row = (col0 + n - 1, row0) if horizontal else (col0, row0 + n - 1)
        if last_row >= sheet.getRows().getCount() or last_col >= sheet.getColumns().getCount():
            ui.message(_TITLE, "The series (%d values) does not fit on the sheet." % n, frame, "warning")
            return

        # non-empty cells BEYOND the selection that the series would overwrite
        sel_len = n_cols if horizontal else n_rows
        if n > sel_len:
            extra = (sheet.getCellRangeByPosition(col0 + sel_len, row0, last_col, last_row) if horizontal
                     else sheet.getCellRangeByPosition(col0, row0 + sel_len, last_col, last_row))
            clash = sum(1 for row in extra.getFormulaArray() for text in row if text != "")
            if clash and not ui.confirm(
                    _TITLE, "%d non-empty cell(s) beyond the selection will be overwritten.\nContinue?" % clash,
                    frame):
                return

        with R.undo_context(model, _TITLE):
            for i, value in enumerate(values):
                col, row = (col0 + i, row0) if horizontal else (col0, row0 + i)
                R.set_cell_value(sheet, col, row, value)
        result = sheet.getCellRangeByPosition(col0, row0, last_col, last_row)
        controller.select(result)
        ui.message(_TITLE, "Filled %d cell(s) in %s:\n%s" % (n, R.address_text(result), S.preview(values)), frame)
