"""Find Missing Sequence Numbers: insert a placeholder row for every missing number.

Flow: choose the key column / step / direction -> gaps are computed (pure, lazily counted) ->
confirm with the gap list -> placeholder rows inserted bottom-up (so earlier addresses stay valid),
the missing number written into the key column, the new rows coloured from the plugin palette
(Clear Highlight removes them). Rows are inserted inside the selection's columns by default
(cells below shift down) or as whole sheet rows. One undo step.
"""

from office_plugin.core.feature import Feature
from office_plugin.services import calc_ranges as R
from office_plugin.services.cellkeys import header_guess
from office_plugin.services.palette import PALETTE

from . import missing as M

_TITLE = "Find Missing Numbers"
_ASC, _DESC = "Ascending", "Descending"
_NO_COLOUR = "No colour"
_CONFIRM_ABOVE = 10_000
_REFUSE_ABOVE = 100_000


class CalcMissingSeqFeature(Feature):
    def execute(self, command, frame, args):
        from com.sun.star.sheet.CellInsertMode import DOWN

        docs = self.services.get("doc_utils")
        ui = self.services.get("ui")
        controller = docs.controller_of(frame) or docs.active_controller()
        if controller is None:
            return
        model = controller.getModel()
        ranges = R.selected_ranges(controller)
        if len(ranges) != 1:
            ui.message(_TITLE, "Select the table (or the numbered column) first.", frame, "warning")
            return
        rng = R.expand_if_single(model, ranges[0])
        a = rng.getRangeAddress()
        sheet = model.getSheets().getByIndex(a.Sheet)
        grid = R.data_grid(rng)
        header = header_guess(grid[0]) if grid else False

        prev = {"header": header, "direction": _ASC, "step": "1", "whole": False, "colour": "Yellow"}
        while True:
            labels = R.column_labels(grid[0], a.StartColumn, prev["header"])
            answer = ui.form(_TITLE, [
                {"type": "label", "text": "Insert a placeholder row for each missing number in %s." % R.address_text(rng)},
                {"type": "checkbox", "name": "header", "value": prev["header"], "label": "The range has a header row"},
                {"type": "dropdown", "name": "key", "label": "Numbered column", "items": labels,
                 "value": prev.get("key", labels[0])},
                {"type": "dropdown", "name": "direction", "label": "Order of the numbers",
                 "items": [_ASC, _DESC], "value": prev["direction"]},
                {"type": "text", "name": "step", "label": "Expected step between numbers", "value": prev["step"]},
                {"type": "dropdown", "name": "colour", "label": "Colour of the inserted rows",
                 "items": [_NO_COLOUR] + list(PALETTE), "value": prev["colour"]},
                {"type": "checkbox", "name": "whole", "value": prev["whole"],
                 "label": "Insert whole sheet rows (otherwise only inside the selection's columns)"},
            ], frame)
            if answer is None:
                return
            prev = dict(answer)
            try:
                step = float(answer["step"].strip().replace(",", "."))
                result = self._scan(grid, labels, answer, step)
            except ValueError as exc:
                ui.message(_TITLE, str(exc), frame, "warning")
                continue
            break

        skip = 1 if answer["header"] else 0
        key_col = labels.index(answer["key"]) if answer["key"] in labels else 0
        summary = self._summary(result)
        if result.total_missing == 0:
            ui.message(_TITLE, "No missing numbers found.\n" + summary, frame)
            return
        if result.total_missing > _REFUSE_ABOVE:
            ui.message(_TITLE, "%d numbers are missing - too many to insert (limit %d).\n%s"
                       % (result.total_missing, _REFUSE_ABOVE, M.describe(result)), frame, "warning")
            return
        question = "%d missing number(s) in %d gap(s):\n%s\n\n%s\n\nInsert %d placeholder row(s)?" % (
            result.total_missing, len(result.gaps), M.describe(result), summary, result.total_missing)
        if result.total_missing > _CONFIRM_ABOVE:
            question = "That is a LOT of rows (%d).\n\n%s" % (result.total_missing, question)
        if not ui.confirm(_TITLE, question, frame):
            return

        colour = None if answer["colour"] == _NO_COLOUR else PALETTE[answer["colour"]]
        with R.undo_context(model, _TITLE):
            for gap in reversed(result.gaps):                    # bottom-up keeps earlier rows valid
                self._insert(sheet, a, skip + gap.before_row, gap, key_col, answer["whole"], colour, DOWN)
        ui.message(_TITLE, "Inserted %d placeholder row(s) for %d gap(s):\n%s"
                   % (result.total_missing, len(result.gaps), M.describe(result)), frame)

    # ---- helpers -------------------------------------------------------------------
    @staticmethod
    def _scan(grid, labels, answer, step):
        skip = 1 if answer["header"] else 0
        key_col = labels.index(answer["key"]) if answer["key"] in labels else 0
        column = [row[key_col] if key_col < len(row) else "" for row in grid[skip:]]
        return M.find_gaps(column, step, answer["direction"] == _DESC)

    @staticmethod
    def _summary(result):
        parts = []
        if result.duplicates:
            parts.append("%d duplicate(s)" % result.duplicates)
        if result.out_of_order:
            parts.append("%d out of order" % result.out_of_order)
        if result.off_step:
            parts.append("%d off the step" % result.off_step)
        if result.skipped:
            parts.append("%d unreadable value(s) skipped" % result.skipped)
        return ("Also noticed (left unchanged): " + ", ".join(parts)) if parts else "No duplicates or out-of-order numbers."

    @staticmethod
    def _insert(sheet, a, row_offset, gap, key_col, whole, colour, DOWN):
        row = a.StartRow + row_offset
        if whole:
            sheet.getRows().insertByIndex(row, gap.count)
        else:
            sheet.insertCells(R.cell_range_address(a.Sheet, a.StartColumn, row, a.EndColumn,
                                                   row + gap.count - 1), DOWN)
        for k, number in enumerate(M.gap_numbers(gap)):
            R.set_cell_value(sheet, a.StartColumn + key_col, row + k, M.format_number(gap.series, number))
        if colour is not None:
            sheet.getCellRangeByPosition(a.StartColumn, row, a.EndColumn, row + gap.count - 1) \
                .setPropertyValue("CellBackColor", colour)
