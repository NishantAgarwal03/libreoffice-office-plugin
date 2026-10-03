"""Remove Duplicates for Calc."""

from office_plugin.core.feature import Feature
from office_plugin.services import calc_ranges as R
from office_plugin.services.cellkeys import CleanOptions

from .dedupe import col_name, find_duplicate_rows, group_runs

_TITLE = "Remove Duplicates"


class CalcDedupeFeature(Feature):
    def execute(self, command, frame, args):
        import uno
        from com.sun.star.sheet.CellDeleteMode import UP

        docs = self.services.get("doc_utils")
        ui = self.services.get("ui")
        controller = docs.controller_of(frame) or docs.active_controller()
        if controller is None:
            return
        model = controller.getModel()
        ranges = R.selected_ranges(controller)           # clipped to the used area: a whole column is cheap
        if len(ranges) != 1:
            ui.message(_TITLE, "Select a single range of cells first.", frame, "warning")
            return

        rng = R.expand_if_single(model, ranges[0])       # like Excel: expand a lone cell
        addr = rng.getRangeAddress()
        sheet = model.getSheets().getByIndex(addr.Sheet)
        data = rng.getDataArray()
        if len(data) < 2:
            ui.message(_TITLE, "The range needs at least two rows.", frame, "warning")
            return

        first = data[0]
        header_guess = any(isinstance(v, str) and v for v in first) and \
            all(isinstance(v, str) for v in first if v != "")
        labels = []
        for i, value in enumerate(first):
            letter = col_name(addr.StartColumn + i)
            labels.append("%s - %s" % (letter, value) if header_guess and value != ""
                          else "Column %s" % letter)
        answer = ui.choose_columns(
            _TITLE, "Rows are duplicates when the selected columns all match.\n"
                    "The first occurrence is kept.",
            labels,
            [("header", "My data has a header row", header_guess),
             ("trim", "Trim and collapse spaces (incl. non-breaking)", True),
             ("ignore_case", "Ignore upper/lower case (Apple = apple)", True),
             ("numeric", "Numbers equal numeric text (101 = \"101\")", True),
             ("nonprint", "Ignore invisible characters", True)],
            frame)
        if answer is None:
            return
        selected, opts = answer
        if not selected:
            ui.message(_TITLE, "Select at least one column.", frame, "warning")
            return

        clean = CleanOptions(opts["trim"], opts["ignore_case"], opts["numeric"], opts["nonprint"])
        dupes = find_duplicate_rows(data, selected, has_header=opts["header"], opts=clean)
        if not dupes:
            ui.message(_TITLE, "No duplicate rows found.", frame)
            return

        undo = model.getUndoManager()
        undo.enterUndoContext(_TITLE)
        try:
            for first_row, last_row in reversed(group_runs(dupes)):   # bottom-up keeps indexes valid
                target = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
                target.Sheet = addr.Sheet
                target.StartColumn, target.EndColumn = addr.StartColumn, addr.EndColumn
                target.StartRow = addr.StartRow + first_row
                target.EndRow = addr.StartRow + last_row
                sheet.removeRange(target, UP)
        finally:
            undo.leaveUndoContext()
        remaining = len(data) - len(dupes) - (1 if opts["header"] else 0)
        ui.message(_TITLE, "%d duplicate row(s) removed; %d unique row(s) remain."
                   % (len(dupes), remaining), frame)
