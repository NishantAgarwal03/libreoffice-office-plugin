"""Compare Lists: union / intersection / differences of two Calc ranges into a new sheet.

Flow: pick List A and List B (click-to-pick wizard, or Ctrl+select both first, or type the
addresses) -> options dialog -> key-column dialog -> counts + confirm -> new sheet.
"""

from office_plugin.core.feature import Feature
from office_plugin.services import calc_ranges as R
from office_plugin.services.pick_ranges import pick_lists
from office_plugin.services.cellkeys import CleanOptions, clean_text, header_guess

from . import setops as S

_TITLE = "Compare Lists"


class CalcSetopsFeature(Feature):
    # ---- entry ---------------------------------------------------------------
    def execute(self, command, frame, args):
        docs = self.services.get("doc_utils")
        ui = self.services.get("ui")
        controller = docs.controller_of(frame) or docs.active_controller()
        if controller is None:
            return
        ranges = R.selected_ranges(controller)
        if len(ranges) >= 2:                       # Ctrl+select of List A and List B
            return self._options(frame, controller, R.address_text(ranges[0]),
                                 R.address_text(ranges[1]))
        first = R.region_from_selection(controller)
        a_text = R.address_text(first) if first is not None else ""
        if not self.option("pick_ranges", True):
            return self._options(frame, controller, a_text, "")

        # Click-to-pick wizard (non-modal, so it reports through a callback); falls back to
        # the dialog with typed addresses when the picker is unavailable.
        pick_lists(ui, controller, a_text, True,
                   lambda a, b: self._guarded(self._options, frame, controller, a, b), _TITLE)

    def _guarded(self, fn, *args):
        try:
            fn(*args)
        except Exception:
            self.log.exception("calc_setops failed")

    # ---- dialogs ---------------------------------------------------------------
    def _options(self, frame, controller, a_text, b_text):
        ui = self.services.get("ui")
        model = controller.getModel()
        prev = {"a": a_text, "b": b_text}
        while True:
            header_default = prev.get("header")
            if header_default is None:
                try:
                    header_default = header_guess(
                        R.data_grid(R.range_from_text(model, controller, prev["a"]))[0])
                except (ValueError, IndexError):
                    header_default = True
            fields = [
                {"type": "label", "text": "Compare two lists and write the result to a NEW sheet.\n"
                                          "Your data is not changed."},
                {"type": "text", "name": "a", "label": "List A (e.g. $Sheet1.$A$1:$C$50)",
                 "value": prev["a"]},
                {"type": "text", "name": "b", "label": "List B (may be on another sheet)",
                 "value": prev["b"]},
                {"type": "checkbox", "name": "header", "label": "Both lists have a header row",
                 "value": header_default},
                {"type": "dropdown", "name": "op", "label": "Operation", "items": list(S.OPERATIONS),
                 "value": prev.get("op", S.INTERSECTION)},
                {"type": "label", "text": "Cleaning (applied when matching values):"},
                {"type": "checkbox", "name": "trim", "value": prev.get("trim", True),
                 "label": "Trim and collapse spaces (incl. non-breaking)"},
                {"type": "checkbox", "name": "ignore_case", "value": prev.get("ignore_case", True),
                 "label": "Ignore upper/lower case"},
                {"type": "checkbox", "name": "numeric", "value": prev.get("numeric", True),
                 "label": "Numbers equal numeric text (101 = \"101\")"},
                {"type": "checkbox", "name": "nonprint", "value": prev.get("nonprint", True),
                 "label": "Ignore invisible characters"},
                {"type": "checkbox", "name": "clean_out", "value": prev.get("clean_out", False),
                 "label": "Also clean text in the result (otherwise keep original values)"},
            ]
            answer = ui.form(_TITLE, fields, frame)
            if answer is None:
                return
            prev = dict(answer)
            try:
                rng_a = R.range_from_text(model, controller, answer["a"])
                rng_b = R.range_from_text(model, controller, answer["b"])
            except ValueError as exc:
                ui.message(_TITLE, "List A/B: %s" % exc, frame, "warning")
                continue
            width_a = rng_a.getRangeAddress().EndColumn - rng_a.getRangeAddress().StartColumn
            width_b = rng_b.getRangeAddress().EndColumn - rng_b.getRangeAddress().StartColumn
            if width_a != width_b:
                ui.message(_TITLE, "Both lists need the same number of columns "
                                   "(A has %d, B has %d)." % (width_a + 1, width_b + 1), frame, "warning")
                continue
            break
        self._run(frame, controller, model, rng_a, rng_b, answer)

    def _run(self, frame, controller, model, rng_a, rng_b, answer):
        ui = self.services.get("ui")
        header = answer["header"]
        grid_a, grid_b = R.data_grid(rng_a), R.data_grid(rng_b)
        addr_a = rng_a.getRangeAddress()
        skip = 1 if header else 0
        rows_a, rows_b = grid_a[skip:], grid_b[skip:]
        if not rows_a and not rows_b:
            ui.message(_TITLE, "Both lists are empty.", frame, "warning")
            return

        labels = R.column_labels(grid_a[0], addr_a.StartColumn, header)
        keys = ui.form(_TITLE, [
            {"type": "label", "text": "Rows match when ALL selected columns match."},
            {"type": "multilist", "name": "cols", "label": "Match on columns", "items": labels,
             "selected": list(range(len(labels))), "rows": min(10, max(3, len(labels)))},
        ], frame)
        if keys is None:
            return
        if not keys["cols"]:
            ui.message(_TITLE, "Select at least one column.", frame, "warning")
            return

        opts = CleanOptions(answer["trim"], answer["ignore_case"], answer["numeric"], answer["nonprint"])
        op = answer["op"]
        result, counts = S.compute(op, rows_a, rows_b, keys["cols"], opts)
        summary = ("List A: %d rows (%d distinct), List B: %d rows (%d distinct)\n"
                   "Found in both: %d\n%s: %d row(s)" % (
                       counts["a_rows"], counts["a_distinct"], counts["b_rows"], counts["b_distinct"],
                       counts["common"], op, counts["result"]))
        if counts["skipped_blank"]:
            summary += "\n(%d row(s) with blank key columns were ignored)" % counts["skipped_blank"]
        if not result:
            ui.message(_TITLE, summary + "\n\nNothing to write.", frame)
            return
        if not ui.confirm(_TITLE, summary + "\n\nWrite the result to a new sheet?", frame):
            return
        name = self._write(controller, model, rng_a, rng_b, header, op, result, answer)
        ui.message(_TITLE, "%d row(s) written to sheet '%s'." % (len(result), name), frame)

    # ---- output ------------------------------------------------------------------
    def _write(self, controller, model, rng_a, rng_b, header, op, result, answer):
        import uno
        addr_a, addr_b = rng_a.getRangeAddress(), rng_b.getRangeAddress()
        ncols = addr_a.EndColumn - addr_a.StartColumn + 1
        skip = 1 if header else 0
        sources = {}

        def src_address(source, first_index, last_index):
            a = sources[source]
            rng = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
            rng.Sheet, rng.StartColumn, rng.EndColumn = a.Sheet, a.StartColumn, a.EndColumn
            rng.StartRow = a.StartRow + skip + first_index
            rng.EndRow = a.StartRow + skip + last_index
            return rng

        with R.undo_context(model, _TITLE):
            sheet = R.add_sheet_after(model, S.SHEET_NAMES[op], addr_a.Sheet)
            sheet_index = list(model.getSheets().getElementNames()).index(sheet.Name)
            # inserting the sheet shifts the index of every later sheet (e.g. List B's):
            # read the live ranges' addresses again
            sources["A"], sources["B"] = rng_a.getRangeAddress(), rng_b.getRangeAddress()
            row = 0
            if header:      # header row of A, copied with its formatting
                self._copy(sheet, sheet_index, row, src_address("A", -1, -1))
                row += 1
            # copy consecutive rows from the same list in one call (keeps number formats/dates)
            run_start = 0
            while run_start < len(result):
                source, first = result[run_start]
                last, run_end = first, run_start
                while (run_end + 1 < len(result) and result[run_end + 1][0] == source
                       and result[run_end + 1][1] == last + 1):
                    run_end += 1
                    last += 1
                self._copy(sheet, sheet_index, row, src_address(source, first, last))
                row += run_end - run_start + 1
                run_start = run_end + 1
            # freeze formulas to values (and optionally clean text) in one pass
            block = sheet.getCellRangeByPosition(0, 0, ncols - 1, row - 1)
            data = block.getDataArray()
            if answer["clean_out"]:
                text_opts = CleanOptions(answer["trim"], False, False, answer["nonprint"])
                data = tuple(tuple(clean_text(v, text_opts) if isinstance(v, str) else v for v in r)
                             for r in data)
            block.setDataArray(data)
            for c in range(ncols):
                sheet.getColumns().getByIndex(c).setPropertyValue("OptimalWidth", True)
            controller.setActiveSheet(sheet)
        return sheet.Name

    @staticmethod
    def _copy(sheet, sheet_index, dest_row, src):
        import uno
        dest = uno.createUnoStruct("com.sun.star.table.CellAddress")
        dest.Sheet, dest.Column, dest.Row = sheet_index, 0, dest_row
        sheet.copyRange(dest, src)
