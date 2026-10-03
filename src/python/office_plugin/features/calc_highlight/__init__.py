"""Highlight duplicates / differences by colouring cells (CellBackColor); Clear removes only
the plugin's own palette colours so colours the user applied survive.

Flow (Run): mode + colour -> pick list A (and B) -> options dialog -> key columns -> paint.
Cell attribute changes are undoable, so Ctrl+Z reverts a whole highlight in one step.
"""

from office_plugin.core.feature import Feature
from office_plugin.services import calc_ranges as R
from office_plugin.services.cellkeys import CleanOptions, header_guess
from office_plugin.services.pick_ranges import pick_lists

from . import marks as M

_TITLE = "Highlight"
_AUTO = "Auto (suggested for the mode)"
_NO_FILL = -1
_MAX_CLEAR_CELLS = 2_000_000


class CalcHighlightFeature(Feature):
    # ---- entry ---------------------------------------------------------------
    def execute(self, command, frame, args):
        docs = self.services.get("doc_utils")
        controller = docs.controller_of(frame) or docs.active_controller()
        if controller is None:
            return
        if command == "Clear":
            return self._clear(frame, controller)
        ui = self.services.get("ui")
        step1 = ui.form(_TITLE, [
            {"type": "label", "text": "What should be highlighted?"},
            {"type": "dropdown", "name": "mode", "label": "Mode", "items": list(M.MODES),
             "value": M.DUP_ALL},
            {"type": "dropdown", "name": "colour", "label": "Colour",
             "items": [_AUTO] + list(M.PALETTE), "value": _AUTO},
        ], frame)
        if step1 is None:
            return
        mode = step1["mode"]
        colour = M.SUGGESTED_COLOUR[mode] if step1["colour"] == _AUTO else step1["colour"]
        want_b = mode in M.NEEDS_B

        ranges = R.selected_ranges(controller)
        if want_b and len(ranges) >= 2:                     # Ctrl+select A and B (same sheet)
            return self._options(frame, controller, mode, colour,
                                 R.address_text(ranges[0]), R.address_text(ranges[1]))
        first = R.region_from_selection(controller)
        a_text = R.address_text(first) if first is not None else ""
        if not self.option("pick_ranges", True):
            return self._options(frame, controller, mode, colour, a_text, "")
        pick_lists(ui, controller, a_text, want_b,
                   lambda a, b: self._guarded(self._options, frame, controller, mode, colour, a, b),
                   _TITLE)

    def _guarded(self, fn, *args):
        try:
            fn(*args)
        except Exception:
            self.log.exception("calc_highlight failed")

    # ---- options -----------------------------------------------------------------
    def _options(self, frame, controller, mode, colour, a_text, b_text):
        ui = self.services.get("ui")
        model = controller.getModel()
        want_b = mode in M.NEEDS_B
        cell_mode = mode == M.CELL_DIFF
        prev = {"a": a_text, "b": b_text}
        while True:
            header_default = prev.get("header")
            if header_default is None:
                try:
                    header_default = header_guess(R.data_grid(R.range_from_text(model, controller, prev["a"]))[0])
                except (ValueError, IndexError):
                    header_default = True
            fields = [{"type": "label", "text": "%s  -  colour: %s" % (mode, colour)},
                      {"type": "text", "name": "a", "label": "List A" if want_b else "List",
                       "value": prev["a"]}]
            if want_b:
                fields.append({"type": "text", "name": "b", "label": "List B", "value": prev["b"]})
            if not cell_mode:
                fields.append({"type": "checkbox", "name": "header", "value": header_default,
                               "label": "The list(s) have a header row (not highlighted)"})
                fields.append({"type": "checkbox", "name": "key_only", "value": prev.get("key_only", False),
                               "label": "Colour only the matching columns (not the whole row)"})
            if mode in (M.IN_BOTH, M.ONLY_IN_A):
                fields.append({"type": "checkbox", "name": "both", "value": prev.get("both", False),
                               "label": "Also colour the rows in B that are %s A" % (
                                   "in" if mode == M.IN_BOTH else "not in")})
            if cell_mode:
                fields.append({"type": "checkbox", "name": "both", "value": prev.get("both", False),
                               "label": "Also colour the differing cells in B"})
            fields += [
                {"type": "label", "text": "Cleaning (applied when comparing):"},
                {"type": "checkbox", "name": "trim", "value": prev.get("trim", True),
                 "label": "Trim and collapse spaces (incl. non-breaking)"},
                {"type": "checkbox", "name": "ignore_case", "value": prev.get("ignore_case", True),
                 "label": "Ignore upper/lower case"},
                {"type": "checkbox", "name": "numeric", "value": prev.get("numeric", True),
                 "label": "Numbers equal numeric text (101 = \"101\")"},
                {"type": "checkbox", "name": "nonprint", "value": prev.get("nonprint", True),
                 "label": "Ignore invisible characters"},
            ]
            answer = ui.form(_TITLE, fields, frame)
            if answer is None:
                return
            prev = dict(answer)
            try:
                rng_a = R.range_from_text(model, controller, answer["a"])
                rng_b = R.range_from_text(model, controller, answer["b"]) if want_b else None
            except ValueError as exc:
                ui.message(_TITLE, "List: %s" % exc, frame, "warning")
                continue
            if rng_b is not None:
                aa, ab = rng_a.getRangeAddress(), rng_b.getRangeAddress()
                same_width = aa.EndColumn - aa.StartColumn == ab.EndColumn - ab.StartColumn
                same_height = aa.EndRow - aa.StartRow == ab.EndRow - ab.StartRow
                if not same_width or (cell_mode and not same_height):
                    what = "the same size" if cell_mode else "the same number of columns"
                    ui.message(_TITLE, "The two ranges must have %s." % what, frame, "warning")
                    continue
            break
        self._apply(frame, controller, model, mode, colour, rng_a, rng_b, answer)

    # ---- run ---------------------------------------------------------------------
    def _apply(self, frame, controller, model, mode, colour, rng_a, rng_b, answer):
        ui = self.services.get("ui")
        opts = CleanOptions(answer["trim"], answer["ignore_case"], answer["numeric"], answer["nonprint"])
        color_value = M.PALETTE[colour]
        addr_a = rng_a.getRangeAddress()
        grid_a = R.data_grid(rng_a)
        grid_b = R.data_grid(rng_b) if rng_b is not None else None
        both = answer.get("both", False)

        if mode == M.CELL_DIFF:
            try:
                diffs = M.cell_diffs(grid_a, grid_b, opts)
            except ValueError as exc:
                ui.message(_TITLE, str(exc), frame, "warning")
                return
            cells_a = {(r, c) for r, c in diffs}
            cells_b = set(cells_a) if both else set()
            skip = 0
        else:
            header = answer["header"]
            skip = 1 if header else 0
            labels_src = grid_a[0] if grid_a else ()
            keys = ui.form(_TITLE, [
                {"type": "label", "text": "Rows match when ALL selected columns match."},
                {"type": "multilist", "name": "cols", "label": "Match on columns",
                 "items": R.column_labels(labels_src, addr_a.StartColumn, header),
                 "selected": list(range(len(labels_src))), "rows": min(10, max(3, len(labels_src)))},
            ], frame)
            if keys is None:
                return
            if not keys["cols"]:
                ui.message(_TITLE, "Select at least one column.", frame, "warning")
                return
            cols = keys["cols"]
            rows_a = grid_a[skip:]
            idx_a, idx_b = self._rows(mode, rows_a, grid_b[skip:] if grid_b else [], cols, opts)
            width = len(grid_a[0]) if grid_a else 0
            target = cols if answer["key_only"] else range(width)
            cells_a = {(r + skip, c) for r in idx_a for c in target}
            cells_b = {(r + skip, c) for r in idx_b for c in target} if both else set()

        if not cells_a and not cells_b:
            ui.message(_TITLE, "Nothing to highlight.", frame)
            return
        with R.undo_context(model, _TITLE):
            painted = self._paint(model, addr_a, cells_a, color_value)
            if rng_b is not None and cells_b:
                painted += self._paint(model, rng_b.getRangeAddress(), cells_b, color_value)
        ui.message(_TITLE, "Highlighted %d cell(s) in %s." % (painted, colour.lower()), frame)

    @staticmethod
    def _rows(mode, rows_a, rows_b, cols, opts):
        if mode == M.DUP_ALL:
            return M.duplicates(rows_a, cols, opts), set()
        if mode == M.DUP_REPEATS:
            return M.duplicates(rows_a, cols, opts, repeats_only=True), set()
        if mode == M.UNIQUE:
            return M.unique(rows_a, cols, opts), set()
        return M.membership(rows_a, rows_b, cols, opts, present=(mode == M.IN_BOTH))

    @staticmethod
    def _paint(model, addr, cells, color_value):
        """Colour `cells` ({(row, col)} relative to `addr`), merging them into rectangles."""
        sheet = model.getSheets().getByIndex(addr.Sheet)
        # rows that share the same set of columns become one rectangle per run
        by_row = {}
        for r, c in cells:
            by_row.setdefault(r, set()).add(c)
        total = 0
        for row_run in _runs_by_pattern(by_row):
            r1, r2, cols = row_run
            for c1, c2 in R.group_runs(cols):
                rng = sheet.getCellRangeByPosition(addr.StartColumn + c1, addr.StartRow + r1,
                                                   addr.StartColumn + c2, addr.StartRow + r2)
                rng.setPropertyValue("CellBackColor", color_value)
                total += (r2 - r1 + 1) * (c2 - c1 + 1)
        return total

    # ---- clear -------------------------------------------------------------------
    def _clear(self, frame, controller):
        ui = self.services.get("ui")
        model = controller.getModel()
        ranges = R.selected_ranges(controller)
        if not ranges:
            ui.message(_TITLE, "Select a range of cells first.", frame, "warning")
            return
        palette = set(M.PALETTE.values())
        cleared = 0
        with R.undo_context(model, "Clear Highlight"):
            for rng in ranges:
                addr = rng.getRangeAddress()
                sheet = model.getSheets().getByIndex(addr.Sheet)
                if addr.StartColumn == addr.EndColumn and addr.StartRow == addr.EndRow:
                    rng = R.used_area(sheet)               # a lone cell: clear the whole used area
                    addr = rng.getRangeAddress()
                if (addr.EndRow - addr.StartRow + 1) * (addr.EndColumn - addr.StartColumn + 1) > _MAX_CLEAR_CELLS:
                    ui.message(_TITLE, "The selection is too large; select a smaller range.", frame, "warning")
                    return
                cleared += self._clear_range(sheet, addr, palette)
        ui.message(_TITLE, "Cleared %d highlighted cell(s)." % cleared if cleared
                   else "No plugin highlights found in the selection.", frame)

    @staticmethod
    def _clear_range(sheet, addr, palette):
        cleared = 0
        width = addr.EndColumn - addr.StartColumn + 1
        for row in range(addr.StartRow, addr.EndRow + 1):
            line = sheet.getCellRangeByPosition(addr.StartColumn, row, addr.EndColumn, row)
            if line.getPropertyState("CellBackColor").value == "AMBIGUOUS_VALUE":
                for col in range(addr.StartColumn, addr.EndColumn + 1):     # mixed row: look cell by cell
                    cell = sheet.getCellByPosition(col, row)
                    if cell.getPropertyValue("CellBackColor") in palette:
                        cell.setPropertyValue("CellBackColor", _NO_FILL)
                        cleared += 1
            elif line.getPropertyValue("CellBackColor") in palette:
                line.setPropertyValue("CellBackColor", _NO_FILL)
                cleared += width
        return cleared


def _runs_by_pattern(by_row):
    """{row: {cols}} -> [(first_row, last_row, sorted cols)] merging consecutive rows with the
    same set of columns, so a tall highlight is one rectangle instead of one per row."""
    out = []
    for row in sorted(by_row):
        cols = sorted(by_row[row])
        if out and out[-1][1] == row - 1 and out[-1][2] == cols:
            out[-1] = (out[-1][0], row, cols)
        else:
            out.append((row, row, cols))
    return out
