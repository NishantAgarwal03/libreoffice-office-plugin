"""Gantt Chart: a live in-sheet Gantt next to a task table.

The timeline is a block of day/week/month columns to the right of the table (one spare column in
between). The bars are conditional-format rules reading the table's Start / Duration (or End) /
% complete columns, so editing a date moves the bar. Dependencies are resolved by this tool and the
computed Start (and End) written into the table; run the command again to recalculate. The mapping
used is remembered in the document (see services/doc_store), Remove Gantt clears the timeline.

Undo note: cell changes are one undo step; LibreOffice does not record clearing/replacing
conditional formats through its API, so after undoing a Remove run Gantt Chart again to restore bars.
"""

from datetime import date

from office_plugin.core.feature import Feature
from office_plugin.services import calc_format as F
from office_plugin.services import calc_ranges as R
from office_plugin.services import doc_store

from . import gantt as G

_TITLE = "Gantt Chart"
_NONE = "(none)"
_MODE_LABELS = {G.DURATION: "Start + Duration (days)", G.END: "Start + End date"}
_STYLES = (("Gantt Bar", 0x9DC3FF), ("Gantt Done", 0x1F4E9E), ("Gantt Milestone", 0x6A3D9A))
_ALL = 1023


def _store_key(sheet_name):
    return "Gantt." + sheet_name


class CalcGanttFeature(Feature):
    def execute(self, command, frame, args):
        docs = self.services.get("doc_utils")
        controller = docs.controller_of(frame) or docs.active_controller()
        if controller is None:
            return
        if command == "Remove":
            return self._remove(frame, controller)
        return self._run(frame, controller)

    # ======================================================================================
    # create / update
    # ======================================================================================
    def _run(self, frame, controller):
        ui = self.services.get("ui")
        model = controller.getModel()
        ranges = R.selected_ranges(controller)
        if len(ranges) != 1:
            ui.message(_TITLE, "Select the task table first (header row and tasks).", frame, "warning")
            return
        sel = ranges[0]
        sel_addr = sel.getRangeAddress()
        sheet = model.getSheets().getByIndex(sel_addr.Sheet)
        saved = doc_store.get(model, _store_key(sheet.getName()))
        lone = sel_addr.StartColumn == sel_addr.EndColumn and sel_addr.StartRow == sel_addr.EndRow
        if lone and saved:                                  # clicking anywhere re-runs on the saved table
            r1, c1, r2, c2 = saved["table"]
            rng = sheet.getCellRangeByPosition(c1, r1, c2, r2)
        else:
            rng = R.expand_if_single(model, sel)
        a = rng.getRangeAddress()
        grid = R.data_grid(rng)
        if len(grid) < 2:
            ui.message(_TITLE, "The table needs a header row and at least one task.", frame, "warning")
            return

        headers = [str(h) for h in grid[0]]
        labels = R.column_labels(grid[0], a.StartColumn, True)
        values = self._default_values(self._defaults(saved, headers, a), labels)
        while True:
            answer = self._ask(ui, frame, labels, values)
            if answer is None:
                return
            problem = self._validate(answer)
            if problem is None:
                break
            ui.message(_TITLE, problem, frame, "warning")
            values = answer                                   # ask again, keeping what was chosen
        self._build(frame, controller, model, sheet, rng, a, grid, labels, answer, saved)

    def _defaults(self, saved, headers, a):
        guess = G.guess_columns(headers)
        mapping = guess
        mode = G.END if (guess["end"] is not None and guess["duration"] is None) else G.DURATION
        scale = G.AUTO
        if saved and saved.get("width") == len(headers):
            mapping, mode, scale = saved["map"], saved["mode"], saved.get("scale", G.AUTO)
        return {"map": mapping, "mode": mode, "scale": scale, "headers": headers}

    @staticmethod
    def _label(labels, index):
        return _NONE if index is None or index >= len(labels) else labels[index]

    def _default_values(self, defaults, labels):
        m, mode = defaults["map"], defaults["mode"]
        second = m.get("duration") if mode == G.DURATION else m.get("end")
        return {"name": self._label(labels, m.get("name")), "id": self._label(labels, m.get("id")),
                "start": self._label(labels, m.get("start")), "mode": _MODE_LABELS[mode],
                "second": self._label(labels, second), "pct": self._label(labels, m.get("pct")),
                "preds": self._label(labels, m.get("preds")), "lag": self._label(labels, m.get("lag")),
                "scale": defaults["scale"]}

    def _ask(self, ui, frame, labels, values):
        items = [_NONE] + labels

        def drop(name, label):
            return {"type": "dropdown", "name": name, "label": label, "items": items, "value": values[name]}

        return ui.form(_TITLE, [
            {"type": "label", "text": "Map your table's columns. A timeline is drawn to the right;\n"
                                      "bars follow the dates live (conditional formatting)."},
            drop("name", "Task name (optional)"),
            drop("start", "Start date"),
            {"type": "dropdown", "name": "mode", "label": "Dates are given as",
             "items": list(_MODE_LABELS.values()), "value": values["mode"]},
            drop("second", "Duration (days) or End date column"),
            drop("pct", "% complete (optional)"),
            drop("id", "Task ID (needed for dependencies)"),
            drop("preds", "Predecessors, e.g. 3 or 3, 5 (optional)"),
            drop("lag", "Lag in days (optional)"),
            {"type": "dropdown", "name": "scale", "label": "Timeline scale",
             "items": [G.AUTO, G.DAY, G.WEEK, G.MONTH], "value": values["scale"]},
        ], frame)

    def _validate(self, answer):
        picked = {f: answer[f] for f in ("name", "id", "start", "second", "pct", "preds", "lag")}
        if picked["start"] == _NONE:
            return "Choose the Start date column."
        if picked["second"] == _NONE:
            return "Choose the Duration or End date column."
        chosen = [v for v in picked.values() if v != _NONE]
        if len(set(chosen)) != len(chosen):
            return "Each column can be used for one purpose only."
        if picked["preds"] != _NONE and picked["id"] == _NONE:
            return "Predecessors need a Task ID column."
        return None

    # ---- build ---------------------------------------------------------------------------
    def _build(self, frame, controller, model, sheet, rng, a, grid, labels, answer, saved):
        ui = self.services.get("ui")
        index = {lab: i for i, lab in enumerate(labels)}
        mode = G.DURATION if answer["mode"] == _MODE_LABELS[G.DURATION] else G.END
        mapping = {f: index.get(answer[f]) for f in ("name", "id", "start", "pct", "preds", "lag")}
        mapping["duration"] = index.get(answer["second"]) if mode == G.DURATION else None
        mapping["end"] = index.get(answer["second"]) if mode == G.END else None

        first_sheet_row = a.StartRow + 2
        tasks, problems = G.read_tasks(grid[1:], mapping, mode, first_sheet_row)
        try:
            updates = G.schedule(tasks, mode)
        except G.GanttError as exc:
            problems += exc.problems
            updates = {}
        if problems:
            ui.message(_TITLE, "Nothing was changed. Please fix:\n" + "\n".join(problems[:12])
                       + ("\n... and %d more" % (len(problems) - 12) if len(problems) > 12 else ""),
                       frame, "warning")
            return

        extent = self._extent(tasks, updates, mode)
        if extent is None:
            ui.message(_TITLE, "No task has a start date.", frame, "warning")
            return
        nd = model.getPropertyValue("NullDate")
        null_ordinal = date(nd.Year, nd.Month, nd.Day).toordinal()
        scale = G.choose_scale(*extent) if answer["scale"] == G.AUTO else answer["scale"]
        try:
            periods = G.period_starts(extent[0], extent[1], scale, null_ordinal)
        except G.GanttError as exc:
            ui.message(_TITLE, exc.problems[0], frame, "warning")
            return

        # where the timeline goes: header row of the table, one spare column after it
        col0 = a.EndColumn + 2
        last_col = col0 + len(periods) - 1
        if last_col >= sheet.getColumns().getCount():
            ui.message(_TITLE, "The timeline does not fit on the sheet.", frame, "warning")
            return
        old = saved["timeline"] if saved else None
        if self._occupied(sheet, a, col0, last_col, old) and not ui.confirm(
                _TITLE, "The timeline area next to the table is not empty.\nOverwrite it?", frame):
            return

        layout = G.Layout(first_row=a.StartRow + 1, header_row=a.StartRow,
                          start=a.StartColumn + mapping["start"],
                          duration=None if mapping["duration"] is None else a.StartColumn + mapping["duration"],
                          end=None if mapping["end"] is None else a.StartColumn + mapping["end"],
                          pct=None if mapping["pct"] is None else a.StartColumn + mapping["pct"],
                          timeline_col=col0)
        with R.undo_context(model, _TITLE):
            if old:
                self._clear_timeline(sheet, old)
            self._write_dates(sheet, a, mapping, mode, tasks, updates)
            header = sheet.getCellRangeByPosition(col0, a.StartRow, last_col, a.StartRow)
            for i, serial in enumerate(periods):
                R.set_cell_value(sheet, col0 + i, a.StartRow, float(serial))
            header.setPropertyValue("NumberFormat", F.ensure_date_format(model, G.HEADER_FORMAT[scale]))
            header.setPropertyValue("RotateAngle", 9000)
            header.getColumns().setPropertyValue("Width", G.COLUMN_WIDTH[scale])
            for name, color in _STYLES:
                F.ensure_cell_style(model, name, color)
            body = sheet.getCellRangeByPosition(col0, a.StartRow + 1, last_col, a.EndRow)
            F.set_conditional_formats(body, G.cf_formulas(layout, mode, scale))
            doc_store.put(model, _store_key(sheet.getName()), {
                "table": [a.StartRow, a.StartColumn, a.EndRow, a.EndColumn], "width": len(grid[0]),
                "map": mapping, "mode": mode, "scale": answer["scale"],
                "timeline": [a.StartRow, col0, a.EndRow, last_col]})
        text = "Gantt chart drawn: %d task(s), %d %s column(s), starting at column %s." % (
            len([t for t in tasks if self._has_dates(t, updates, mode)]), len(periods), scale.lower(),
            R.col_name(col0))
        if updates:
            text += "\nRescheduled %d task(s) from their predecessors." % len(updates)
        ui.message(_TITLE, text, frame)

    # ---- helpers -------------------------------------------------------------------------
    @staticmethod
    def _has_dates(task, updates, mode):
        return task.start is not None or task.row in updates

    @staticmethod
    def _extent(tasks, updates, mode):
        first = last = None
        for t in tasks:
            if t.row in updates:
                start, finish = updates[t.row].start, updates[t.row].end
            elif t.start is not None:
                days = G.span_days(t, mode) or 0
                start, finish = t.start, G.finish_of(t.start, days)
            else:
                continue
            first = start if first is None else min(first, start)
            last = finish if last is None else max(last, finish)
        return None if first is None else (first, last)

    @staticmethod
    def _occupied(sheet, a, col0, last_col, old):
        """True when the target area holds cells that are not the previous timeline."""
        block = sheet.getCellRangeByPosition(col0, a.StartRow, last_col, a.EndRow)
        for r, row in enumerate(block.getFormulaArray()):
            for c, text in enumerate(row):
                if text == "":
                    continue
                inside_old = old and old[0] <= a.StartRow + r <= old[2] and old[1] <= col0 + c <= old[3]
                if not inside_old:
                    return True
        return False

    @staticmethod
    def _clear_timeline(sheet, old):
        block = sheet.getCellRangeByPosition(old[1], old[0], old[3], old[2])
        F.clear_conditional_formats(block)
        block.clearContents(_ALL)

    @staticmethod
    def _write_dates(sheet, a, mapping, mode, tasks, updates):
        by_row = {t.row: t for t in tasks}
        for row, update in updates.items():
            sheet_row = a.StartRow + 1 + row
            R.set_cell_value(sheet, a.StartColumn + mapping["start"], sheet_row, float(update.start))
            task = by_row[row]
            # End-mode: a blank end marks a milestone, so never write one for it
            if mode == G.END and task.end is not None:
                R.set_cell_value(sheet, a.StartColumn + mapping["end"], sheet_row, float(update.end))

    # ======================================================================================
    # remove
    # ======================================================================================
    def _remove(self, frame, controller):
        ui = self.services.get("ui")
        model = controller.getModel()
        sheet = controller.getActiveSheet()
        saved = doc_store.get(model, _store_key(sheet.getName()))
        if not saved:
            ui.message(_TITLE, "No Gantt chart made by this tool was found on this sheet.", frame)
            return
        t = saved["timeline"]
        if not ui.confirm(_TITLE, "Remove the Gantt timeline (columns %s to %s)?\n"
                                  "The task table and the computed dates stay."
                          % (R.col_name(t[1]), R.col_name(t[3])), frame):
            return
        with R.undo_context(model, "Remove Gantt"):
            self._clear_timeline(sheet, t)
            doc_store.remove(model, _store_key(sheet.getName()))
        ui.message(_TITLE, "Gantt timeline removed.", frame)
