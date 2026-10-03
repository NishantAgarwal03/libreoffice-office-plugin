"""Set Decimals: every selected number gets the same number of decimals.

 * values are rounded (typed numbers stay numbers);
 * formulas stay formulas, wrapped in ROUND / ROUNDUP / ROUNDDOWN (an existing outer wrapper is
   replaced, never nested), or optionally left alone, or unwrapped;
 * the cells are displayed with exactly that many decimals (the cell's own number format is kept
   and only its decimals part is rewritten; percent formats round the percentage).
Dates, times, text, booleans, fractions, errors, text-valued formulas and array formulas are skipped
and counted. Cell setters only, so Ctrl+Z reverts a run in one step.
"""

from collections import Counter

from office_plugin.core.feature import Feature
from office_plugin.services import calc_ranges as R
from office_plugin.services import decimals as D

_TITLE = "Set Decimals"
_WRAP, _LEAVE, _REMOVE = "Wrap in ROUND (formulas stay formulas)", "Leave formulas as they are", "Remove the ROUND wrapper"
_DATE, _TIME, _TEXT, _FRACTION, _LOGICAL = 2, 4, 256, 64, 1024
_VALUE_CELLS, _FORMULA_CELLS = 1, 16
_RESULT_STRING, _RESULT_ERROR = 2, 4


class CalcDecimalsFeature(Feature):
    def execute(self, command, frame, args):
        docs = self.services.get("doc_utils")
        ui = self.services.get("ui")
        controller = docs.controller_of(frame) or docs.active_controller()
        if controller is None:
            return
        ranges = R.selected_ranges(controller)
        if not ranges:
            ui.message(_TITLE, "Select the cells first.", frame, "warning")
            return
        answer = ui.form(_TITLE, [
            {"type": "label", "text": "Give the selected numbers the same number of decimals."},
            {"type": "dropdown", "name": "places", "label": "Decimal places",
             "items": [str(i) for i in range(D.MAX_DECIMALS + 1)], "value": "2"},
            {"type": "dropdown", "name": "mode", "label": "Rounding", "items": list(D.MODES), "value": D.ROUND},
            {"type": "dropdown", "name": "formulas", "label": "Formulas", "items": [_WRAP, _LEAVE, _REMOVE],
             "value": _WRAP},
            {"type": "checkbox", "name": "values", "value": True, "label": "Round typed numbers (values)"},
            {"type": "checkbox", "name": "display", "value": True,
             "label": "Show exactly this many decimals"},
        ], frame)
        if answer is None:
            return
        stats = self._apply(controller.getModel(), ranges, int(answer["places"]), answer)
        ui.message(_TITLE, self._summary(stats, int(answer["places"]), answer), frame)

    # ------------------------------------------------------------------------------------------
    def _apply(self, model, ranges, places, opts):
        stats = Counter()
        formats = model.getNumberFormats()
        info_cache, key_cache = {}, {}

        def info(key):
            if key not in info_cache:
                props = formats.getByKey(key)
                code, locale, kind = props.getPropertyValue("FormatString"), props.getPropertyValue("Locale"), \
                    props.getPropertyValue("Type")
                skip = bool(kind & (_DATE | _TIME | _TEXT | _FRACTION | _LOGICAL))
                info_cache[key] = (code, locale, skip, D.is_percent(code))
            return info_cache[key]

        def show(cell, key, code, locale):
            new_key = key_cache.get((key, places))
            if new_key is None:
                new_code = D.format_with_decimals(code, places)
                new_key = formats.queryKey(new_code, locale, False)
                if new_key == -1:
                    new_key = formats.addNew(new_code, locale)
                key_cache[(key, places)] = new_key
            if new_key != key:
                cell.setPropertyValue("NumberFormat", new_key)
                stats["display"] += 1

        with R.undo_context(model, _TITLE):
            for rng in ranges:
                for cell in self._cells(rng, _VALUE_CELLS):
                    key = cell.getPropertyValue("NumberFormat")
                    code, locale, skip, percent = info(key)
                    if skip:
                        stats["skipped"] += 1
                        continue
                    if opts["values"]:
                        old = cell.getValue()
                        new = D.round_value(old, min(places + (2 if percent else 0), D.MAX_DECIMALS), opts["mode"])
                        if new != old:
                            cell.setValue(new)
                            stats["rounded"] += 1
                    if opts["display"]:
                        show(cell, key, code, locale)
                for cell in self._cells(rng, _FORMULA_CELLS):
                    key = cell.getPropertyValue("NumberFormat")
                    code, locale, skip, percent = info(key)
                    result_type = cell.getPropertyValue("FormulaResultType2")
                    if cell.getArrayFormula() not in ("", None):
                        stats["array"] += 1
                        continue
                    if skip or result_type in (_RESULT_STRING, _RESULT_ERROR) or cell.getError():
                        stats["skipped"] += 1
                        continue
                    eff = min(places + (2 if percent else 0), D.MAX_DECIMALS)
                    formula = cell.getFormula()
                    if opts["formulas"] == _WRAP:
                        new, status = D.wrap_round(formula, eff, opts["mode"])
                        if status != "same":
                            cell.setFormula(new)
                            stats["wrapped" if status == "wrapped" else "updated"] += 1
                    elif opts["formulas"] == _REMOVE:
                        new = D.remove_round(formula)
                        if new is not None:
                            cell.setFormula(new)
                            stats["unwrapped"] += 1
                        continue                                   # removing a wrapper never touches the display
                    if opts["display"]:
                        show(cell, key, code, locale)
        return stats

    @staticmethod
    def _cells(rng, flags):
        found = rng.queryContentCells(flags)
        if found.getCount() == 0:
            return []
        enum = found.getCells().createEnumeration()
        cells = []
        while enum.hasMoreElements():
            cells.append(enum.nextElement())
        return cells

    @staticmethod
    def _summary(stats, places, opts):
        lines = []
        if stats["rounded"]:
            lines.append("%d value(s) rounded to %d decimal(s)." % (stats["rounded"], places))
        if stats["wrapped"]:
            lines.append("%d formula(s) wrapped in ROUND." % stats["wrapped"])
        if stats["updated"]:
            lines.append("%d existing ROUND wrapper(s) updated." % stats["updated"])
        if stats["unwrapped"]:
            lines.append("%d ROUND wrapper(s) removed." % stats["unwrapped"])
        if stats["display"]:
            lines.append("%d cell(s) now show %d decimal(s)." % (stats["display"], places))
        if not lines:
            lines.append("Nothing needed changing.")
        skipped = []
        if stats["skipped"]:
            skipped.append("%d not numbers of this kind (dates, text, errors...)" % stats["skipped"])
        if stats["array"]:
            skipped.append("%d array formula(s)" % stats["array"])
        if skipped:
            lines.append("Skipped: " + "; ".join(skipped) + ".")
        return "\n".join(lines)
