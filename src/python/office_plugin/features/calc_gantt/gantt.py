"""Gantt chart logic (pure Python, no uno).

Purpose: everything about a Gantt that is not talking to Calc: guessing which column is which,
         reading/validating task rows, finish-to-start scheduling, timeline periods and the exact
         conditional-format formulas that draw the bars.
Inputs:  task grid rows (cell values), a column mapping, date serials (days since the document's
         NullDate), the scale.
Outputs: Task records, problems, schedule updates, period start serials, formula strings.
Must not: import uno.
"""

import math
import re
from collections import namedtuple
from datetime import date

from office_plugin.services.calc_ranges import col_name

DURATION, END = "duration", "end"          # how the table gives dates: Start+Duration or Start+End
DAY, WEEK, MONTH, AUTO = "Day", "Week", "Month", "Auto"
MAX_COLUMNS = 400
FIELDS = ("name", "id", "start", "duration", "end", "pct", "preds", "lag")

Task = namedtuple("Task", "row id start duration end preds lag")      # row: 0-based index in the table data
Update = namedtuple("Update", "row start end")


class GanttError(ValueError):
    """Carries a list of human-readable problems."""

    def __init__(self, problems):
        self.problems = list(problems)
        ValueError.__init__(self, "\n".join(self.problems))


# ======================================================================================
# column guessing
# ======================================================================================
def _tokens(text):
    return set(re.findall(r"[a-z0-9]+|[#%]", str(text).lower()))


_RULES = (   # (field, token test) in claim order
    ("pct", lambda t: bool(t & {"%", "pct", "percent", "progress", "complete", "completed", "done"})),
    ("preds", lambda t: bool(t & {"pred", "predecessor", "predecessors", "depends", "dependency", "dependencies", "after"})),
    ("lag", lambda t: bool(t & {"lag", "delay", "lead"})),
    ("id", lambda t: bool(t & {"id", "no", "#", "wbs", "code", "sr", "sl"})),
    ("duration", lambda t: bool(t & {"duration", "dur", "days", "length"})),
    ("end", lambda t: bool(t & {"end", "finish", "finishes", "ends", "to"})),
    ("start", lambda t: bool(t & {"start", "begin", "starts", "from"})),
    ("name", lambda t: bool(t & {"task", "activity", "name", "description", "item", "work"})),
)


def guess_columns(headers):
    """{field: column index or None} from header texts; each column is claimed at most once."""
    taken, found = set(), {f: None for f in FIELDS}
    for field, test in _RULES:
        for index, text in enumerate(headers):
            if index not in taken and test(_tokens(text)):
                found[field] = index
                taken.add(index)
                break
    return found


# ======================================================================================
# values
# ======================================================================================
def normalize_id(value):
    if value is None or value == "":
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return str(int(value)) if float(value) == int(value) else str(value)
    text = str(value).strip()
    return normalize_id(float(text)) if re.fullmatch(r"\d+\.0+", text) else text


def parse_predecessors(value):
    """3 / '3' / '3, 5' / '3;5 7' / '3FS' -> ['3', '5', ...] (order kept, duplicates removed)."""
    if value is None or value == "":
        return []
    if isinstance(value, (int, float)):
        return [normalize_id(value)]
    out = []
    for part in re.split(r"[,;/\s]+", str(value).strip()):
        part = re.sub(r"(?i)fs$", "", part.strip())
        part = normalize_id(part)
        if part and part not in out:
            out.append(part)
    return out


def normalize_pct(value):
    """None/'' -> 0.0; 0.5 -> 0.5; 50 -> 0.5; '50%' -> 0.5; clamped to 0..1."""
    if value is None or value == "":
        return 0.0
    if isinstance(value, str):
        text = value.strip().replace(",", ".")
        try:
            value = float(text[:-1]) / 100 if text.endswith("%") else float(text)
        except ValueError:
            return 0.0
    value = float(value)
    value = value / 100.0 if value > 1 else value
    return max(0.0, min(1.0, value))


def done_days(duration, pct):
    """Whole days shown as completed (half rounds up, like Calc's ROUND)."""
    return int(math.floor(duration * normalize_pct(pct) + 0.5))


def _date_serial(value):
    if value is None or value == "":
        return None
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError("not a date")
    return int(math.floor(float(value)))


def read_tasks(rows, mapping, mode, first_sheet_row=2):
    """Parse table data rows into Tasks. -> (tasks, problems).

    `rows` are the data rows (header excluded); `mapping` maps FIELDS to column indexes/None;
    `first_sheet_row` is the 1-based sheet row of rows[0] (for messages)."""
    tasks, problems = [], []

    def cell(row, field):
        index = mapping.get(field)
        return row[index] if index is not None and index < len(row) else ""

    for i, row in enumerate(rows):
        label = "row %d" % (first_sheet_row + i)
        if all(v == "" or v is None for v in row):
            continue                                            # blank line inside the table
        try:
            start = _date_serial(cell(row, "start"))
        except ValueError:
            problems.append("%s: the start is not a date" % label)
            continue
        duration = end = None
        try:
            if mode == DURATION:
                raw = cell(row, "duration")
                duration = None if raw in ("", None) else int(round(float(raw)))
                if duration is not None and duration < 0:
                    problems.append("%s: the duration is negative" % label)
                    continue
            else:
                end = _date_serial(cell(row, "end"))
                if start is not None and end is not None and end < start:
                    problems.append("%s: the end is before the start" % label)
                    continue
        except (ValueError, TypeError):
            problems.append("%s: the %s is not valid" % (label, "duration" if mode == DURATION else "end date"))
            continue
        lag_raw = cell(row, "lag")
        try:
            lag = 0 if lag_raw in ("", None) else int(round(float(lag_raw)))
        except (ValueError, TypeError):
            problems.append("%s: the lag is not a number" % label)
            continue
        tasks.append(Task(i, normalize_id(cell(row, "id")), start, duration, end,
                          parse_predecessors(cell(row, "preds")), lag))
    return tasks, problems


def span_days(task, mode):
    """Length in days: 0 = milestone, None = unknown (cannot be scheduled)."""
    if mode == DURATION:
        return task.duration
    if task.start is None:
        return None
    return 0 if task.end is None else task.end - task.start + 1          # End-mode: blank end = milestone


def finish_of(start, days):
    return start if days <= 0 else start + days - 1


# ======================================================================================
# scheduling (finish-to-start)
# ======================================================================================
def schedule(tasks, mode):
    """Compute start/finish of every task that has predecessors.

    -> {row: Update(row, start, finish)} for tasks WITH predecessors. Raises GanttError with all
    problems (unknown / self / duplicate ids, cycles, tasks lacking dates or duration) and writes nothing.
    """
    problems = []
    if not any(t.preds for t in tasks):
        return {}
    by_id, rows_of = {}, {}
    for t in tasks:
        if t.id:
            rows_of.setdefault(t.id, []).append(t.row)
            by_id.setdefault(t.id, t)
    for ident, rows in rows_of.items():
        if len(rows) > 1:
            problems.append("ID %s is used more than once (rows %s)" % (ident, ", ".join(str(r + 2) for r in rows)))
    for t in tasks:
        if t.preds and not t.id:
            problems.append("row %d has predecessors but no ID" % (t.row + 2))
        for p in t.preds:
            if p == t.id:
                problems.append("row %d lists itself as a predecessor" % (t.row + 2))
            elif p not in by_id:
                problems.append("row %d: predecessor %s does not exist" % (t.row + 2, p))
    if problems:
        raise GanttError(problems)

    state, order, stack = {}, [], []

    def visit(task):
        state[task.row] = 1
        stack.append(task)
        for p in task.preds:
            nxt = by_id[p]
            if state.get(nxt.row) == 1:
                cycle = [x.id or str(x.row + 2) for x in stack[stack.index(nxt):]] + [nxt.id]
                raise GanttError(["circular predecessors: " + " -> ".join(cycle)])
            if nxt.row not in state:
                visit(nxt)
        stack.pop()
        state[task.row] = 2
        order.append(task)

    for t in tasks:
        if t.row not in state:
            visit(t)

    finish, updates = {}, {}
    for t in order:                                       # predecessors always come first
        days = span_days(t, mode)
        if t.preds:
            if days is None:
                problems.append("row %d needs a duration (or start and end) to be scheduled" % (t.row + 2))
                continue
            missing = [p for p in t.preds if by_id[p].row not in finish]
            if missing:
                continue                                  # a predecessor already reported a problem
            start = max(finish[by_id[p].row] for p in t.preds) + 1 + t.lag
            updates[t.row] = Update(t.row, start, finish_of(start, days))
            finish[t.row] = finish_of(start, days)
        else:
            if t.start is None:
                problems.append("row %d needs a start date (it has no predecessors)" % (t.row + 2))
                continue
            finish[t.row] = finish_of(t.start, span_days(t, mode) or 0)
    if problems:
        raise GanttError(problems)
    return updates


# ======================================================================================
# timeline
# ======================================================================================
def choose_scale(first, last):
    """Auto scale from the number of days covered."""
    days = last - first + 1
    return DAY if days <= 120 else (WEEK if days <= 730 else MONTH)


def _date(serial, null_ordinal):
    return date.fromordinal(null_ordinal + serial)


def _serial(d, null_ordinal):
    return d.toordinal() - null_ordinal


def _too_many(count):
    raise GanttError(["The timeline needs about %d columns (limit %d): choose a coarser scale "
                      "(Week or Month)." % (count, MAX_COLUMNS)])


def period_starts(first, last, scale, null_ordinal):
    """Serials of the period starts covering first..last plus one extra period. Raises GanttError
    when more than MAX_COLUMNS columns would be needed (checked before building anything)."""
    if last < first:
        first, last = last, first
    if scale == DAY:
        if last - first + 2 > MAX_COLUMNS:
            _too_many(last - first + 2)
        return list(range(first, last + 2))
    if scale == WEEK:
        cur = first - _date(first, null_ordinal).weekday()              # Monday of the first week
        count = (last - cur) // 7 + 2
        if count > MAX_COLUMNS:
            _too_many(count)
        return [cur + 7 * i for i in range(count)]
    if scale == MONTH:
        d = _date(first, null_ordinal)
        year, month, result = d.year, d.month, []
        while True:
            start = _serial(date(year, month, 1), null_ordinal)
            result.append(start)
            if start > last:
                return result
            if len(result) >= MAX_COLUMNS:
                _too_many(len(result) + 1)
            year, month = (year + 1, 1) if month == 12 else (year, month + 1)
    raise ValueError("unknown scale %r" % scale)


HEADER_FORMAT = {DAY: "DD-MMM", WEEK: "DD-MMM", MONTH: "MMM-YY"}
COLUMN_WIDTH = {DAY: 500, WEEK: 700, MONTH: 900}                     # 1/100 mm


# ======================================================================================
# conditional-format formulas
# ======================================================================================
Layout = namedtuple("Layout", "first_row header_row start duration end pct timeline_col")
# first_row / header_row are 0-based sheet rows; the rest are 0-based column indexes (None = not mapped)


def cf_formulas(layout, mode, scale):
    """[(formula, style_name)] in priority order, relative to the top-left timeline data cell.

    Styles: 'Gantt Milestone', 'Gantt Done', 'Gantt Bar'. API grammar (';' separators)."""
    row = layout.first_row + 1                                       # 1-based, relative (no $)
    s = "$%s%d" % (col_name(layout.start), row)
    p = "%s$%d" % (col_name(layout.timeline_col), layout.header_row + 1)
    q = {DAY: p, WEEK: "%s+6" % p, MONTH: "EOMONTH(%s;0)" % p}[scale]
    out = []
    if mode == DURATION:
        d = "$%s%d" % (col_name(layout.duration), row)
        out.append(("AND(ISNUMBER(%s);ISNUMBER(%s);%s=0;%s<=%s;%s>=%s)" % (s, d, d, p, s, q, s), "Gantt Milestone"))
        days, active = d, "%s>0" % d
        finish = "%s+%s-1" % (s, d)
    else:
        e = "$%s%d" % (col_name(layout.end), row)
        out.append(("AND(ISNUMBER(%s);NOT(ISNUMBER(%s));%s<=%s;%s>=%s)" % (s, e, p, s, q, s), "Gantt Milestone"))
        days, active = "(%s-%s+1)" % (e, s), "ISNUMBER(%s)" % e
        finish = e
    if layout.pct is not None:
        pct = "$%s%d" % (col_name(layout.pct), row)
        done = "ROUND(%s*IF(%s>1;%s/100;%s);0)" % (days, pct, pct, pct)
        out.append(("AND(ISNUMBER(%s);%s;%s>0;%s<=%s+%s-1;%s>=%s)" % (s, active, done, p, s, done, q, s),
                    "Gantt Done"))
    out.append(("AND(ISNUMBER(%s);%s;%s<=%s;%s>=%s)" % (s, active, p, finish, q, s), "Gantt Bar"))
    return out


def expected_cell(task_start, days, pct, period_start, period_end):
    """Python model of what the formulas should draw for one cell: 'milestone' / 'done' / 'bar' / ''.
    Used by tests to compare against what LibreOffice renders."""
    if task_start is None:
        return ""
    if days == 0:
        return "milestone" if period_start <= task_start <= period_end else ""
    finish = task_start + days - 1
    if not (period_start <= finish and period_end >= task_start):
        return ""
    d = done_days(days, pct)
    if d > 0 and period_start <= task_start + d - 1:
        return "done"
    return "bar"
