import unittest
from datetime import date

import tests  # noqa: F401
from office_plugin.features.calc_gantt import gantt as G

NULL = date(1899, 12, 30).toordinal()


def ser(y, m, d):
    return date(y, m, d).toordinal() - NULL


def task(row, ident, start=None, duration=None, end=None, preds=(), lag=0):
    return G.Task(row, ident, start, duration, end, list(preds), lag)


class GuessColumnsTests(unittest.TestCase):
    def test_typical_headers(self):
        found = G.guess_columns(["Task ID", "Task Name", "Start Date", "End Date", "Duration (days)",
                                 "% Complete", "Predecessors", "Lag"])
        self.assertEqual(found, {"id": 0, "name": 1, "start": 2, "end": 3, "duration": 4,
                                 "pct": 5, "preds": 6, "lag": 7})

    def test_partial_headers_and_single_claim(self):
        found = G.guess_columns(["Activity", "Start", "Finish", "Progress"])
        self.assertEqual((found["name"], found["start"], found["end"], found["pct"]), (0, 1, 2, 3))
        self.assertIsNone(found["duration"])
        self.assertIsNone(found["preds"])
        self.assertEqual(sorted(i for i in found.values() if i is not None), [0, 1, 2, 3])   # no column twice

    def test_unrecognised(self):
        self.assertTrue(all(v is None for v in G.guess_columns(["Foo", "Bar"]).values()))


class ValueTests(unittest.TestCase):
    def test_ids_and_predecessors(self):
        self.assertEqual(G.normalize_id(3.0), "3")
        self.assertEqual(G.normalize_id(" A-1 "), "A-1")
        self.assertEqual(G.normalize_id("7.0"), "7")
        self.assertEqual(G.parse_predecessors(3.0), ["3"])
        self.assertEqual(G.parse_predecessors("3, 5;7 3"), ["3", "5", "7"])
        self.assertEqual(G.parse_predecessors("3FS, 4fs"), ["3", "4"])
        self.assertEqual(G.parse_predecessors(""), [])
        self.assertEqual(G.parse_predecessors(None), [])

    def test_percent(self):
        self.assertEqual(G.normalize_pct(0.5), 0.5)
        self.assertEqual(G.normalize_pct(50), 0.5)
        self.assertEqual(G.normalize_pct("50%"), 0.5)
        self.assertEqual(G.normalize_pct(""), 0.0)
        self.assertEqual(G.normalize_pct(250), 1.0)
        self.assertEqual(G.normalize_pct(-3), 0.0)
        self.assertEqual(G.normalize_pct("abc"), 0.0)

    def test_done_days_rounds_half_up(self):
        self.assertEqual(G.done_days(3, 0.5), 2)        # 1.5 -> 2
        self.assertEqual(G.done_days(4, 100), 4)
        self.assertEqual(G.done_days(10, 0.25), 3)      # 2.5 -> 3
        self.assertEqual(G.done_days(1, 0.4), 0)
        self.assertEqual(G.done_days(5, 0), 0)


class ReadTasksTests(unittest.TestCase):
    MAP = {"id": 0, "start": 1, "duration": 2, "preds": 3, "lag": 4, "end": None, "pct": None, "name": None}

    def test_reads_rows_and_skips_blank_lines(self):
        rows = [(1.0, 46000.0, 3.0, "", ""), ("", "", "", "", ""), (2.0, "", 2.0, "1", 1.0)]
        tasks, problems = G.read_tasks(rows, self.MAP, G.DURATION)
        self.assertEqual(problems, [])
        self.assertEqual([(t.row, t.id, t.start, t.duration, t.preds, t.lag) for t in tasks],
                         [(0, "1", 46000, 3, [], 0), (2, "2", None, 2, ["1"], 1)])

    def test_problems_are_reported_with_sheet_rows(self):
        rows = [(1.0, "tomorrow", 3.0, "", ""), (2.0, 46000.0, -1.0, "", ""), (3.0, 46000.0, "x", "", "")]
        tasks, problems = G.read_tasks(rows, self.MAP, G.DURATION)
        self.assertEqual(tasks, [])
        self.assertEqual(len(problems), 3)
        self.assertIn("row 2: the start is not a date", problems[0])
        self.assertIn("row 3: the duration is negative", problems[1])
        self.assertIn("row 4: the duration is not valid", problems[2])

    def test_end_mode_checks_order(self):
        mapping = dict(self.MAP, duration=None, end=2)
        _, problems = G.read_tasks([(1.0, 46010.0, 46000.0, "", "")], mapping, G.END)
        self.assertIn("end is before the start", problems[0])
        tasks, problems = G.read_tasks([(1.0, 46000.0, "", "", "")], mapping, G.END)
        self.assertEqual((problems, tasks[0].end), ([], None))               # blank end = milestone


class ScheduleTests(unittest.TestCase):
    def test_no_predecessors_means_nothing_to_do(self):
        self.assertEqual(G.schedule([task(0, "1", 100, 3)], G.DURATION), {})

    def test_chain(self):
        tasks = [task(0, "1", 100, 3), task(1, "2", None, 2, preds=["1"]), task(2, "3", None, 4, preds=["2"])]
        u = G.schedule(tasks, G.DURATION)
        self.assertEqual((u[1].start, u[1].end), (103, 104))          # finish of 1 is 102 -> starts 103
        self.assertEqual((u[2].start, u[2].end), (105, 108))

    def test_several_predecessors_take_the_latest_finish(self):
        tasks = [task(0, "1", 100, 3), task(1, "2", 100, 10), task(2, "3", None, 1, preds=["1", "2"])]
        self.assertEqual(G.schedule(tasks, G.DURATION)[2].start, 110)

    def test_lag_positive_and_negative(self):
        tasks = [task(0, "1", 100, 3), task(1, "2", None, 1, preds=["1"], lag=2), task(2, "3", None, 1, preds=["1"], lag=-1)]
        u = G.schedule(tasks, G.DURATION)
        self.assertEqual(u[1].start, 105)
        self.assertEqual(u[2].start, 102)

    def test_milestone_and_order_independence(self):
        tasks = [task(0, "M", None, 0, preds=["B"]), task(1, "B", None, 2, preds=["A"]), task(2, "A", 100, 3)]
        u = G.schedule(tasks, G.DURATION)                              # successors listed before predecessors
        self.assertEqual((u[1].start, u[1].end), (103, 104))
        self.assertEqual((u[0].start, u[0].end), (105, 105))           # milestone: one day, finish == start

    def test_end_mode_derives_duration_from_existing_dates(self):
        tasks = [task(0, "1", 100, end=102), task(1, "2", 50, end=54, preds=["1"]),
                 task(2, "3", 60, None, preds=["2"])]                  # blank end -> milestone
        u = G.schedule(tasks, G.END)
        self.assertEqual((u[1].start, u[1].end), (103, 107))           # 5 days kept
        self.assertEqual((u[2].start, u[2].end), (108, 108))

    def test_rerun_is_idempotent(self):
        tasks = [task(0, "1", 100, 3), task(1, "2", 1, 2, preds=["1"]), task(2, "3", 1, 4, preds=["2"])]
        first = G.schedule(tasks, G.DURATION)
        again = [tasks[0]] + [t._replace(start=first[t.row].start) for t in tasks[1:]]
        self.assertEqual(G.schedule(again, G.DURATION), first)

    def test_problems(self):
        cases = {
            "unknown": [task(0, "1", 100, 3, preds=["9"])],
            "self": [task(0, "1", 100, 3, preds=["1"])],
            "duplicate": [task(0, "1", 100, 3), task(1, "1", 100, 3), task(2, "2", None, 1, preds=["1"])],
            "no id": [task(0, "", None, 1, preds=["1"]), task(1, "1", 100, 1)],
            "no duration": [task(0, "1", 100, 3), task(1, "2", None, None, preds=["1"])],
            "root without start": [task(0, "1", None, 3), task(1, "2", None, 1, preds=["1"])],
        }
        for label, tasks in cases.items():
            with self.subTest(label), self.assertRaises(G.GanttError) as ctx:
                G.schedule(tasks, G.DURATION)
            self.assertTrue(ctx.exception.problems)

    def test_cycle_is_named(self):
        tasks = [task(0, "A", None, 1, preds=["B"]), task(1, "B", None, 1, preds=["A"])]
        with self.assertRaises(G.GanttError) as ctx:
            G.schedule(tasks, G.DURATION)
        self.assertIn("circular predecessors: A -> B -> A", ctx.exception.problems[0])

    def test_error_lists_every_problem(self):
        tasks = [task(0, "1", 100, 3, preds=["8"]), task(1, "2", 100, 3, preds=["9"])]
        with self.assertRaises(G.GanttError) as ctx:
            G.schedule(tasks, G.DURATION)
        self.assertEqual(len(ctx.exception.problems), 2)


class TimelineTests(unittest.TestCase):
    def test_scale_choice(self):
        self.assertEqual(G.choose_scale(0, 119), G.DAY)
        self.assertEqual(G.choose_scale(0, 120), G.WEEK)
        self.assertEqual(G.choose_scale(0, 729), G.WEEK)
        self.assertEqual(G.choose_scale(0, 730), G.MONTH)

    def test_day_periods_include_one_extra(self):
        s = ser(2026, 10, 5)
        self.assertEqual(G.period_starts(s, s + 2, G.DAY, NULL), [s, s + 1, s + 2, s + 3])

    def test_week_periods_start_on_monday(self):
        wed = ser(2026, 10, 7)                                           # a Wednesday
        starts = G.period_starts(wed, wed + 10, G.WEEK, NULL)
        self.assertEqual(starts[0], ser(2026, 10, 5))                     # Monday of that week
        self.assertTrue(all(date.fromordinal(NULL + x).weekday() == 0 for x in starts))
        self.assertEqual(starts, [ser(2026, 10, 5), ser(2026, 10, 12), ser(2026, 10, 19)])   # covers wed+10 = 17th, +1 extra

    def test_month_periods_including_leap_february_and_year_end(self):
        starts = G.period_starts(ser(2023, 12, 15), ser(2024, 3, 2), G.MONTH, NULL)
        self.assertEqual([date.fromordinal(NULL + x) for x in starts],
                         [date(2023, 12, 1), date(2024, 1, 1), date(2024, 2, 1), date(2024, 3, 1), date(2024, 4, 1)])

    def test_reversed_arguments_and_caps(self):
        self.assertEqual(G.period_starts(10, 5, G.DAY, NULL), G.period_starts(5, 10, G.DAY, NULL))
        with self.assertRaises(G.GanttError) as ctx:
            G.period_starts(0, 500, G.DAY, NULL)
        self.assertIn("coarser scale", ctx.exception.problems[0])
        with self.assertRaises(G.GanttError):
            G.period_starts(ser(1990, 1, 1), ser(2060, 1, 1), G.WEEK, NULL)
        self.assertLessEqual(len(G.period_starts(ser(2000, 1, 1), ser(2030, 1, 1), G.MONTH, NULL)), G.MAX_COLUMNS)


class FormulaTests(unittest.TestCase):
    # table: header row 0, data from row 1; Start=C, Duration=D, Pct=E; timeline starts at G
    L = G.Layout(first_row=1, header_row=0, start=2, duration=3, end=None, pct=4, timeline_col=6)

    def test_duration_mode_day_scale(self):
        f = G.cf_formulas(self.L, G.DURATION, G.DAY)
        d = "ROUND($D2*IF($E2>1;$E2/100;$E2);0)"
        self.assertEqual(f, [
            ("AND(ISNUMBER($C2);ISNUMBER($D2);$D2=0;G$1<=$C2;G$1>=$C2)", "Gantt Milestone"),
            ("AND(ISNUMBER($C2);$D2>0;%s>0;G$1<=$C2+%s-1;G$1>=$C2)" % (d, d), "Gantt Done"),
            ("AND(ISNUMBER($C2);$D2>0;G$1<=$C2+$D2-1;G$1>=$C2)", "Gantt Bar")])

    def test_period_end_by_scale(self):
        self.assertIn("G$1+6>=$C2", G.cf_formulas(self.L, G.DURATION, G.WEEK)[-1][0])
        self.assertIn("EOMONTH(G$1;0)>=$C2", G.cf_formulas(self.L, G.DURATION, G.MONTH)[-1][0])

    def test_without_progress_column_there_is_no_done_rule(self):
        f = G.cf_formulas(self.L._replace(pct=None), G.DURATION, G.DAY)
        self.assertEqual([s for _, s in f], ["Gantt Milestone", "Gantt Bar"])

    def test_end_mode(self):
        layout = G.Layout(first_row=3, header_row=2, start=1, duration=None, end=2, pct=3, timeline_col=5)
        f = G.cf_formulas(layout, G.END, G.DAY)
        self.assertEqual(f[0], ("AND(ISNUMBER($B4);NOT(ISNUMBER($C4));F$3<=$B4;F$3>=$B4)", "Gantt Milestone"))
        self.assertEqual(f[-1], ("AND(ISNUMBER($B4);ISNUMBER($C4);F$3<=$C4;F$3>=$B4)", "Gantt Bar"))
        self.assertIn("ROUND(($C4-$B4+1)*IF($D4>1;$D4/100;$D4);0)", f[1][0])

    def test_python_model_matches_the_documented_example(self):
        base = 46300
        row = lambda s, d, p: [G.expected_cell(s, d, p, base + i, base + i) or "." for i in range(6)]
        self.assertEqual(row(base, 3, 0.5), ["done", "done", "bar", ".", ".", "."])
        self.assertEqual(row(base + 2, 0, 0), [".", ".", "milestone", ".", ".", "."])
        self.assertEqual(row(base + 1, 4, 100), [".", "done", "done", "done", "done", "."])
        self.assertEqual(G.expected_cell(None, 3, 0, 0, 10), "")
        # a week period overlapping the bar at all shows it
        self.assertEqual(G.expected_cell(base + 3, 2, 0, base, base + 6), "bar")


if __name__ == "__main__":
    unittest.main()
