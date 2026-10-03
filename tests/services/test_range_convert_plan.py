import unittest

import tests  # noqa: F401
from office_plugin.services.range_convert import plan


def accepts_numbers(data, formula):
    return isinstance(data, float)


def times_two(value):
    if value < 0:
        raise ValueError("negative")
    return value * 2


class PlanTests(unittest.TestCase):
    def test_converts_skips_and_reports(self):
        data = [(1.0, "text", ""), (-5.0, 3.0, "")]
        formulas = [("1", "text", ""), ("-5", "=A1+B1", "")]
        p = plan(data, formulas, accepts_numbers, times_two)
        self.assertEqual(p.skipped, 1)                                    # "text"
        self.assertEqual([(i.row, i.col, i.result, i.error) for i in p.items],
                         [(0, 0, 2.0, None), (1, 0, None, "negative"), (1, 1, 6.0, None)])

    def test_empty_cells_are_ignored_not_counted(self):
        p = plan([("", None)], [("", "")], accepts_numbers, times_two)
        self.assertEqual((p.items, p.skipped), ([], 0))

    def test_a_formula_returning_empty_text_is_a_cell(self):
        p = plan([("",)], [('=""',)], accepts_numbers, times_two)
        self.assertEqual(p.skipped, 1)                                    # has a formula, result is text -> skipped

    def test_skip_set_is_counted_and_never_converted(self):
        data = [(1.0, 2.0), (3.0, "")]
        p = plan(data, [("1", "2"), ("3", "")], accepts_numbers, times_two, skip={(0, 1), (1, 1)})
        self.assertEqual([(i.row, i.col) for i in p.items], [(0, 0), (1, 0)])
        self.assertEqual(p.skipped, 1)                                    # (1, 1) is empty: not counted

    def test_arithmetic_errors_are_reported_too(self):
        p = plan([(1.0,)], [("1",)], accepts_numbers, lambda v: v / 0)
        self.assertEqual(len(p.items), 1)
        self.assertIsNotNone(p.items[0].error)


if __name__ == "__main__":
    unittest.main()
