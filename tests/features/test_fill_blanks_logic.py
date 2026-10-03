import unittest

import tests  # noqa: F401
from office_plugin.features.calc_fill_blanks.fill import count_cells, plan_fill
from office_plugin.services.cellkeys import is_blank


class FillDownTests(unittest.TestCase):
    def test_basic_runs(self):
        grid = [("North", 1.0), ("", 2.0), ("", 3.0), ("South", 4.0), ("", 5.0)]
        self.assertEqual(plan_fill(grid), [(1, 0, 2, 0, "North"), (4, 0, 4, 0, "South")])

    def test_leading_blanks_stay_blank(self):
        grid = [("",), ("",), ("A",), ("",)]
        self.assertEqual(plan_fill(grid), [(3, 0, 3, 0, "A")])

    def test_each_column_independent_and_numbers_kept_as_numbers(self):
        grid = [(5.0, "x"), ("", ""), ("", "y"), (7.0, "")]
        runs = plan_fill(grid)
        self.assertIn((1, 0, 2, 0, 5.0), runs)
        self.assertIn((1, 1, 1, 1, "x"), runs)
        self.assertIn((3, 1, 3, 1, "y"), runs)
        self.assertEqual(count_cells(runs), 4)

    def test_zero_is_a_value_not_blank(self):
        self.assertEqual(plan_fill([(0.0,), ("",)]), [(1, 0, 1, 0, 0.0)])

    def test_whitespace_predicate(self):
        grid = [("A",), (" ",), (" ",), ("B",)]
        self.assertEqual(plan_fill(grid), [])                      # whitespace is a value by default
        runs = plan_fill(grid, is_blank=lambda v: is_blank(v, whitespace=True))
        self.assertEqual(runs, [(1, 0, 2, 0, "A")])

    def test_ragged_rows_and_empty(self):
        self.assertEqual(plan_fill([]), [])
        self.assertEqual(plan_fill([("A", "B"), ("",)]),                # missing cell = blank
                         [(1, 0, 1, 0, "A"), (1, 1, 1, 1, "B")])


class FillRightTests(unittest.TestCase):
    def test_header_row(self):
        grid = [("2026", "", "", "2027", "", ""), ("a", "b", "c", "d", "e", "f")]
        self.assertEqual(plan_fill(grid, "right"), [(0, 1, 0, 2, "2026"), (0, 4, 0, 5, "2027")])

    def test_leading_blank_stays(self):
        self.assertEqual(plan_fill([("", "", "X", "")], "right"), [(0, 3, 0, 3, "X")])

    def test_bad_direction(self):
        with self.assertRaises(ValueError):
            plan_fill([("a",)], "left")


if __name__ == "__main__":
    unittest.main()
