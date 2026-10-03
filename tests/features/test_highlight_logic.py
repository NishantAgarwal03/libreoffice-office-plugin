import unittest

import tests  # noqa: F401
from office_plugin.features.calc_highlight import marks as M
from office_plugin.services.cellkeys import CleanOptions, EXACT

ROWS = [("A",), ("b",), ("a ",), ("C",), ("B",), ("a",), ("",), ("",)]


class DuplicateTests(unittest.TestCase):
    def test_every_occurrence(self):
        self.assertEqual(M.duplicates(ROWS, [0]), {0, 1, 2, 4, 5})      # A/a/a, b/B ; C unique

    def test_repeats_only_keeps_first_occurrence_unmarked(self):
        self.assertEqual(M.duplicates(ROWS, [0], repeats_only=True), {2, 4, 5})

    def test_unique(self):
        self.assertEqual(M.unique(ROWS, [0]), {3})

    def test_blank_keys_never_marked(self):
        self.assertNotIn(6, M.duplicates(ROWS, [0]))
        self.assertNotIn(6, M.unique(ROWS, [0]))

    def test_cleaning_changes_result(self):
        self.assertEqual(M.duplicates(ROWS, [0], EXACT), set())           # exact: nothing equal
        exact_case = CleanOptions(True, False, True, True)
        self.assertEqual(M.duplicates(ROWS, [0], exact_case), {2, 5})     # trim only: "a " == "a", "A" differs

    def test_multi_column_key(self):
        rows = [("x", 1.0), ("x", 2.0), ("x", 1.0)]
        self.assertEqual(M.duplicates(rows, [0, 1]), {0, 2})
        self.assertEqual(M.duplicates(rows, [0]), {0, 1, 2})


class MembershipTests(unittest.TestCase):
    A = [("T-01",), ("T-02",), ("T-03",), ("",)]
    B = [("t-02 ",), ("T-04",), ("101",)]

    def test_present_and_absent(self):
        self.assertEqual(M.membership(self.A, self.B, [0], present=True), ({1}, {0}))
        self.assertEqual(M.membership(self.A, self.B, [0], present=False), ({0, 2}, {1, 2}))

    def test_numeric_text(self):
        self.assertEqual(M.membership([(101.0,)], [("101",)], [0])[0], {0})


class CellDiffTests(unittest.TestCase):
    def test_differences(self):
        a = [("x", 1.0), ("y", 2.0)]
        b = [("X ", "1"), ("z", 2.0)]
        self.assertEqual(M.cell_diffs(a, b), {(1, 0)})                   # cleaned: only y vs z differs
        self.assertEqual(M.cell_diffs(a, b, EXACT), {(0, 0), (0, 1), (1, 0)})

    def test_shape_mismatch(self):
        with self.assertRaises(ValueError):
            M.cell_diffs([("a",)], [("a",), ("b",)])
        with self.assertRaises(ValueError):
            M.cell_diffs([("a", "b")], [("a",)])


class ConstantsTests(unittest.TestCase):
    def test_every_mode_has_a_suggested_colour_in_the_palette(self):
        for mode in M.MODES:
            self.assertIn(M.SUGGESTED_COLOUR[mode], M.PALETTE)
        self.assertEqual(len(set(M.PALETTE.values())), len(M.PALETTE))


if __name__ == "__main__":
    unittest.main()
