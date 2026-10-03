import unittest

import tests  # noqa: F401
from office_plugin.features.calc_remove_blanks import blanks as B

T, F = True, False


class MaskTests(unittest.TestCase):
    MASK = [[F, F, T],
            [T, T, T],      # blank row
            [F, T, F],
            [T, T, T],      # blank row
            [T, T, T]]      # blank row

    def test_blank_rows(self):
        self.assertEqual(B.blank_rows(self.MASK), [1, 3, 4])

    def test_blank_cols(self):
        self.assertEqual(B.blank_cols(self.MASK), [])
        self.assertEqual(B.blank_cols([[F, T, F], [F, T, T]]), [1])

    def test_empty_and_ragged(self):
        self.assertEqual(B.blank_rows([]), [])
        self.assertEqual(B.blank_cols([]), [])
        self.assertEqual(B.blank_rows([[]]), [])              # a zero-width row is not "blank"
        self.assertEqual(B.blank_cols([[F, T], [F]]), [1])    # missing trailing cell = blank

    def test_runs(self):
        self.assertEqual(B.runs([5, 3, 4, 9]), [(3, 5), (9, 9)])

    def test_cell_runs_up_per_column(self):
        mask = [[F, T], [T, T], [T, F], [F, F]]
        self.assertEqual(B.blank_cell_runs(mask, "up"), [(0, 1, 2), (1, 0, 1)])

    def test_cell_runs_left_per_row(self):
        mask = [[F, T, T, F, T], [T, F, F, F, F]]
        self.assertEqual(B.blank_cell_runs(mask, "left"), [(0, 1, 2), (0, 4, 4), (1, 0, 0)])

    def test_count_and_bad_axis(self):
        self.assertEqual(B.count_blank_cells(self.MASK), 11)
        with self.assertRaises(ValueError):
            B.blank_cell_runs(self.MASK, "diagonal")


if __name__ == "__main__":
    unittest.main()
