import unittest

import tests  # noqa: F401
from office_plugin.features.calc_transpose import transpose as T
from office_plugin.features.calc_transpose.transpose import Rect


def cells(rects):
    return {(r, c) for x in rects for r in range(x.row1, x.row2 + 1) for c in range(x.col1, x.col2 + 1)}


class GeometryTests(unittest.TestCase):
    SRC = Rect(0, 0, 2, 1)          # A1:B3  (3 rows x 2 cols)

    def test_transposed_rect_swaps_shape(self):
        self.assertEqual(T.transposed_rect(self.SRC, 0, 0), Rect(0, 0, 1, 2))      # A1:C2
        self.assertEqual(T.transposed_rect(self.SRC, 5, 10), Rect(5, 10, 6, 12))
        self.assertEqual(T.cell_count(T.transposed_rect(self.SRC, 0, 0)), 6)

    def test_degenerate_shapes(self):
        row = Rect(0, 0, 0, 4)                                  # 1 x 5
        self.assertEqual(T.transposed_rect(row, 0, 0), Rect(0, 0, 4, 0))             # 5 x 1
        self.assertEqual(T.transposed_rect(Rect(2, 2, 2, 2), 2, 2), Rect(2, 2, 2, 2))

    def test_square_in_place_overwrites_nothing_new(self):
        sq = Rect(0, 0, 2, 2)
        self.assertEqual(T.rect_difference(T.transposed_rect(sq, 0, 0), sq), [])

    def test_in_place_nonsquare_difference(self):
        dest = T.transposed_rect(self.SRC, 0, 0)                 # A1:C2 vs source A1:B3
        self.assertEqual(cells(T.rect_difference(dest, self.SRC)), {(0, 2), (1, 2)})   # only column C

    def test_difference_properties(self):
        a = Rect(1, 1, 6, 6)
        for b in (Rect(0, 0, 3, 3), Rect(2, 2, 3, 3), Rect(5, 0, 9, 9), Rect(7, 7, 8, 8), Rect(1, 1, 6, 6)):
            parts = T.rect_difference(a, b)
            expected = cells([a]) - cells([b])
            self.assertEqual(cells(parts), expected, b)
            self.assertEqual(sum(T.cell_count(p) for p in parts), len(expected), "parts overlap: %s" % (b,))

    def test_intersects_and_fits(self):
        self.assertTrue(T.intersects(Rect(0, 0, 2, 2), Rect(2, 2, 4, 4)))
        self.assertFalse(T.intersects(Rect(0, 0, 1, 1), Rect(2, 2, 4, 4)))
        self.assertTrue(T.fits(Rect(0, 0, 9, 9), 9, 9))
        self.assertFalse(T.fits(Rect(0, 0, 10, 0), 9, 9))
        self.assertFalse(T.fits(Rect(0, -1, 0, 0), 9, 9))

    def test_swap_grid_and_describe(self):
        self.assertEqual(T.swap_grid([("a", 1), ("b", 2), ("c", 3)]), [("a", "b", "c"), (1, 2, 3)])
        self.assertEqual(T.swap_grid([]), [])
        self.assertEqual(T.swap_grid([(1, 2, 3)]), [(1,), (2,), (3,)])
        self.assertIn("3 row(s) x 2 column(s)  ->  2 row(s) x 3 column(s)", T.describe(self.SRC))


if __name__ == "__main__":
    unittest.main()
