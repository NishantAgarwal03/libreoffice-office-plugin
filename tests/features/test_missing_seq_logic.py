import unittest

import tests  # noqa: F401
from office_plugin.features.calc_missing_seq import missing as M


def numbers(result):
    return [(g.before_row, M.gap_numbers(g)) for g in result.gaps]


class ParseTests(unittest.TestCase):
    def test_numbers_and_text(self):
        self.assertEqual(M.parse(7.0, 0).number, 7)
        self.assertIsNone(M.parse(7.5, 0))
        self.assertIsNone(M.parse("", 0))
        self.assertIsNone(M.parse("abc", 0))
        inv = M.parse("INV-0012", 3)
        self.assertEqual((inv.number, inv.series), (12, M.Series("INV-", "", False, 4)))
        grn = M.parse("GRN/045/26", 0)
        self.assertEqual((grn.number, grn.series.prefix, grn.series.suffix), (26, "GRN/045/", ""))   # trailing number
        self.assertEqual(M.parse("12-A", 0).series.suffix, "-A")
        self.assertEqual(M.parse("007", 0).series.width, 3)
        self.assertEqual(M.parse("10", 0).series.width, 0)

    def test_format_number(self):
        self.assertEqual(M.format_number(M.Series("INV-", "", False, 4), 13), "INV-0013")
        self.assertEqual(M.format_number(M.Series("", "", True, 0), 13), 13.0)
        self.assertEqual(M.format_number(M.Series("A", "x", False, 0), 9), "A9x")


class GapTests(unittest.TestCase):
    def test_simple_gaps(self):
        r = M.find_gaps([1.0, 2.0, 5.0, 6.0, 9.0])
        self.assertEqual(numbers(r), [(2, [3, 4]), (4, [7, 8])])
        self.assertEqual(r.total_missing, 4)

    def test_no_gap_and_empty(self):
        self.assertEqual(M.find_gaps([1.0, 2.0, 3.0]).gaps, [])
        self.assertEqual(M.find_gaps([]).gaps, [])
        self.assertEqual(M.find_gaps(["", "x"]).skipped, 1)

    def test_duplicates_and_out_of_order_do_not_create_false_gaps(self):
        r = M.find_gaps([1.0, 2.0, 2.0, 5.0, 4.0, 6.0])
        self.assertEqual(numbers(r), [(3, [3, 4])])                  # 4 appears later, but order is the order
        self.assertEqual((r.duplicates, r.out_of_order), (1, 1))
        # running extreme: after 5, the out-of-order 4 does not make 6 look like a gap
        self.assertEqual(M.find_gaps([1.0, 5.0, 4.0, 6.0]).total_missing, 3)

    def test_blank_and_text_cells_are_skipped_not_gaps(self):
        r = M.find_gaps([1.0, "", "n/a", 2.0, 3.0])
        self.assertEqual(r.gaps, [])
        self.assertEqual(r.skipped, 1)                                # "n/a"; the blank is not counted

    def test_text_with_prefix_and_padding(self):
        r = M.find_gaps(["INV-0001", "INV-0002", "INV-0005"])
        self.assertEqual(numbers(r), [(2, [3, 4])])
        self.assertEqual([M.format_number(r.gaps[0].series, n) for n in M.gap_numbers(r.gaps[0])],
                         ["INV-0003", "INV-0004"])
        self.assertEqual(M.label(r.gaps[0]), "INV-0003-INV-0004")

    def test_prefix_change_starts_a_new_series(self):
        r = M.find_gaps(["A-1", "A-2", "B-10", "B-12"])
        self.assertEqual(numbers(r), [(3, [11])])                    # nothing between A-2 and B-10

    def test_step_and_off_step(self):
        r = M.find_gaps([5.0, 10.0, 20.0], step=5)
        self.assertEqual(numbers(r), [(2, [15])])
        r = M.find_gaps([5.0, 10.0, 12.0], step=5)
        self.assertEqual((r.gaps, r.off_step), ([], 1))              # 12 is not on the 5-step grid
        with self.assertRaises(ValueError):
            M.find_gaps([1.0], step=0)
        with self.assertRaises(ValueError):
            M.find_gaps([1.0], step=1.5)

    def test_descending(self):
        r = M.find_gaps([10.0, 9.0, 6.0, 5.0], descending=True)
        self.assertEqual(numbers(r), [(2, [8, 7])])

    def test_huge_gap_is_counted_without_building_a_list(self):
        r = M.find_gaps([1.0, 1000000000.0])
        self.assertEqual(r.total_missing, 999999998)
        self.assertEqual(r.gaps[0].count, 999999998)

    def test_labels(self):
        r = M.find_gaps([1.0, 2.0, 5.0, 7.0, 20.0, 22.0])
        self.assertEqual(M.describe(r), "3-4, 6, 8-19, 21")
        many = M.find_gaps([float(i * 2) for i in range(12)])
        self.assertIn("(3 more gap(s))", M.describe(many))


if __name__ == "__main__":
    unittest.main()
