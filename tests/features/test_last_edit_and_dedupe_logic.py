import unittest

import tests  # noqa: F401
from office_plugin.core.position import Position
from office_plugin.features.calc_dedupe.dedupe import col_name, find_duplicate_rows, group_runs
from office_plugin.features.last_edit.ring import EditRing
from office_plugin.services.cellkeys import EXACT, CleanOptions


def wr(y, doc="d"):
    return Position("writer", doc, container=0, coords=(y,))


class EditRingTests(unittest.TestCase):
    def test_keeps_last_three_distinct_places(self):
        r = EditRing(size=3)
        for y in (0, 100000, 200000, 300000):
            r.record(wr(y))
        self.assertEqual(len(r), 3)
        self.assertEqual(r.next_target().coords, (300000,))

    def test_nearby_edits_merge(self):
        r = EditRing()
        r.record(wr(1000))
        r.record(wr(1100))                      # same spot (typing)
        self.assertEqual(len(r), 1)
        self.assertEqual(r.next_target().coords, (1100,))

    def test_cycle_newest_first_and_wraps(self):
        r = EditRing()
        for y in (0, 100000, 200000):
            r.record(wr(y))
        seen = [r.next_target(wr(999999)).coords[0] for _ in range(4)]
        self.assertEqual(seen, [200000, 100000, 0, 200000])

    def test_first_press_skips_where_you_already_are(self):
        r = EditRing()
        r.record(wr(0))
        r.record(wr(100000))
        self.assertEqual(r.next_target(wr(100000)).coords, (0,))

    def test_new_edit_resets_cycle_and_drop_doc(self):
        r = EditRing()
        r.record(wr(0))
        r.record(wr(100000))
        r.next_target()
        r.next_target()
        r.record(wr(300000))
        self.assertEqual(r.next_target(wr(999999)).coords, (300000,))
        r.drop_doc("d")
        self.assertIsNone(r.next_target())


class DedupeTests(unittest.TestCase):
    ROWS = [("Name", "City"), ("Ann", "Pune"), ("ann ", "Pune"), ("Bob", "Pune"),
            ("Ann", "Delhi"), ("", ""), ("", ""), ("Bob", "Pune")]

    def test_all_columns_ignore_case_and_whitespace(self):
        self.assertEqual(find_duplicate_rows(self.ROWS, [0, 1]), [2, 7])

    def test_single_column(self):
        self.assertEqual(find_duplicate_rows(self.ROWS, [0]), [2, 4, 7])

    def test_case_sensitive(self):
        rows = [("h",), ("A",), ("a",)]
        self.assertEqual(find_duplicate_rows(rows, [0], opts=CleanOptions(True, False, True, True)), [])
        self.assertEqual(find_duplicate_rows(rows, [0], opts=CleanOptions(True, True, True, True)), [2])

    def test_shared_cleaning_rules(self):
        rows = [("h",), ("101",), (101.0,), ("a b",), ("a b",), (" a  b ",), ("x​y",), ("xy",),
                ("  ",), ("",)]
        self.assertEqual(find_duplicate_rows(rows, [0]), [2, 4, 5, 7])     # blank keys never match
        self.assertEqual(find_duplicate_rows(rows, [0], opts=EXACT), [])
        self.assertEqual(find_duplicate_rows(rows, [0], opts=CleanOptions(True, True, False, True)), [4, 5, 7])
        self.assertEqual(find_duplicate_rows(rows, [0], opts=CleanOptions(True, True, True, False)), [2, 4, 5])

    def test_header_flag(self):
        rows = [("x",), ("x",), ("y",)]
        self.assertEqual(find_duplicate_rows(rows, [0], has_header=True), [])
        self.assertEqual(find_duplicate_rows(rows, [0], has_header=False), [1])

    def test_numbers_and_blank_keys(self):
        rows = [(1.0, "a"), (1.0, "b"), (2.0, ""), ("", "")]
        self.assertEqual(find_duplicate_rows(rows, [0], has_header=False), [1])
        self.assertEqual(find_duplicate_rows(rows, [], has_header=False), [])

    def test_group_runs_and_col_name(self):
        self.assertEqual(group_runs([2, 3, 4, 9, 11, 12]), [(2, 4), (9, 9), (11, 12)])
        self.assertEqual(group_runs([]), [])
        self.assertEqual([col_name(i) for i in (0, 25, 26, 51, 52, 701, 702)],
                         ["A", "Z", "AA", "AZ", "BA", "ZZ", "AAA"])


if __name__ == "__main__":
    unittest.main()
