import unittest

import tests  # noqa: F401
from office_plugin.services import calc_ranges as R
from office_plugin.services.cellkeys import (EXACT, CleanOptions, header_guess, is_blank,
                                             key_is_blank, normalize, row_key)


class NormalizeTests(unittest.TestCase):
    def test_whitespace_and_case(self):
        self.assertEqual(normalize("  Ceramic   Tile "), normalize("ceramic tile"))
        self.assertEqual(normalize("A B"), normalize("a b"))               # non-breaking space
        self.assertNotEqual(normalize("ceramic tile", CleanOptions(False, False, True, True)),
                            normalize("Ceramic Tile", CleanOptions(False, False, True, True)))

    def test_trim_off_keeps_spaces(self):
        opts = CleanOptions(trim=False, ignore_case=True, numeric=True, strip_nonprint=True)
        self.assertNotEqual(normalize("a ", opts), normalize("a", opts))
        self.assertEqual(normalize("a ", opts), normalize("A ", opts))

    def test_invisible_characters(self):
        self.assertEqual(normalize("ab​c"), normalize("abc"))             # zero-width space
        self.assertEqual(normalize("ab\x07c"), normalize("abc"))               # control char
        keep = CleanOptions(True, True, True, strip_nonprint=False)
        self.assertNotEqual(normalize("ab​c", keep), normalize("abc", keep))

    def test_numeric_text_equals_number(self):
        self.assertEqual(normalize("101"), normalize(101.0))
        self.assertEqual(normalize("101.0"), normalize(101))
        self.assertEqual(normalize(" 2.50 "), normalize(2.5))
        self.assertNotEqual(normalize("101", EXACT), normalize(101.0, EXACT))
        self.assertNotEqual(normalize("101", CleanOptions(True, True, False, True)), normalize(101))

    def test_non_numbers_stay_text(self):
        for text in ("1,5", "nan", "inf", "1e3", "12a", "--1", ""):
            self.assertEqual(normalize(text)[0], "s", text)

    def test_blank_forms_are_equal(self):
        self.assertEqual(normalize(None), normalize(""))
        self.assertEqual(normalize("   "), normalize(""))

    def test_row_key_columns_and_short_rows(self):
        self.assertEqual(row_key(("A", " x ", 1.0), [1, 2]), row_key(("zz", "X", "1"), [1, 2]))
        self.assertEqual(row_key(("A",), [0, 3]), (("s", "a"), ("s", "")))


class BlankTests(unittest.TestCase):
    def test_is_blank(self):
        self.assertTrue(is_blank("") and is_blank(None))
        self.assertFalse(is_blank(0.0))
        self.assertFalse(is_blank("  "))
        self.assertTrue(is_blank("  ", whitespace=True))
        self.assertTrue(is_blank(" ​ ", whitespace=True))
        self.assertFalse(is_blank("x", whitespace=True))

    def test_key_is_blank(self):
        self.assertTrue(key_is_blank(row_key(("", " "), [0, 1])))
        self.assertFalse(key_is_blank(row_key(("", 0.0), [0, 1])))


class HeaderAndLabelTests(unittest.TestCase):
    def test_header_guess(self):
        self.assertTrue(header_guess(("Name", "City", "")))
        self.assertFalse(header_guess(("Name", 5.0)))
        self.assertFalse(header_guess(("", "")))

    def test_column_labels(self):
        self.assertEqual(R.column_labels(("Item", "Qty"), 1, True), ["B - Item", "C - Qty"])
        self.assertEqual(R.column_labels(("Item", ""), 0, True), ["A - Item", "Column B"])
        self.assertEqual(R.column_labels(("Item", "Qty"), 0, False), ["Column A", "Column B"])


class RangeHelperTests(unittest.TestCase):
    def test_col_name_and_runs(self):
        self.assertEqual([R.col_name(i) for i in (0, 25, 26, 701, 702)], ["A", "Z", "AA", "ZZ", "AAA"])
        self.assertEqual(R.group_runs([5, 3, 4, 9]), [(3, 5), (9, 9)])
        self.assertEqual(R.group_runs([]), [])

    def test_split_sheet_ref(self):
        self.assertEqual(R.split_sheet_ref("Sheet2.A1:C9"), ("Sheet2", "A1:C9"))
        self.assertEqual(R.split_sheet_ref("$Sheet2.$A$1:$C$9"), ("Sheet2", "A1:C9"))
        self.assertEqual(R.split_sheet_ref("'My Sheet'.B2"), ("My Sheet", "B2"))
        self.assertEqual(R.split_sheet_ref("'It''s'.B2"), ("It's", "B2"))
        self.assertEqual(R.split_sheet_ref("$A$1:$B$9"), (None, "A1:B9"))
        self.assertEqual(R.split_sheet_ref("A1"), (None, "A1"))

    def test_unique_name(self):
        self.assertEqual(R.unique_name(["Union"], "Intersection"), "Intersection")
        self.assertEqual(R.unique_name(["Union", "Union_2"], "Union"), "Union_3")


if __name__ == "__main__":
    unittest.main()
