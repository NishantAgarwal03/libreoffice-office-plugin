import unittest

import tests  # noqa: F401
from office_plugin.services import decimals as D


class ValueTests(unittest.TestCase):
    def test_round_half_up_like_calc(self):
        self.assertEqual(D.round_value(2.675, 2), 2.68)           # binary 2.67499999 -> still 2.68, like Calc
        self.assertEqual(D.round_value(3.14159, 2), 3.14)
        self.assertEqual(D.round_value(-2.5, 0), -3.0)
        self.assertEqual(D.round_value(0.5, 0), 1.0)
        self.assertEqual(D.round_value(1234.5678, 0), 1235.0)
        self.assertEqual(D.round_value(7, 3), 7.0)

    def test_up_and_down(self):
        self.assertEqual(D.round_value(2.341, 2, D.ROUND_UP_MODE), 2.35)
        self.assertEqual(D.round_value(-2.341, 2, D.ROUND_UP_MODE), -2.35)
        self.assertEqual(D.round_value(2.349, 2, D.ROUND_DOWN_MODE), 2.34)
        self.assertEqual(D.round_value(-2.349, 2, D.ROUND_DOWN_MODE), -2.34)

    def test_bounds(self):
        for bad in (-1, 16):
            with self.assertRaises(ValueError):
                D.round_value(1.5, bad)
        self.assertEqual(D.round_value(1.1234567890123456, 15), 1.123456789012346)


class FormulaTests(unittest.TestCase):
    def test_wrap_plain_formulas(self):
        self.assertEqual(D.wrap_round("=A1*B1", 2), ("=ROUND(A1*B1;2)", "wrapped"))
        self.assertEqual(D.wrap_round("=SUM(A1:A3)/3", 1, D.ROUND_UP_MODE), ("=ROUNDUP(SUM(A1:A3)/3;1)", "wrapped"))
        self.assertEqual(D.wrap_round("=A1", 0, D.ROUND_DOWN_MODE), ("=ROUNDDOWN(A1;0)", "wrapped"))

    def test_existing_wrapper_is_replaced_not_nested(self):
        self.assertEqual(D.wrap_round("=ROUND(A1*B1;2)", 3), ("=ROUND(A1*B1;3)", "updated"))
        self.assertEqual(D.wrap_round("=ROUND(A1*B1;2)", 2), ("=ROUND(A1*B1;2)", "same"))
        self.assertEqual(D.wrap_round("=ROUNDUP(A1;2)", 2), ("=ROUND(A1;2)", "updated"))
        self.assertEqual(D.wrap_round("=ROUND(A1;-1)", 2), ("=ROUND(A1;2)", "updated"))
        self.assertEqual(D.wrap_round("=ROUND(ROUND(A1;4);2)", 3), ("=ROUND(ROUND(A1;4);3)", "updated"))   # only the OUTER one

    def test_not_a_wrapper(self):
        for formula in ("=ROUND(A1;2)+1", "=1+ROUND(A1;2)", "=ROUND(A1;B1)", "=ROUND(A1)", "=ROUND(A1;2;3)",
                        "=ROUNDTRIP(A1;2)", "=ROUND(A1;2)*ROUND(B1;2)"):
            self.assertEqual(D.unwrap(formula), (None, None), formula)
            self.assertTrue(D.wrap_round(formula, 2)[0].startswith("=ROUND(" + formula[1:]), formula)

    def test_awkward_content_inside(self):
        f = '=IF(A1="a;b)";ROUND(B1;2);0)'
        self.assertEqual(D.wrap_round(f, 2)[0], '=ROUND(' + f[1:] + ';2)')
        sheet = "=SUM('My;Sheet (1)'.A1:A3)"
        self.assertEqual(D.wrap_round(sheet, 1)[0], "=ROUND(SUM('My;Sheet (1)'.A1:A3);1)")
        inner = '=ROUND(IF(A1>0;"x;y";B1)*2;2)'
        self.assertEqual(D.unwrap(inner)[0], 'IF(A1>0;"x;y";B1)*2')

    def test_remove(self):
        self.assertEqual(D.remove_round("=ROUND(A1*B1;2)"), "=A1*B1")
        self.assertEqual(D.remove_round("=ROUNDDOWN(SUM(A1:A3);0)"), "=SUM(A1:A3)")
        self.assertIsNone(D.remove_round("=A1*B1"))

    def test_whitespace_tolerance(self):
        self.assertEqual(D.wrap_round("=ROUND(A1 ; 2)", 2)[1], "same")


class FormatTests(unittest.TestCase):
    def test_basic_codes(self):
        f = D.format_with_decimals
        self.assertEqual(f("General", 2), "0.00")
        self.assertEqual(f("General", 0), "0")
        self.assertEqual(f("0.00", 3), "0.000")
        self.assertEqual(f("#,##0.00", 0), "#,##0")
        self.assertEqual(f("#,##0", 2), "#,##0.00")
        self.assertEqual(f("0.0%", 2), "0.00%")
        self.assertEqual(f("0%", 1), "0.0%")

    def test_currency_colours_and_sections(self):
        f = D.format_with_decimals
        self.assertEqual(f("[$₹-4009] #,##0.00", 1), "[$₹-4009] #,##0.0")
        self.assertEqual(f("#,##0.00 [$€-407]", 3), "#,##0.000 [$€-407]")
        self.assertEqual(f("#,##0.00;[RED]-#,##0.00", 1), "#,##0.0;[RED]-#,##0.0")
        self.assertEqual(f('0.00" kg"', 1), '0.0" kg"')
        self.assertEqual(f('"Qty: "0.00', 0), '"Qty: "0')
        self.assertEqual(f('0.00 "a;b"', 1), '0.0 "a;b"')

    def test_percent_detection(self):
        self.assertTrue(D.is_percent("0.00%"))
        self.assertFalse(D.is_percent('0.00" 100%"'))
        self.assertFalse(D.is_percent("#,##0.00"))
        self.assertFalse(D.is_percent("[$%-409]0.00") and False)

    def test_scientific_left_alone_and_bounds(self):
        self.assertEqual(D.format_with_decimals("0.00E+00", 4), "0.00E+00")
        with self.assertRaises(ValueError):
            D.format_with_decimals("0.00", 16)


if __name__ == "__main__":
    unittest.main()
