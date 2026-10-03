import unittest
from decimal import Decimal as D

import tests  # noqa: F401
from office_plugin.services.numunits import parse_units


class ParseUnitsTests(unittest.TestCase):
    def test_the_users_examples(self):
        for text, value in (("1 lakh", 100000), ("1L", 100000), ("2 cr", 20000000), ("2cr", 20000000),
                            ("1 lac", 100000), ("2 crore", 20000000), ("2 Crores", 20000000), ("3 lakhs", 300000)):
            self.assertEqual(parse_units(text), D(value), text)

    def test_decimals_and_international(self):
        self.assertEqual(parse_units("1.5k"), D(1500))
        self.assertEqual(parse_units("2.5L"), D(250000))
        self.assertEqual(parse_units("0.75 cr"), D(7500000))
        self.assertEqual(parse_units("3M"), D(3000000))
        self.assertEqual(parse_units("2.3 mn"), D(2300000))
        self.assertEqual(parse_units("1 bn"), D(10 ** 9))
        self.assertEqual(parse_units("1 arab"), D(10 ** 9))
        self.assertEqual(parse_units("1 kharab"), D(10 ** 11))
        self.assertEqual(parse_units("1 trillion"), D(10 ** 12))
        self.assertEqual(parse_units(".5k"), D(500))

    def test_exact_arithmetic(self):
        self.assertEqual(parse_units("0.1L"), D(10000))
        self.assertEqual(parse_units("1.1 cr"), D(11000000))
        self.assertEqual(str(parse_units("2.55k")), "2550.00")

    def test_plain_numbers_and_grouping(self):
        self.assertEqual(parse_units("1500"), D(1500))
        self.assertEqual(parse_units("12,34,567"), D(1234567))          # Indian grouping
        self.assertEqual(parse_units("1,234,567.5"), D("1234567.5"))
        self.assertEqual(parse_units("1,25,000 "), D(125000))

    def test_currency_and_sign(self):
        self.assertEqual(parse_units("₹ 1,25,000"), D(125000))
        self.assertEqual(parse_units("Rs. 2 cr"), D(20000000))
        self.assertEqual(parse_units("Rs 5L"), D(500000))
        self.assertEqual(parse_units("$2.5M"), D(2500000))
        self.assertEqual(parse_units("-2 lakh"), D(-200000))
        self.assertEqual(parse_units("(2.5L)"), D(-250000))
        self.assertEqual(parse_units("+3k"), D(3000))

    def test_compound_amounts(self):
        self.assertEqual(parse_units("1 crore 25 lakh"), D(12500000))
        self.assertEqual(parse_units("2 lakh 50 thousand"), D(250000))
        self.assertEqual(parse_units("2L, 50K"), D(250000))
        self.assertEqual(parse_units("1 cr and 5 lakh"), D(10500000))

    def test_second_unit_multiplies(self):
        self.assertEqual(parse_units("1 lakh crore"), D(10 ** 12))
        self.assertEqual(parse_units("2.5 lakh cr"), D(25 * 10 ** 11))
        self.assertEqual(parse_units("2 L cr"), D(2 * 10 ** 12))
        self.assertEqual(parse_units("1 cr 25 lakh"), D(12500000))           # a number in between: still a sum
        self.assertEqual(parse_units("3 lakh and 50 thousand"), D(350000))

    def test_spacing_and_case(self):
        self.assertEqual(parse_units("  2   CR  "), D(20000000))
        self.assertEqual(parse_units("2Cr."), D(20000000))

    def test_rejects_garbage(self):
        for bad in ("", "   ", "abc", "lakh", "2 xyz", "1,5L", "1..5k", "2 cr 3 qq", "₹", "--5"):
            with self.subTest(bad), self.assertRaises(ValueError):
                parse_units(bad)
        with self.assertRaises(ValueError):
            parse_units(None)


if __name__ == "__main__":
    unittest.main()
