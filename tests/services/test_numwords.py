import random
import unittest
from decimal import Decimal as D

import tests  # noqa: F401
from office_plugin.services import numwords as N


class ToWordsTests(unittest.TestCase):
    def test_small_numbers(self):
        cases = {0: "zero", 5: "five", 13: "thirteen", 20: "twenty", 21: "twenty-one", 99: "ninety-nine",
                 100: "one hundred", 101: "one hundred one", 999: "nine hundred ninety-nine"}
        for number, words in cases.items():
            self.assertEqual(N.to_words(number, case=N.LOWER), words, number)

    def test_indian_grouping(self):
        t = lambda n: N.to_words(n, N.INDIAN, case=N.LOWER)
        self.assertEqual(t(1000), "one thousand")
        self.assertEqual(t(100000), "one lakh")
        self.assertEqual(t(123456), "one lakh twenty-three thousand four hundred fifty-six")
        self.assertEqual(t(1234567), "twelve lakh thirty-four thousand five hundred sixty-seven")
        self.assertEqual(t(10000000), "one crore")
        self.assertEqual(t(250000000), "twenty-five crore")
        self.assertEqual(t(1000000000), "one arab")
        self.assertEqual(t(10 ** 13), "one neel")
        self.assertEqual(t(100100100), "ten crore one lakh one hundred")

    def test_international_grouping(self):
        t = lambda n: N.to_words(n, N.INTERNATIONAL, case=N.LOWER)
        self.assertEqual(t(1000000), "one million")
        self.assertEqual(t(1234567), "one million two hundred thirty-four thousand five hundred sixty-seven")
        self.assertEqual(t(5000000000), "five billion")
        self.assertEqual(t(10 ** 12), "one trillion")

    def test_and_option(self):
        self.assertEqual(N.to_words(105, N.INTERNATIONAL, case=N.LOWER, use_and=True), "one hundred and five")
        self.assertEqual(N.to_words(1005, N.INTERNATIONAL, case=N.LOWER, use_and=True), "one thousand and five")
        self.assertEqual(N.to_words(123456, N.INDIAN, case=N.LOWER, use_and=True),
                         "one lakh twenty-three thousand four hundred and fifty-six")

    def test_negative_and_decimals(self):
        self.assertEqual(N.to_words(-42, case=N.LOWER), "minus forty-two")
        self.assertEqual(N.to_words(12.45, case=N.LOWER), "twelve point four five")
        self.assertEqual(N.to_words(0.05, case=N.LOWER), "zero point zero five")
        self.assertEqual(N.to_words(D("7.50"), case=N.LOWER), "seven point five")
        self.assertEqual(N.to_words(-0.0, case=N.LOWER), "zero")

    def test_case_styles(self):
        self.assertEqual(N.to_words(123456, case=N.SENTENCE), "One lakh twenty-three thousand four hundred fifty-six")
        self.assertEqual(N.to_words(23, case=N.TITLE), "Twenty-Three")
        self.assertEqual(N.to_words(23, case=N.UPPER), "TWENTY-THREE")

    def test_currency(self):
        t = lambda v, **k: N.to_words(v, currency="Rupees / Paise", case=N.SENTENCE, **k)
        self.assertEqual(t(123456.78), "Rupees one lakh twenty-three thousand four hundred fifty-six and Paise seventy-eight")
        self.assertEqual(t(5, only=True), "Rupees five only")
        self.assertEqual(t(0.5), "Rupees zero and Paise fifty")
        self.assertEqual(t(2.995), "Rupees three")                  # rounds to 3.00
        self.assertEqual(N.to_words(1.5, N.INTERNATIONAL, "Dollars / Cents", N.TITLE, only=True),
                         "Dollars One And Cents Fifty Only")

    def test_floating_point_noise_reads_like_the_cell_shows(self):
        self.assertEqual(N.to_words(1.1 + 2.2, case=N.LOWER), "three point three")        # stored 3.3000000000000003
        self.assertEqual(N.to_words(0.1 + 0.2, case=N.LOWER), "zero point three")
        self.assertEqual(N.to_words(100 * 1.1, case=N.LOWER), "one hundred ten")
        self.assertEqual(N.to_words(1234.5678, case=N.LOWER), "one thousand two hundred thirty-four point five six seven eight")
        self.assertEqual(N.to_words(1.1 + 2.2, currency="Rupees / Paise"), "Rupees three and Paise thirty")

    def test_limits(self):
        with self.assertRaises(ValueError):
            N.to_words(10 ** 15)
        with self.assertRaises(ValueError):
            N.to_words(float("nan"))
        with self.assertRaises(ValueError):
            N.to_words(True)
        self.assertTrue(N.to_words(10 ** 15 - 1).startswith("Ninety-nine neel"))


class ParseWordsTests(unittest.TestCase):
    def test_both_systems(self):
        p = N.parse_words
        self.assertEqual(p("one lakh twenty-three thousand four hundred fifty-six"), D(123456))
        self.assertEqual(p("twelve crore"), D(120000000))
        self.assertEqual(p("two million five hundred thousand"), D(2500000))
        self.assertEqual(p("Five Hundred and Fifty"), D(550))
        self.assertEqual(p("one lac"), D(100000))
        self.assertEqual(p("3 lakhs".replace("3", "three")), D(300000))

    def test_consecutive_scale_words_multiply(self):
        p = N.parse_words
        self.assertEqual(p("one lakh crore"), D(10 ** 12))                  # was silently read as 10,100,000
        self.assertEqual(p("two lakh crore"), D(2 * 10 ** 12))
        self.assertEqual(p("twenty-five lakh crore"), D(25 * 10 ** 12))
        self.assertEqual(p("one lakh crore fifty lakh"), D(10 ** 12 + 5 * 10 ** 6))
        self.assertEqual(p("five lakh fifty thousand"), D(550000))         # ordinary sequences unchanged
        self.assertEqual(p("one crore one lakh"), D(10100000))

    def test_negative_decimal_and_currency(self):
        p = N.parse_words
        self.assertEqual(p("minus forty-two"), D(-42))
        self.assertEqual(p("twelve point four five"), D("12.45"))
        self.assertEqual(p("point five"), D("0.5"))
        self.assertEqual(p("Rupees one lakh and Paise fifty Only"), D("100000.50"))
        self.assertEqual(p("twelve rupees and fifty paise"), D("12.50"))
        self.assertEqual(p("rupees zero and paise fifty"), D("0.50"))
        self.assertEqual(p("fifty paise"), D("0.50"))
        self.assertEqual(p("five dollars and twenty cents only"), D("5.20"))

    def test_errors(self):
        for bad in ("", "   ", "banana", "twenty bananas", "point", "one point two x", "rupees and paise"):
            with self.subTest(bad), self.assertRaises(ValueError):
                N.parse_words(bad)
        with self.assertRaises(ValueError):
            N.parse_words("five rupees and one hundred paise")

    def test_round_trip_property(self):
        rng = random.Random(7)
        values = [0, 7, 19, 20, 99, 100, 999, 1000, 12345, 100000, 999999, 10000000, 123456789, 10 ** 12 + 5,
                  10 ** 14 + 12345] + [rng.randrange(0, 10 ** rng.randrange(1, 15)) for _ in range(300)]
        for system in (N.INDIAN, N.INTERNATIONAL):
            for v in values:
                for use_and in (False, True):
                    words = N.to_words(v, system, use_and=use_and)
                    self.assertEqual(N.parse_words(words), D(v), (system, v, words))

    def test_round_trip_decimals_and_currency(self):
        for v in ("0.5", "12.45", "100.01", "0.07", "99999.99"):
            self.assertEqual(N.parse_words(N.to_words(D(v), N.INDIAN, "Rupees / Paise")), D(v), v)
        for v in ("12.45", "0.05", "1234567.891"):
            self.assertEqual(N.parse_words(N.to_words(D(v))), D(v), v)
        self.assertEqual(N.parse_words(N.to_words(-250.5, N.INDIAN, "Rupees / Paise")), D("-250.50"))


if __name__ == "__main__":
    unittest.main()
