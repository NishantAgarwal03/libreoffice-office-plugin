import unittest

import tests  # noqa: F401
from office_plugin.services.functions import call


class NumToWordsTests(unittest.TestCase):
    def test_defaults_are_indian_plain_sentence_case(self):
        self.assertEqual(call("NUMTOWORDS", 123456.0), "One lakh twenty-three thousand four hundred fifty-six")
        self.assertEqual(call("numtowords", 0.0), "Zero")                       # name is case-insensitive

    def test_system_currency_and_letters(self):
        self.assertEqual(call("NUMTOWORDS", 1234567.0, "International"),
                         "One million two hundred thirty-four thousand five hundred sixty-seven")
        self.assertEqual(call("NUMTOWORDS", 1234.5, "", "INR"), "Rupees one thousand two hundred thirty-four and Paise fifty only")
        self.assertEqual(call("NUMTOWORDS", 5.0, None, "usd", "upper"), "DOLLARS FIVE ONLY")
        self.assertEqual(call("NUMTOWORDS", 23.0, "intl", None, "title"), "Twenty-Three")
        self.assertEqual(call("NUMTOWORDS", 7.0, None, None, "lower"), "seven")

    def test_numeric_text_and_noise(self):
        self.assertEqual(call("NUMTOWORDS", "42"), "Forty-two")
        self.assertEqual(call("NUMTOWORDS", 1.1 + 2.2), "Three point three")

    def test_errors_become_value_errors(self):
        bad = [("NUMTOWORDS", None), ("NUMTOWORDS", ""), ("NUMTOWORDS", "abc"), ("NUMTOWORDS", True),
               ("NUMTOWORDS", 1.0, "klingon"), ("NUMTOWORDS", 1.0, None, "yen"), ("NUMTOWORDS", 1.0, None, None, "shout"),
               ("NUMTOWORDS", 1e20), ("NOPE", 1.0)]
        for args in bad:
            with self.subTest(args), self.assertRaises(ValueError):
                call(*args)


class SkippedArgumentTests(unittest.TestCase):
    def test_zero_or_none_for_an_option_means_default(self):
        # =NUMTOWORDS(5;;"usd") may reach us with the skipped option as None or 0
        for skipped in (None, 0, 0.0, ""):
            with self.subTest(skipped=skipped):
                self.assertEqual(call("NUMTOWORDS", 5.0, skipped, "usd", skipped), "Dollars five only")

    def test_error_text_is_the_cell_error_literal(self):
        from office_plugin.services.functions import ERROR_TEXT
        self.assertEqual(ERROR_TEXT, "#VALUE!")


class GeneratedBasicTests(unittest.TestCase):
    def test_no_reserved_word_parameters_and_no_assignment_to_optionals(self):
        # one compile error blanks EVERY function in the My Macros library, silently (found the hard way)
        import re
        from office_plugin.features.calc_functions import basic_source
        text = basic_source.source()
        self.assertIsNone(re.search(r"(?i)\bOptional\s+(system|name|type|error|date|string)\b", text))
        for param in re.findall(r"(?i)Optional\s+(\w+)\s+As", text):
            self.assertIsNone(re.search(r"(?im)^\s*%s\s*=" % param, text), param)


class ParseFunctionTests(unittest.TestCase):
    def test_words_to_number(self):
        self.assertEqual(call("WORDSTONUM", "one lakh twenty-three thousand"), 123000.0)
        self.assertEqual(call("WORDSTONUM", "Rupees twelve and Paise fifty Only"), 12.5)
        self.assertEqual(call("WORDSTONUM", "one lakh crore"), 1e12)
        self.assertIsInstance(call("WORDSTONUM", "five"), float)

    def test_units_to_number(self):
        self.assertEqual(call("UNITSTONUM", "2cr"), 20000000.0)
        self.assertEqual(call("UNITSTONUM", "1.5k"), 1500.0)
        self.assertEqual(call("UNITSTONUM", "Rs. 2 crore"), 20000000.0)
        self.assertEqual(call("UNITSTONUM", 1500.0), 1500.0)                    # a numeric cell passes through

    def test_errors(self):
        for fn, arg in (("WORDSTONUM", ""), ("WORDSTONUM", "banana"), ("WORDSTONUM", None),
                        ("UNITSTONUM", "abc"), ("UNITSTONUM", None), ("UNITSTONUM", "1,5L")):
            with self.subTest(fn=fn, arg=arg), self.assertRaises(ValueError):
                call(fn, arg)


if __name__ == "__main__":
    unittest.main()
