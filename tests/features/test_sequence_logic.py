import unittest
from decimal import Decimal as D

import tests  # noqa: F401
from office_plugin.features.calc_sequence import sequence as S


def series(start="1", step=None, end=None, count=None, default=None):
    p = lambda t, n: S.parse_number(t, n)
    return S.build_series(p(start, "Start"), p(step, "Step"), p(end, "End"), p(count, "Count"), default)


class ParseTests(unittest.TestCase):
    def test_parse(self):
        self.assertEqual(S.parse_number(" 2,5 ", "x"), D("2.5"))       # decimal comma
        self.assertIsNone(S.parse_number("", "x"))
        self.assertIsNone(S.parse_number(None, "x"))
        for bad in ("abc", "1e999999", "NaN", "Infinity", "1 2"):
            with self.assertRaises(ValueError, msg=bad):
                S.parse_number(bad, "Start")


class SeriesTests(unittest.TestCase):
    def test_start_step_count(self):
        self.assertEqual(series("1", "2", count="4"), [1, 3, 5, 7])
        self.assertEqual(series("10", "-3", count="4"), [10, 7, 4, 1])

    def test_start_step_end_is_inclusive_only_when_the_step_lands(self):
        self.assertEqual(series("1", "2", end="7"), [1, 3, 5, 7])
        self.assertEqual(series("1", "2", end="8"), [1, 3, 5, 7])
        self.assertEqual(series("5", "-1", end="2"), [5, 4, 3, 2])
        self.assertEqual(series("1", "1", end="1"), [1])

    def test_start_end_count_computes_the_step(self):
        self.assertEqual(series("0", end="1", count="5"), [0, 0.25, 0.5, 0.75, 1])
        self.assertEqual(series("1", end="10", count="10"), list(range(1, 11)))
        self.assertEqual(series("5", end="9", count="1"), [5])

    def test_default_count_from_the_selection(self):
        self.assertEqual(series("1", "1", default=3), [1, 2, 3])
        self.assertEqual(series("100", default=2), [100, 101])         # step defaults to 1

    def test_decimal_steps_are_exact(self):
        self.assertEqual(series("0", "0.1", count="4"), [0, 0.1, 0.2, 0.3])
        self.assertEqual(series("0.5", "0.25", end="1.25"), [0.5, 0.75, 1, 1.25])

    def test_all_three_given_stops_at_whichever_comes_first(self):
        self.assertEqual(series("1", "1", end="3", count="10"), [1, 2, 3])
        self.assertEqual(series("1", "1", end="30", count="3"), [1, 2, 3])

    def test_end_without_step_picks_the_direction(self):
        self.assertEqual(series("5", end="3"), [5, 4, 3])
        self.assertEqual(series("1", end="3"), [1, 2, 3])

    def test_errors(self):
        cases = [dict(start="", step="1", count="3"),                  # no start
                 dict(start="1", step="0", count="3"),
                 dict(start="1", step="-1", end="5"),                  # wrong way
                 dict(start="1", step="1", count="0"),
                 dict(start="1", step="1", count="2.5"),
                 dict(start="1", step="1"),                            # nothing to bound it
                 dict(start="3", end="3", count="5"),                  # step would be 0
                 dict(start="1", step="1", count="2000000")]
        for case in cases:
            with self.subTest(case), self.assertRaises(ValueError):
                series(**case)

    def test_integral_results_are_ints(self):
        self.assertTrue(all(isinstance(v, int) for v in series("1", "1", count="3")))
        self.assertTrue(all(isinstance(v, float) for v in series("0.5", "1", count="3")))

    def test_preview(self):
        self.assertEqual(S.preview([1, 2, 3]), "1, 2, 3")
        self.assertEqual(S.preview(list(range(1, 51))), "1, 2, 3, 4 ... 50  (50 values)")


if __name__ == "__main__":
    unittest.main()
