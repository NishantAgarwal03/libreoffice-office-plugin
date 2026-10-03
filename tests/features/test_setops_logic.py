import unittest

import tests  # noqa: F401
from office_plugin.features.calc_setops import setops as S
from office_plugin.services.cellkeys import CleanOptions, EXACT

A = [("T-01", "Ceramic Tile", 10.0), ("T-02", "Paint", 5.0), ("T-03", "Putty", 2.0),
     ("T-02", "Paint", 5.0), ("", "", "")]
B = [("t-02 ", "paint", 5.0), ("T-04", "Cement", 1.0), ("101", "x", 1.0), ("T-01", "ceramic tile", 9.0)]


def src(result):
    return [(s, i) for s, i in result]


class OperationTests(unittest.TestCase):
    def run_op(self, op, keys=(0,), opts=CleanOptions(), a=A, b=B):
        return S.compute(op, a, b, list(keys), opts)

    def test_intersection(self):
        result, counts = self.run_op(S.INTERSECTION)
        self.assertEqual(src(result), [("A", 0), ("A", 1)])       # T-01, T-02 (first occurrence)
        self.assertEqual(counts["common"], 2)

    def test_a_minus_b_and_b_minus_a(self):
        self.assertEqual(src(self.run_op(S.A_MINUS_B)[0]), [("A", 2)])           # T-03
        self.assertEqual(src(self.run_op(S.B_MINUS_A)[0]), [("B", 1), ("B", 2)])  # T-04, 101

    def test_union_order_and_distinct(self):
        result, counts = self.run_op(S.UNION)
        self.assertEqual(src(result), [("A", 0), ("A", 1), ("A", 2), ("B", 1), ("B", 2)])
        self.assertEqual(counts["result"], 5)

    def test_symmetric_difference(self):
        self.assertEqual(src(self.run_op(S.SYM_DIFF)[0]), [("A", 2), ("B", 1), ("B", 2)])

    def test_multi_column_key(self):
        # (code, name): T-01/"Ceramic Tile" matches B's T-01/"ceramic tile" only when case is ignored
        self.assertEqual(src(self.run_op(S.INTERSECTION, keys=(0, 1))[0]), [("A", 0), ("A", 1)])
        exact = CleanOptions(True, False, True, True)
        self.assertEqual(src(self.run_op(S.INTERSECTION, keys=(0, 1), opts=exact)[0]), [])

    def test_clean_options_change_matching(self):
        raw_a, raw_b = [("T-02",)], [("t-02 ",)]
        self.assertEqual(self.run_op(S.INTERSECTION, a=raw_a, b=raw_b)[1]["common"], 1)
        self.assertEqual(self.run_op(S.INTERSECTION, a=raw_a, b=raw_b, opts=EXACT)[1]["common"], 0)

    def test_numeric_text_matches_number(self):
        a, b = [(101.0,), (7.0,)], [("101",), ("7.5",)]
        self.assertEqual(src(self.run_op(S.INTERSECTION, a=a, b=b)[0]), [("A", 0)])
        no_num = CleanOptions(True, True, False, True)
        self.assertEqual(src(self.run_op(S.INTERSECTION, a=a, b=b, opts=no_num)[0]), [])

    def test_blank_key_rows_never_participate(self):
        a, b = [("",), ("x",)], [("",), ("y",)]
        result, counts = self.run_op(S.UNION, a=a, b=b)
        self.assertEqual(src(result), [("A", 1), ("B", 1)])
        self.assertEqual(counts["skipped_blank"], 2)

    def test_empty_lists_and_errors(self):
        self.assertEqual(self.run_op(S.UNION, a=[], b=[])[0], [])
        self.assertEqual(self.run_op(S.INTERSECTION, a=A, b=[])[0], [])
        self.assertEqual(len(self.run_op(S.A_MINUS_B, a=A, b=[])[0]), 3)
        with self.assertRaises(ValueError):
            S.compute("Nope", A, B, [0])
        with self.assertRaises(ValueError):
            S.compute(S.UNION, A, B, [])

    def test_operations_are_consistent(self):
        union = {(s, i) for s, i in self.run_op(S.UNION)[0]}
        parts = {(s, i) for op in (S.A_MINUS_B, S.B_MINUS_A) for s, i in self.run_op(op)[0]}
        both = {(s, i) for s, i in self.run_op(S.INTERSECTION)[0]}
        self.assertEqual(union, parts | both)


if __name__ == "__main__":
    unittest.main()
