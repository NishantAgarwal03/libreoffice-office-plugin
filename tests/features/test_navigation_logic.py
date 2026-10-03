import math
import unittest

import tests  # noqa: F401
from office_plugin.core.position import Position
from office_plugin.features.navigation.history import History, JumpClassifier


def calc(row, col=0, sheet=0, doc="d1"):
    return Position("calc", doc, container=sheet, coords=(row, col))


def writer(y, doc="d1"):
    return Position("writer", doc, container=0, coords=(y,))


class PositionTests(unittest.TestCase):
    def test_same_place_and_distance(self):
        self.assertTrue(calc(1, 1).same_place(calc(1, 1)))
        self.assertFalse(calc(1, 1).same_place(calc(1, 2)))
        self.assertEqual(calc(0).distance(calc(50)), 1.0)           # 50 rows = one screen
        self.assertEqual(calc(0).distance(calc(0, 10)), 0.5)        # 10 of 20 columns
        self.assertTrue(math.isinf(calc(0).distance(calc(0, sheet=1))))
        self.assertTrue(math.isinf(calc(0).distance(calc(0, doc="d2"))))
        self.assertEqual(writer(0).distance(writer(40000)), 1.0)


class HistoryTests(unittest.TestCase):
    def test_back_then_forward_roundtrip(self):
        h = History()
        a, b, c = calc(1), calc(100), calc(200)
        h.record_jump(a)                      # jumped from A to B
        self.assertIs(h.back(b), a)           # Alt+Left from B -> A
        self.assertIs(h.forward(a), b)        # Alt+Right from A -> B
        h.record_jump(b)                      # B -> C
        self.assertIs(h.back(c), b)
        self.assertIs(h.back(b), a)           # chain continues (a was pushed earlier? no) ...
        self.assertIsNone(h.back(a))

    def test_new_jump_clears_forward(self):
        h = History()
        a, b, c = calc(1), calc(100), calc(200)
        h.record_jump(a)
        h.back(b)                             # now at A, forward = [B]
        self.assertTrue(h.can_forward())
        h.record_jump(a)                      # user jumps again from A
        self.assertFalse(h.can_forward())

    def test_dedupe_cap_and_skip_current(self):
        h = History(max_len=3)
        for r in (1, 2, 3, 4):
            h.record_jump(calc(r * 100))
        self.assertEqual(len(h), 3)           # cap
        h2 = History()
        h2.record_jump(calc(5))
        h2.record_jump(calc(5))               # duplicate neighbour ignored
        self.assertEqual(len(h2), 1)
        h3 = History()
        h3.record_jump(calc(5))
        self.assertIsNone(h3.back(calc(5)))   # entry equal to current is skipped, nothing left

    def test_empty_history(self):
        h = History()
        self.assertIsNone(h.back(calc(1)))
        self.assertIsNone(h.forward(calc(1)))
        self.assertFalse(h.can_back())

    def test_drop_doc(self):
        h = History()
        h.record_jump(calc(1, doc="x"))
        h.record_jump(calc(2, doc="y"))
        h.drop_doc("x")
        self.assertEqual(len(h), 1)

    def test_cross_document_entries(self):
        h = History()
        h.record_jump(calc(1, doc="book1"))
        target = h.back(calc(0, doc="book2"))
        self.assertEqual(target.doc_id, "book1")
        self.assertEqual(h.forward(target).doc_id, "book2")


class JumpClassifierTests(unittest.TestCase):
    def setUp(self):
        self.c = JumpClassifier()

    def test_click_followed_by_move_is_a_jump(self):
        self.c.note_release(10.0)
        self.assertEqual(self.c.is_jump(calc(1), calc(30), 10.2), (True, "click"))

    def test_click_far_in_time_is_not(self):
        self.c.note_release(10.0)
        self.assertFalse(self.c.is_jump(calc(1), calc(10), 12.0)[0])

    def test_tiny_move_after_click_is_ignored(self):
        self.c.note_release(10.0)
        self.assertEqual(self.c.is_jump(calc(5), calc(5, 0), 10.1)[0], False)   # same place
        self.assertEqual(self.c.is_jump(writer(100), writer(500), 10.1)[0], False)  # ~4 mm

    def test_keyboard_navigation_never_jumps(self):
        self.c.note_key(5.0)
        self.assertEqual(self.c.is_jump(calc(1), calc(5000), 5.1), (False, "keyboard"))
        self.assertEqual(self.c.is_jump(calc(1), calc(1, sheet=2), 5.1), (False, "keyboard"))

    def test_sheet_change_and_far_move_without_input(self):
        self.assertEqual(self.c.is_jump(calc(1), calc(1, sheet=1), 1.0), (True, "container"))
        self.assertEqual(self.c.is_jump(calc(1), calc(500), 1.0), (True, "far"))
        self.assertEqual(self.c.is_jump(calc(1), calc(10), 1.0), (False, "near"))

    def test_options_can_disable_heuristics(self):
        c = JumpClassifier(record_far=False, record_container=False)
        self.assertFalse(c.is_jump(calc(1), calc(500), 1.0)[0])
        self.assertFalse(c.is_jump(calc(1), calc(1, sheet=1), 1.0)[0])

    def test_no_previous_position(self):
        self.assertEqual(self.c.is_jump(None, calc(1), 1.0), (False, "no-move"))


class SilentJumpTests(unittest.TestCase):
    """Writer moves the cursor on a link/Navigator jump without telling listeners."""

    def setUp(self):
        self.c = JumpClassifier()

    def test_click_since_cache_then_moved(self):
        self.c.note_release(5.0)
        self.assertEqual(self.c.is_silent_jump(writer(0), writer(200000), cache_time=1.0),
                         (True, "silent-click"))

    def test_click_before_cache_does_not_count(self):
        self.c.note_release(0.5)       # the click happened before the cache was refreshed
        self.assertFalse(self.c.is_silent_jump(writer(0), writer(500), cache_time=1.0)[0])

    def test_tiny_move_after_click_ignored(self):
        self.c.note_release(5.0)
        self.assertFalse(self.c.is_silent_jump(writer(0), writer(300), cache_time=1.0)[0])

    def test_far_silent_move_without_input_is_a_jump(self):      # Navigator / Go To
        self.assertEqual(self.c.is_silent_jump(writer(0), writer(200000), cache_time=1.0),
                         (True, "silent-far"))

    def test_keyboard_since_cache_suppresses(self):
        self.c.note_key(3.0)
        self.assertEqual(self.c.is_silent_jump(writer(0), writer(200000), cache_time=1.0),
                         (False, "keyboard"))

    def test_same_place_and_near_moves(self):
        self.assertEqual(self.c.is_silent_jump(writer(5), writer(5), 1.0), (False, "no-move"))
        self.assertEqual(self.c.is_silent_jump(writer(0), writer(2000), 1.0), (False, "near"))

    def test_sheet_change_silent(self):
        self.assertEqual(self.c.is_silent_jump(calc(1), calc(1, sheet=2), 1.0), (True, "silent-container"))


if __name__ == "__main__":
    unittest.main()
