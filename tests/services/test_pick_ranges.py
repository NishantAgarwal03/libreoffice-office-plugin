import unittest

import tests  # noqa: F401
from office_plugin.services.pick_ranges import pick_lists


class FakePicker:
    """Plays the user: answers each non-modal pick with the next scripted reply."""

    def __init__(self, replies, available=True):
        self.replies, self.available, self.calls = list(replies), available, []

    def pick_range(self, controller, title, initial, callback):
        if not self.available:
            return False
        self.calls.append((title, initial))
        callback(self.replies.pop(0))        # a real picker calls back later; same contract
        return True


class PickListsTests(unittest.TestCase):
    def run_wizard(self, picker, want_b=True):
        result = []
        pick_lists(picker, None, "$S.$A$1:$B$2", want_b, lambda a, b: result.append((a, b)), "T")
        return result

    def test_picks_a_then_b(self):
        picker = FakePicker(["$S.$A$1:$C$9", "$S.$E$1:$G$9"])
        self.assertEqual(self.run_wizard(picker), [("$S.$A$1:$C$9", "$S.$E$1:$G$9")])
        self.assertEqual(picker.calls[0], ("T - select list A", "$S.$A$1:$B$2"))   # A starts at the selection
        self.assertEqual(picker.calls[1][0], "T - select list B")

    def test_single_list_mode_skips_b(self):
        picker = FakePicker(["$S.$A$1:$C$9"])
        self.assertEqual(self.run_wizard(picker, want_b=False), [("$S.$A$1:$C$9", "")])
        self.assertEqual(len(picker.calls), 1)

    def test_abort_at_either_step_never_calls_back(self):
        self.assertEqual(self.run_wizard(FakePicker([None])), [])
        self.assertEqual(self.run_wizard(FakePicker(["$S.$A$1"] + [None])), [])

    def test_unavailable_picker_falls_back_to_typed_addresses(self):
        self.assertEqual(self.run_wizard(FakePicker([], available=False)), [("$S.$A$1:$B$2", "")])

    def test_b_picker_unavailable_after_a_keeps_a(self):
        class HalfPicker(FakePicker):
            def pick_range(self, controller, title, initial, callback):
                if "list B" in title:
                    return False
                return super().pick_range(controller, title, initial, callback)

        self.assertEqual(self.run_wizard(HalfPicker(["$S.$A$1:$C$9"])), [("$S.$A$1:$C$9", "")])


if __name__ == "__main__":
    unittest.main()
