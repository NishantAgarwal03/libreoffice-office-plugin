"""Shared "pick list A (and list B) with the mouse" wizard used by the Calc compare tools.

Purpose: chain LibreOffice's non-modal range-selection popups (A, then optionally B) and hand the
         chosen addresses to a callback; fall back to typed addresses when picking is unavailable.
Inputs:  ui service, controller, the initial text for A, whether B is needed, callback.
Outputs: callback(a_text, b_text) -- never called when the user aborts a pick.
Must not: import uno (the ui service owns all UNO access).
"""


def pick_lists(ui, controller, a_initial, want_b, done, title="Compare"):
    """Start the wizard. `done(a_text, b_text)`; b_text is '' when B is not wanted or when the
    picker is unavailable (the caller's dialog then shows editable text boxes)."""
    def got_a(a_text):
        if a_text is None:
            return
        if not want_b:
            done(a_text, "")
            return

        def got_b(b_text):
            if b_text is not None:
                done(a_text, b_text)

        if not ui.pick_range(controller, "%s - select list B" % title, "", got_b):
            done(a_text, "")

    first_title = "%s - select list A" % title if want_b else "%s - select the list" % title
    if not ui.pick_range(controller, first_title, a_initial, got_a):
        done(a_initial, "")
