"""Back/forward history (two-stack, browser style) and the jump classifier.

Pure Python: operates on core.position.Position objects, no uno import.
"""

import math


class History:
    def __init__(self, max_len=100):
        self.max_len = max_len
        self._back = []
        self._fwd = []

    # ---- recording -------------------------------------------------------
    def record_jump(self, source):
        """The user jumped away from `source`; remember it and drop the forward stack."""
        if self._back and self._back[-1].same_place(source):
            self._fwd.clear()
            return
        self._back.append(source)
        if len(self._back) > self.max_len:
            del self._back[0]
        self._fwd.clear()

    # ---- navigation ------------------------------------------------------
    def can_back(self, current=None):
        return any(not p.same_place(current) for p in self._back) if current else bool(self._back)

    def can_forward(self, current=None):
        return any(not p.same_place(current) for p in self._fwd) if current else bool(self._fwd)

    def back(self, current):
        """Pop the previous place (skipping entries equal to `current`); push `current` forward."""
        return self._step(self._back, self._fwd, current)

    def forward(self, current):
        return self._step(self._fwd, self._back, current)

    def _step(self, src, dst, current):
        while src and current is not None and src[-1].same_place(current):
            src.pop()
        if not src:
            return None
        target = src.pop()
        if current is not None and not (dst and dst[-1].same_place(current)):
            dst.append(current)
            if len(dst) > self.max_len:
                del dst[0]
        return target

    def drop_doc(self, doc_id):
        """Forget entries of a closed document that cannot be reopened."""
        self._back = [p for p in self._back if p.doc_id != doc_id]
        self._fwd = [p for p in self._fwd if p.doc_id != doc_id]

    def __len__(self):
        return len(self._back) + len(self._fwd)


class JumpClassifier:
    """Decides whether a selection change was a 'jump' (hyperlink, Name Box, Find...)
    rather than ordinary cursor movement.

    Signals (fed by the navigation feature):
      * mouse release shortly before the change  -> a click activated a link
      * keyboard activity in the document window  -> ordinary keyboard navigation
      * a change of sheet/page, or a move of more than `far_distance` screens
    """

    def __init__(self, post_release_s=0.6, key_suppress_s=0.3,
                 click_min_distance=0.05, far_distance=1.0,
                 record_far=True, record_container=True):
        self.post_release_s = post_release_s
        self.key_suppress_s = key_suppress_s
        self.click_min_distance = click_min_distance
        self.far_distance = far_distance
        self.record_far = record_far
        self.record_container = record_container
        self._last_release = None
        self._last_key = None

    def note_release(self, t):
        self._last_release = t

    def note_key(self, t):
        self._last_key = t

    def last_release(self):
        return self._last_release

    def is_silent_jump(self, cached, live, cache_time):
        """Writer moves the view cursor on a link/Navigator jump WITHOUT notifying selection
        listeners, so we notice it later by comparing the live position with the cache.

        `cache_time` is when `cached` was recorded. Returns (bool, reason):
          * a mouse click happened since then and the cursor moved -> link click
          * no keyboard activity since then and the cursor moved far / to another place
            container -> Navigator, Go To, programmatic jump
        """
        if cached is None or live is None or cached.same_place(live):
            return False, "no-move"
        d = cached.distance(live)
        if (self._last_release is not None and self._last_release > cache_time
                and d >= self.click_min_distance):
            return True, "silent-click"
        keyed = self._last_key is not None and self._last_key > cache_time
        if keyed:
            return False, "keyboard"
        if math.isinf(d):
            return (True, "silent-container") if self.record_container else (False, "container-off")
        if self.record_far and d >= self.far_distance:
            return True, "silent-far"
        return False, "near"

    def is_jump(self, prev, new, now):
        """-> (bool, reason). `prev`/`new` are Positions in the same document."""
        if prev is None or new is None or prev.same_place(new):
            return False, "no-move"
        d = prev.distance(new)
        released = self._last_release is not None and (now - self._last_release) <= self.post_release_s
        if released and d >= self.click_min_distance:
            return True, "click"
        recent_key = self._last_key is not None and (now - self._last_key) <= self.key_suppress_s
        if recent_key:
            return False, "keyboard"
        if math.isinf(d):
            return (True, "container") if self.record_container else (False, "container-off")
        if self.record_far and d >= self.far_distance:
            return True, "far"
        return False, "near"
