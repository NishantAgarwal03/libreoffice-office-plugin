"""Alt+Left / Alt+Right: back/forward through places you jumped from.

How the 'source' of a hyperlink jump is captured (see README, "Navigation design"):
every selection change refreshes a per-document last-known Position, so when a jump is
detected the place we jumped *from* is already known. A change counts as a jump when it
follows a mouse release (link click), or switches sheet/slide, or moves more than a
screen without keyboard navigation (Name Box, Navigator, Find...).
"""

import time

from office_plugin.core.feature import Feature

from .history import History, JumpClassifier

_CROSS_DOC_WINDOW_S = 1.5   # a new window opening this soon after a click = link to another document
_SELF_IGNORE_S = 0.4        # ignore selection events caused by our own restore


class NavigationFeature(Feature):
    def start(self):
        self.history = History(max_len=int(self.option("max_entries", 100)))
        self.classifier = JumpClassifier(
            record_far=bool(self.option("record_far_jumps", True)),
            record_container=bool(self.option("record_sheet_changes", True)))
        self.cache = {}            # doc_id -> last known Position
        self.cache_time = {}       # doc_id -> monotonic time the cache entry was set
        self.current = None        # most recent Position in any document
        self._ignore_until = {}    # doc_id -> monotonic deadline (events from our own restore)
        self._restoring = False
        self.subscribe("selection_changed", self._on_selection)
        self.subscribe("mouse_pressed", self._on_input)
        self.subscribe("mouse_released", self._on_release)
        self.subscribe("key_pressed", self._on_key)
        self.subscribe("view_created", self._on_view_created)
        self.subscribe("view_closed", self._on_view_closed)

    # ---- event handlers --------------------------------------------------
    def _on_input(self, controller=None, **_):
        self._sync(controller)

    def _on_release(self, controller=None, modifiers=0, t=None, **_):
        self._sync(controller)
        self.classifier.note_release(t if t is not None else time.monotonic())

    def _on_key(self, controller=None, t=None, **_):
        self._sync(controller)          # before noting the key: sync judges what happened *before* it
        self.classifier.note_key(t if t is not None else time.monotonic())

    def _sync(self, controller):
        """Catch moves that LibreOffice did not announce (Writer link / Navigator jumps).

        Runs on every user input and before Back/Forward, comparing the live position with
        the cached one; an unannounced move counts as a jump *from the cached place*.
        """
        if controller is None or self._restoring:
            return
        live = self.services.get("positions").capture(controller)
        if live is None:
            return
        cached = self.cache.get(live.doc_id)
        now = time.monotonic()
        if cached is None or now < self._ignore_until.get(live.doc_id, 0.0):
            self._remember(live, now)
            return
        jump, reason = self.classifier.is_silent_jump(cached, live, self.cache_time.get(live.doc_id, 0.0))
        if not live.same_place(cached):
            self.log.debug("navigation: silent move %s -> %s jump=%s (%s)", cached, live, jump, reason)
            if jump:
                self.history.record_jump(cached)
                self.refresh_state()
        self._remember(live, now)

    def _remember(self, pos, now):
        self.cache[pos.doc_id] = pos
        self.cache_time[pos.doc_id] = now
        self.current = pos

    def _on_selection(self, controller=None, model=None, page=None, **_):
        if controller is None:
            return
        positions = self.services.get("positions")
        pos = positions.capture(controller, page=page)
        if pos is None:
            return
        now = time.monotonic()
        prev = self.cache.get(pos.doc_id)
        self._remember(pos, now)
        if self._restoring or now < self._ignore_until.get(pos.doc_id, 0.0):
            return
        jump, reason = self.classifier.is_jump(prev, pos, now)
        self.log.debug("navigation: %s -> %s  jump=%s (%s)", prev, pos, jump, reason)
        if jump:
            self.history.record_jump(prev)
            self.refresh_state()

    def _on_view_created(self, controller=None, model=None, **_):
        """Link to another document: the new window opens right after a click."""
        released = self.classifier.last_release()
        if released is None or self.current is None:
            return
        if time.monotonic() - released > _CROSS_DOC_WINDOW_S:
            return
        docs = self.services.get("doc_utils")
        if model is None or docs.doc_id(model) == self.current.doc_id:
            return
        self.log.debug("navigation: cross-document jump from %s", self.current)
        self.history.record_jump(self.current)
        self.refresh_state()

    def _on_view_closed(self, controller=None, model=None, **_):
        if model is None:
            return
        doc_id = self.services.get("doc_utils").doc_id(model)
        self.history.drop_doc(doc_id)
        self.cache.pop(doc_id, None)
        self.cache_time.pop(doc_id, None)
        if self.current is not None and self.current.doc_id == doc_id:
            self.current = None
        self.refresh_state()

    # ---- commands ----------------------------------------------------------
    def execute(self, command, frame, args):
        docs = self.services.get("doc_utils")
        positions = self.services.get("positions")
        controller = docs.controller_of(frame) or docs.active_controller()
        self._sync(controller)
        current = positions.capture(controller) if controller is not None else None
        step = self.history.back if command == "Back" else self.history.forward
        while True:
            target = step(current)
            if target is None:
                return
            if self._go(target):
                break
            self.log.info("navigation: dropping unreachable entry %r", target)
        self.refresh_state()

    def _go(self, target):
        self._restoring = True
        try:
            ok = self.services.get("positions").restore(target)
        finally:
            self._restoring = False
        now = time.monotonic()
        self._ignore_until[target.doc_id] = now + _SELF_IGNORE_S
        if ok:      # we are the ones who moved: refresh the cache so it is not mistaken for a silent jump
            landed = self.services.get("positions").capture_for(target)
            if landed is not None:
                self._remember(landed, now)
        return ok

    def is_enabled(self, command, frame):
        return self.history.can_back() if command == "Back" else self.history.can_forward()
