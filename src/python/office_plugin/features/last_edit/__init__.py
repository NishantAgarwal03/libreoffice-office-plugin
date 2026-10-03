"""Go to last edit: remembers the last 3 distinct edit locations per document."""

from office_plugin.core.feature import Feature

from .ring import EditRing


class LastEditFeature(Feature):
    def start(self):
        self.rings = {}  # doc_id -> EditRing
        self.subscribe("doc_modified", self._on_modified)
        self.subscribe("view_closed", self._on_closed)

    def _ring(self, doc_id):
        return self.rings.setdefault(doc_id, EditRing(size=int(self.option("history_size", 3))))

    def _on_modified(self, controller=None, model=None, **_):
        if controller is None:
            return
        pos = self.services.get("positions").capture(controller)
        if pos is None or pos.kind not in ("writer", "calc"):
            return
        self._ring(pos.doc_id).record(pos)

    def _on_closed(self, controller=None, model=None, **_):
        if model is not None:
            self.rings.pop(self.services.get("doc_utils").doc_id(model), None)

    def execute(self, command, frame, args):
        docs = self.services.get("doc_utils")
        positions = self.services.get("positions")
        controller = docs.controller_of(frame) or docs.active_controller()
        if controller is None:
            return
        current = positions.capture(controller)
        doc_id = docs.doc_id(controller.getModel())
        ring = self.rings.get(doc_id)
        if not ring:
            return
        for _ in range(len(ring)):
            target = ring.next_target(current)
            if target is None or positions.restore(target):
                return
        self.log.info("last_edit: no reachable edit location")

    def is_enabled(self, command, frame):
        return True
