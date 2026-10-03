"""Ring of the last N distinct edit locations (pure Python)."""


class EditRing:
    def __init__(self, size=3, merge_distance=0.02):
        self.size = size
        self.merge_distance = merge_distance  # edits this close are the same "place"
        self._items = []   # oldest ... newest
        self._cycle = None  # index into newest-first order while cycling

    def record(self, pos):
        """Called after every edit. Edits near the previous one just move it."""
        self._cycle = None
        if self._items:
            last = self._items[-1]
            if last.same_doc(pos) and last.distance(pos) <= self.merge_distance:
                self._items[-1] = pos
                return
        self._items.append(pos)
        if len(self._items) > self.size:
            del self._items[0]

    def reset_cycle(self):
        self._cycle = None

    def next_target(self, current=None):
        """Next location to visit, newest first, wrapping. Skips the spot we are already at
        on the first press (Word behaviour)."""
        order = list(reversed(self._items))
        if not order:
            return None
        if self._cycle is None:
            idx = 0
            if current is not None and order[0].same_doc(current) and \
                    order[0].distance(current) <= self.merge_distance and len(order) > 1:
                idx = 1
        else:
            idx = (self._cycle + 1) % len(order)
        self._cycle = idx
        return order[idx]

    def __len__(self):
        return len(self._items)

    def drop_doc(self, doc_id):
        self._items = [p for p in self._items if p.doc_id != doc_id]
        self._cycle = None
