"""Pure value object describing "a place in a document".

Used by navigation (back/forward) and last_edit. It carries only plain data plus an
opaque `handle` that the UNO-side service (services/positions.py) knows how to
restore; nothing here imports `uno`.
"""

import math

# One "screen" of movement per kind, used to normalise distances so a single
# threshold works across apps (Calc rows/cols, Writer document Y in 1/100 mm).
_SCREEN = {
    "calc": (50.0, 20.0),   # rows, columns
    "writer": (40000.0,),   # 1/100 mm  (~40 cm)
}


class Position:
    __slots__ = ("kind", "doc_id", "container", "coords", "handle", "label")

    def __init__(self, kind, doc_id, container=None, coords=(), handle=None, label=""):
        self.kind = kind            # "writer" | "calc" | "impress" | "draw"
        self.doc_id = doc_id        # stable per-open-document id
        self.container = container  # sheet index / page index / None
        self.coords = tuple(coords)  # (row, col) for calc, (y,) for writer, () otherwise
        self.handle = handle        # opaque restore data for services.positions
        self.label = label

    def same_doc(self, other):
        return other is not None and self.doc_id == other.doc_id

    def same_place(self, other):
        return (
            other is not None
            and self.kind == other.kind
            and self.doc_id == other.doc_id
            and self.container == other.container
            and self.coords == other.coords
        )

    def distance(self, other):
        """Normalised distance in 'screens'; inf when in another doc/sheet/page."""
        if other is None or self.kind != other.kind or self.doc_id != other.doc_id:
            return math.inf
        if self.container != other.container:
            return math.inf
        scale = _SCREEN.get(self.kind)
        if not scale or not self.coords or not other.coords:
            return 0.0
        return max(abs(a - b) / s for a, b, s in zip(self.coords, other.coords, scale))

    def __repr__(self):
        return "Position(%s, doc=%s, container=%r, coords=%r)" % (
            self.kind, self.doc_id, self.container, self.coords)
