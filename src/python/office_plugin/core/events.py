"""In-process event bus. Pure Python; the UNO listeners that feed it live in uno_events.py.

Known event names (payload = keyword args):
    selection_changed  controller, model          cursor/selection/page moved
    doc_modified       controller, model          document content changed
    view_created       controller, model          a new document window appeared
    view_closed        controller, model
    mouse_pressed      controller, modifiers, t
    mouse_released     controller, modifiers, t
    key_pressed        controller, key, modifiers, t
    frame_activated    controller, model
"""


class EventBus:
    def __init__(self, log=None):
        self._subs = {}
        self._log = log

    def subscribe(self, name, callback):
        """Returns a token for unsubscribe()."""
        self._subs.setdefault(name, []).append(callback)
        return (name, callback)

    def unsubscribe(self, token):
        name, callback = token
        try:
            self._subs.get(name, []).remove(callback)
        except ValueError:
            pass

    def publish(self, name, **payload):
        # copy: a handler may (un)subscribe while we iterate
        for callback in list(self._subs.get(name, ())):
            try:
                callback(**payload)
            except Exception:  # one bad subscriber must not break the others
                if self._log:
                    self._log.exception("event handler failed for %s", name)
