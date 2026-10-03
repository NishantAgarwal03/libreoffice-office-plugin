"""TEMPLATE - copy this folder to features/<your_feature_id>/ (see README.md).

Folder names starting with '_' are ignored by the runtime and by build_oxt.py.
"""

from office_plugin.core.feature import Feature


class MyFeature(Feature):
    def start(self):
        # Optional: react to events. self.subscribe() is undone automatically in stop().
        # self.subscribe("selection_changed", self._on_selection)
        pass

    def execute(self, command, frame, args):
        docs = self.services.get("doc_utils")
        controller = docs.controller_of(frame) or docs.active_controller()
        self.log.info("%s.%s called on %s", self.id, command, controller)

    def is_enabled(self, command, frame):
        return True
