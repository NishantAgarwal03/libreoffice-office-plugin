"""Select Visible Cells Only (Excel's Alt+;)."""

from office_plugin.core.feature import Feature


class CalcVisibleFeature(Feature):
    def execute(self, command, frame, args):
        docs = self.services.get("doc_utils")
        controller = docs.controller_of(frame) or docs.active_controller()
        if controller is None:
            return
        selection = controller.getSelection()
        if not hasattr(selection, "queryVisibleCells"):
            self.services.get("ui").message(
                "Select Visible Cells", "Select a range of cells first.", frame, "warning")
            return
        visible = selection.queryVisibleCells()
        if visible.getCount() == 0:
            self.services.get("ui").message(
                "Select Visible Cells", "The selection contains no visible cells.", frame, "warning")
            return
        controller.select(visible)
