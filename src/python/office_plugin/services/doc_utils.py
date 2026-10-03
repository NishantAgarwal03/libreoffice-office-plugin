"""Document/controller helpers shared by features (needs a running LibreOffice)."""

_SERVICES = (
    ("writer", "com.sun.star.text.TextDocument"),
    ("calc", "com.sun.star.sheet.SpreadsheetDocument"),
    ("impress", "com.sun.star.presentation.PresentationDocument"),
    ("draw", "com.sun.star.drawing.DrawingDocument"),
)


class DocUtils:
    def __init__(self, ctx):
        self.ctx = ctx
        self._fallback_ids = []  # [(model, id)] when RuntimeUID is unavailable

    # ---- desktop ---------------------------------------------------------
    def desktop(self):
        smgr = self.ctx.getServiceManager()
        return smgr.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)

    def create(self, service):
        return self.ctx.getServiceManager().createInstanceWithContext(service, self.ctx)

    # ---- document identity ----------------------------------------------
    @staticmethod
    def kind(model):
        """'writer' | 'calc' | 'impress' | 'draw' | None."""
        if model is None:
            return None
        for kind, service in _SERVICES:
            try:
                if model.supportsService(service):
                    return kind
            except Exception:
                return None
        return None

    def doc_id(self, model):
        """Stable id of an open document (URL is empty for unsaved docs)."""
        try:
            uid = model.RuntimeUID
            if uid:
                return str(uid)
        except Exception:
            pass
        for known, ident in self._fallback_ids:
            if known == model:
                return ident
        ident = "doc%d" % (len(self._fallback_ids) + 1)
        self._fallback_ids.append((model, ident))
        return ident

    # ---- frames / controllers -------------------------------------------
    def active_frame(self):
        return self.desktop().getCurrentFrame()

    def controller_of(self, frame):
        return frame.getController() if frame is not None else None

    def active_controller(self):
        return self.controller_of(self.active_frame())

    def model_of(self, controller):
        return controller.getModel() if controller is not None else None

    def open_models(self):
        """All open document models that we support (skips Basic IDE, Start Center...)."""
        models = []
        enum = self.desktop().getComponents().createEnumeration()
        while enum.hasMoreElements():
            comp = enum.nextElement()
            if self.kind(comp):
                models.append(comp)
        return models

    def find_model(self, doc_id):
        for model in self.open_models():
            if self.doc_id(model) == doc_id:
                return model
        return None

    def bring_to_front(self, controller):
        try:
            frame = controller.getFrame()
            win = frame.getContainerWindow()
            win.setVisible(True)
            win.toFront()
            win.setFocus()
        except Exception:
            pass
