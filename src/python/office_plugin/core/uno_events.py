"""Bridges LibreOffice listeners onto the pure EventBus (the only module that knows UNO
listener interfaces). Requires a running LibreOffice."""

import time

import uno  # noqa: F401  (activates the `com.sun.star` import hook)
import unohelper
from com.sun.star.awt import XKeyHandler, XMouseClickHandler
from com.sun.star.beans import XPropertyChangeListener
from com.sun.star.document import XDocumentEventListener
from com.sun.star.util import XModifyListener
from com.sun.star.view import XSelectionChangeListener

_ATTACH_EVENTS = {"OnViewCreated", "OnFocus", "OnLoadFinished", "OnNew", "OnLoad"}


class _Selection(unohelper.Base, XSelectionChangeListener):
    def __init__(self, hub, controller):
        self.hub, self.controller = hub, controller

    def selectionChanged(self, ev):
        self.hub._emit("selection_changed", controller=self.controller)

    def disposing(self, ev):
        self.hub._forget(self.controller)


class _PageProp(unohelper.Base, XPropertyChangeListener):
    """Impress/Draw: switching slides is a property change, not a selection change."""

    def __init__(self, hub, controller):
        self.hub, self.controller = hub, controller

    def propertyChange(self, ev):
        # The event fires BEFORE controller.getCurrentPage() reflects the change, so hand
        # the new page to the subscribers explicitly.
        self.hub._emit("selection_changed", controller=self.controller, page=ev.NewValue)

    def disposing(self, ev):
        pass


class _Mouse(unohelper.Base, XMouseClickHandler):
    def __init__(self, hub, controller):
        self.hub, self.controller = hub, controller

    def mousePressed(self, ev):
        self.hub._emit("mouse_pressed", controller=self.controller,
                       modifiers=ev.Modifiers, t=time.monotonic())
        return False  # never consume

    def mouseReleased(self, ev):
        self.hub._emit("mouse_released", controller=self.controller,
                       modifiers=ev.Modifiers, t=time.monotonic())
        return False

    def disposing(self, ev):
        pass


class _Key(unohelper.Base, XKeyHandler):
    def __init__(self, hub, controller):
        self.hub, self.controller = hub, controller

    def keyPressed(self, ev):
        self.hub._emit("key_pressed", controller=self.controller, key=ev.KeyCode,
                       modifiers=ev.Modifiers, t=time.monotonic())
        return False

    def keyReleased(self, ev):
        return False

    def disposing(self, ev):
        pass


class _Modify(unohelper.Base, XModifyListener):
    def __init__(self, hub, model):
        self.hub, self.model = hub, model

    def modified(self, ev):
        controller = None
        try:
            controller = self.model.getCurrentController()
        except Exception:
            pass
        self.hub._emit("doc_modified", controller=controller, model=self.model)

    def disposing(self, ev):
        pass


class _DocEvents(unohelper.Base, XDocumentEventListener):
    def __init__(self, hub):
        self.hub = hub

    def documentEventOccured(self, ev):
        try:
            self.hub._on_document_event(ev)
        except Exception:
            self.hub.log.exception("document event failed")

    def disposing(self, ev):
        pass


class _Attached:
    def __init__(self, controller, model):
        self.controller, self.model = controller, model
        self.listeners = {}


class UnoEventHub:
    def __init__(self, ctx, events, doc_utils, log):
        self.ctx, self.events, self.docs, self.log = ctx, events, doc_utils, log
        self._attached = []
        self._modify_models = []   # [(model, listener)]
        self._doc_listener = None
        self._broadcaster = None

    # ---- lifecycle -------------------------------------------------------
    def start(self):
        self._doc_listener = _DocEvents(self)
        self._broadcaster = self.ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.frame.GlobalEventBroadcaster", self.ctx)
        self._broadcaster.addDocumentEventListener(self._doc_listener)
        for model in self.docs.open_models():       # documents opened before we started
            self._attach_model(model)

    def stop(self):
        if self._broadcaster and self._doc_listener:
            try:
                self._broadcaster.removeDocumentEventListener(self._doc_listener)
            except Exception:
                pass
        for rec in list(self._attached):
            self._detach(rec)
        for model, listener in self._modify_models:
            try:
                model.removeModifyListener(listener)
            except Exception:
                pass
        self._modify_models = []

    # ---- document events -------------------------------------------------
    def _on_document_event(self, ev):
        name = ev.EventName
        model = ev.Source
        if name in _ATTACH_EVENTS and self.docs.kind(model):
            controller = self._view_controller(ev) or model.getCurrentController()
            self._attach_model(model, controller)
            if name == "OnViewCreated":
                self._emit("view_created", controller=controller, model=model)
            elif name == "OnFocus":
                self._emit("frame_activated", controller=controller, model=model)
        elif name in ("OnViewClosed", "OnUnload") and self.docs.kind(model):
            controller = self._view_controller(ev)
            self._emit("view_closed", controller=controller, model=model)
            for rec in list(self._attached):
                if rec.model == model:
                    self._detach(rec)

    @staticmethod
    def _view_controller(ev):
        """ev.ViewController, or None (it is already disposed when a window is closing)."""
        try:
            return ev.ViewController
        except Exception:
            return None

    # ---- attach / detach -------------------------------------------------
    def _attach_model(self, model, controller=None):
        controller = controller or model.getCurrentController()
        if controller is None:
            return
        if any(rec.controller == controller for rec in self._attached):
            return
        rec = _Attached(controller, model)
        self._try(rec, "selection", _Selection(self, controller),
                  lambda l: controller.addSelectionChangeListener(l))
        self._try(rec, "mouse", _Mouse(self, controller),
                  lambda l: controller.addMouseClickHandler(l))
        self._try(rec, "key", _Key(self, controller),
                  lambda l: controller.addKeyHandler(l))
        if self.docs.kind(model) in ("impress", "draw"):
            self._try(rec, "page", _PageProp(self, controller),
                      lambda l: controller.addPropertyChangeListener("CurrentPage", l))
        if not any(m == model for m, _ in self._modify_models):
            listener = _Modify(self, model)
            try:
                model.addModifyListener(listener)
                self._modify_models.append((model, listener))
            except Exception:
                self.log.debug("addModifyListener failed", exc_info=True)
        self._attached.append(rec)
        self._emit("selection_changed", controller=controller)   # warm position caches
        self.log.debug("attached to %s (%s): %s", self.docs.doc_id(model),
                       self.docs.kind(model), sorted(rec.listeners))

    def _try(self, rec, name, listener, adder):
        try:
            adder(listener)
            rec.listeners[name] = listener
        except Exception:
            self.log.debug("could not add %s listener", name, exc_info=True)

    def _detach(self, rec):
        c = rec.controller
        removers = {
            "selection": lambda l: c.removeSelectionChangeListener(l),
            "mouse": lambda l: c.removeMouseClickHandler(l),
            "key": lambda l: c.removeKeyHandler(l),
            "page": lambda l: c.removePropertyChangeListener("CurrentPage", l),
        }
        for name, listener in rec.listeners.items():
            try:
                removers[name](listener)
            except Exception:
                pass
        if rec in self._attached:
            self._attached.remove(rec)

    def _forget(self, controller):
        for rec in list(self._attached):
            if rec.controller == controller:
                self._attached.remove(rec)

    # ---- emit ------------------------------------------------------------
    def _emit(self, name, controller=None, **payload):
        model = payload.pop("model", None)
        if model is None and controller is not None:
            try:
                model = controller.getModel()
            except Exception:
                model = None
        self.events.publish(name, controller=controller, model=model, **payload)
