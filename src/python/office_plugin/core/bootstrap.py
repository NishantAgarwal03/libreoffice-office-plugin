"""Builds the whole plugin once per LibreOffice process (singleton)."""

import os

from . import log as log_mod
from .dispatch import Router
from .events import EventBus
from .registry import Registry
from .services import ServiceContainer
from .settings import Settings

_plugin = None


class Plugin:
    def __init__(self, ctx, profile_dir=None):
        self.ctx = ctx
        base = os.path.join(profile_dir, "office_plugin") if profile_dir else None
        self.settings = Settings(os.path.join(base, "settings.json") if base else None)
        self.log = log_mod.setup(base, debug=self.settings.debug)
        self.events = EventBus(self.log)
        self.services = ServiceContainer()
        self.registry = Registry(ctx, self.services, self.events, self.settings, self.log)
        self.router = Router(self.registry, self.log)
        self.hub = None
        self.status_notifier = lambda feature_id, command=None: None  # set by officeplugin.py
        self._register_services()

    def _register_services(self):
        s = self.services
        s.register("log", lambda: self.log)
        s.register("doc_utils", self._make_doc_utils)
        s.register("positions", self._make_positions)
        s.register("ui", self._make_ui)

    # service factories import uno-dependent modules lazily
    def _make_doc_utils(self):
        from office_plugin.services.doc_utils import DocUtils
        return DocUtils(self.ctx)

    def _make_positions(self):
        from office_plugin.services.positions import PositionService
        return PositionService(self.ctx, self.services.get("doc_utils"), self.log)

    def _make_ui(self):
        from office_plugin.services.ui import UiService
        return UiService(self.ctx, self.services.get("doc_utils"))

    def start(self):
        self.log.info("starting Office Plugin")
        self.registry.load()
        for feature in self.registry.features.values():
            feature._notify = lambda fid, command=None: self.status_notifier(fid, command)
        if self.ctx is not None:
            from .uno_events import UnoEventHub
            self.hub = UnoEventHub(self.ctx, self.events, self.services.get("doc_utils"), self.log)
            try:
                self.hub.start()
            except Exception:
                self.log.exception("could not start UNO event hub")
        if self.registry.failed:
            self.log.warning("features failed to load: %s", self.registry.failed)

    def stop(self):
        if self.hub:
            self.hub.stop()
        self.registry.stop()


def _profile_dir(ctx):
    try:
        import uno
        subst = ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.util.PathSubstitution", ctx)
        return uno.fileUrlToSystemPath(subst.substituteVariables("$(user)", True))
    except Exception:
        return None


def ensure_started(ctx):
    """Idempotent; safe to call from the protocol handler and from the startup job."""
    global _plugin
    if _plugin is None:
        _plugin = Plugin(ctx, _profile_dir(ctx))
        _plugin.start()
    return _plugin
