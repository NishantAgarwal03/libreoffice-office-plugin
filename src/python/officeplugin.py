"""UNO entry point of the Office Plugin extension.

Registers two components:
  * org.officeplugin.ProtocolHandler  -> handles vnd.officeplugin:<feature>.<command> URLs
  * org.officeplugin.StartupJob       -> starts the plugin when LibreOffice shows its first window

Everything else lives in the `office_plugin` package next to this file.
"""

import os
import sys

import uno
import unohelper
from com.sun.star.frame import FeatureStateEvent, XDispatch, XDispatchProvider
from com.sun.star.lang import XInitialization, XServiceInfo
from com.sun.star.task import XJob

# pythonloader puts this file's directory on sys.path; make it explicit & robust.
_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from office_plugin.core import bootstrap  # noqa: E402
from office_plugin.core.manifest import parse_command_url  # noqa: E402

PROTOCOL_IMPL = "org.officeplugin.ProtocolHandler"
JOB_IMPL = "org.officeplugin.StartupJob"
FUNCTION_IMPL = "org.officeplugin.FunctionJob"
PROTOCOL_PREFIX = "vnd.officeplugin:"


class _StatusHub:
    """Keeps toolbar/menu status listeners so features can refresh enabled state."""

    def __init__(self):
        self._listeners = {}  # url -> [(listener, dispatch)]

    def add(self, url, listener, dispatch):
        self._listeners.setdefault(url, []).append((listener, dispatch))

    def remove(self, url, listener):
        self._listeners[url] = [(l, d) for l, d in self._listeners.get(url, []) if l != listener]

    def refresh(self, feature_id, command=None):
        prefix = PROTOCOL_PREFIX + feature_id + "."
        for url, entries in list(self._listeners.items()):
            if url.startswith(prefix) and (command is None or url == prefix + command):
                for listener, dispatch in list(entries):
                    try:
                        listener.statusChanged(dispatch.state(url))
                    except Exception:
                        pass


_hub = _StatusHub()


class _Dispatch(unohelper.Base, XDispatch):
    def __init__(self, plugin, frame, url):
        self.plugin, self.frame, self.url = plugin, frame, url
        self.url_struct = None

    def _app(self):
        try:
            docs = self.plugin.services.get("doc_utils")
            ctrl = docs.controller_of(self.frame)
            return docs.kind(ctrl.getModel()) if ctrl else None
        except Exception:
            return None

    def state(self, url):
        ev = FeatureStateEvent()
        ev.FeatureURL = self.url_struct
        ev.FeatureDescriptor = url
        ev.IsEnabled = self.plugin.router.is_enabled(url, self.frame, self._app())
        ev.Requery = False
        ev.State = None
        return ev

    def dispatch(self, url, args):
        self.plugin.router.execute(url.Complete, self.frame, args)

    def addStatusListener(self, listener, url):
        self.url_struct = url
        _hub.add(url.Complete, listener, self)
        listener.statusChanged(self.state(url.Complete))

    def removeStatusListener(self, listener, url):
        _hub.remove(url.Complete, listener)


class ProtocolHandler(unohelper.Base, XDispatchProvider, XInitialization, XServiceInfo):
    def __init__(self, ctx):
        self.ctx = ctx
        self.frame = None
        self.plugin = bootstrap.ensure_started(ctx)
        self.plugin.status_notifier = _hub.refresh

    # XInitialization: LibreOffice passes the frame this handler serves
    def initialize(self, args):
        if args:
            self.frame = args[0]

    # XDispatchProvider
    def queryDispatch(self, url, target_frame_name, search_flags):
        if not url.Complete.startswith(PROTOCOL_PREFIX):
            return None
        if self.plugin.router.resolve(url.Complete) is None:
            return None
        return _Dispatch(self.plugin, self.frame, url.Complete)

    def queryDispatches(self, requests):
        return tuple(self.queryDispatch(r.FeatureURL, r.FrameName, r.SearchFlags)
                     for r in requests)

    # XServiceInfo
    def getImplementationName(self):
        return PROTOCOL_IMPL

    def supportsService(self, name):
        return name in self.getSupportedServiceNames()

    def getSupportedServiceNames(self):
        return ("com.sun.star.frame.ProtocolHandler",)


class StartupJob(unohelper.Base, XJob, XServiceInfo):
    def __init__(self, ctx):
        self.ctx = ctx

    def execute(self, args):
        plugin = bootstrap.ensure_started(self.ctx)
        plugin.status_notifier = _hub.refresh
        return None

    def getImplementationName(self):
        return JOB_IMPL

    def supportsService(self, name):
        return name in self.getSupportedServiceNames()

    def getSupportedServiceNames(self):
        return ("com.sun.star.task.Job",)


class FunctionJob(unohelper.Base, XJob, XServiceInfo):
    """Called by the Basic module that 'Number Functions' installs in My Macros, so cell formulas
    (=NUMTOWORDS(...), =WORDSTONUM(...), =UNITSTONUM(...)) can reach the Python logic.
    execute(NamedValue fn, a, b, c, d) -> str | float. Bad input returns the TEXT "#VALUE!": LibreOffice turns a
    Basic runtime error or exception into a blank cell and cannot raise a real cell error from Basic."""

    def __init__(self, ctx):
        self.ctx = ctx

    def execute(self, args):
        from office_plugin.services import functions
        values = {nv.Name: nv.Value for nv in args}
        try:
            return functions.call(values.get("fn"), values.get("a"), values.get("b"), values.get("c"), values.get("d"))
        except Exception:
            return functions.ERROR_TEXT

    def getImplementationName(self):
        return FUNCTION_IMPL

    def supportsService(self, name):
        return name in self.getSupportedServiceNames()

    def getSupportedServiceNames(self):
        return (FUNCTION_IMPL,)


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    ProtocolHandler, PROTOCOL_IMPL, ("com.sun.star.frame.ProtocolHandler",))
g_ImplementationHelper.addImplementation(
    StartupJob, JOB_IMPL, ("com.sun.star.task.Job",))
g_ImplementationHelper.addImplementation(
    FunctionJob, FUNCTION_IMPL, (FUNCTION_IMPL,))
