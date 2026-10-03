"""Base class every feature subclasses.

Rules (see README "How to add a feature"):
  * depend only on office_plugin.core.* and office_plugin.services.*, never on another feature
  * do not import `uno` at module level (lazy-import inside methods)
  * talk to other parts of the plugin through `self.events` / `self.services`
"""


class Feature:
    def __init__(self, ctx, services, events, settings, manifest, log):
        self.ctx = ctx                # com.sun.star.uno.XComponentContext (None in unit tests)
        self.services = services
        self.events = events
        self.settings = settings
        self.manifest = manifest
        self.log = log
        self.id = manifest["id"]
        self._subs = []
        self._notify = lambda feature_id, command=None: None  # injected by the plugin

    # ---- lifecycle -------------------------------------------------------
    def start(self):
        """Subscribe to events here (use self.subscribe so stop() can undo it)."""

    def stop(self):
        for token in self._subs:
            self.events.unsubscribe(token)
        self._subs = []

    # ---- commands --------------------------------------------------------
    def execute(self, command, frame, args):
        """Handle vnd.officeplugin:<id>.<command>."""
        raise NotImplementedError("%s.%s" % (self.id, command))

    def is_enabled(self, command, frame):
        """Whether the command's button/menu entry is enabled for `frame`."""
        return True

    # ---- helpers ---------------------------------------------------------
    def subscribe(self, event_name, callback):
        self._subs.append(self.events.subscribe(event_name, callback))

    def option(self, key, default=None):
        return self.settings.option(self.id, key, default)

    def refresh_state(self, command=None):
        """Ask LibreOffice to re-query enabled state of our toolbar/menu entries."""
        self._notify(self.id, command)
