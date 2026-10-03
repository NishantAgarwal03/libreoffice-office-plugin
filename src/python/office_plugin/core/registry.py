"""Discovers, validates and instantiates features from the `features` package.

A feature that fails to import/construct/start is logged and skipped; the rest keep
working.
"""

import importlib
import os

from . import manifest as manifest_mod


class Registry:
    def __init__(self, ctx, services, events, settings, log,
                 package="office_plugin.features", features_dir=None):
        self.ctx, self.services, self.events = ctx, services, events
        self.settings, self.log = settings, log
        self.package = package
        if features_dir is None:
            pkg = importlib.import_module(package)
            features_dir = os.path.dirname(pkg.__file__)
        self.features_dir = features_dir
        self.features = {}   # id -> Feature instance (started)
        self.manifests = {}  # id -> manifest dict (all discovered, even disabled)
        self.failed = {}     # id -> error text

    def load(self):
        for name in sorted(os.listdir(self.features_dir)):
            path = os.path.join(self.features_dir, name, "feature.json")
            if name.startswith("_") or not os.path.isfile(path):
                continue
            try:
                man = manifest_mod.load_manifest(path)
                if man["id"] != name:
                    raise manifest_mod.ManifestError("id %r != folder %r" % (man["id"], name))
                self.manifests[name] = man
                if not self.settings.feature_enabled(name, man["default_enabled"]):
                    self.log.info("feature %s disabled in settings", name)
                    continue
                missing = [s for s in man["requires"] if not self.services.has(s)]
                if missing:
                    raise manifest_mod.ManifestError("missing services %s" % missing)
                module = importlib.import_module("%s.%s" % (self.package, name))
                cls = getattr(module, man["class"])
                feature = cls(self.ctx, self.services, self.events, self.settings, man, self.log)
                feature.start()
                self.features[name] = feature
                self.log.info("feature %s started", name)
            except Exception as exc:  # isolate: one broken feature must not take down the rest
                self.failed[name] = "%s: %s" % (type(exc).__name__, exc)
                self.log.exception("feature %s failed to load", name)

    def stop(self):
        for feature in self.features.values():
            try:
                feature.stop()
            except Exception:
                self.log.exception("feature %s failed to stop", feature.id)
        self.features.clear()

    def get(self, feature_id):
        return self.features.get(feature_id)
