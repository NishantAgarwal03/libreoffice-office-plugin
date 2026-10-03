"""Routes vnd.officeplugin:<feature>.<command> to the owning feature.

Pure Python. officeplugin.py wraps this in the UNO XDispatch objects.
"""

from . import manifest as manifest_mod


class Router:
    def __init__(self, registry, log):
        self.registry = registry
        self.log = log

    def resolve(self, url):
        """-> (feature, command_name) or None when the URL is not ours/unknown."""
        parsed = manifest_mod.parse_command_url(url)
        if not parsed:
            return None
        feature = self.registry.get(parsed[0])
        if feature is None:
            return None
        if not any(c["name"] == parsed[1] for c in feature.manifest["commands"]):
            return None
        return feature, parsed[1]

    def execute(self, url, frame, args=()):
        target = self.resolve(url)
        if target is None:
            self.log.warning("no handler for %s", url)
            return False
        feature, command = target
        try:
            feature.execute(command, frame, args)
            return True
        except Exception:
            self.log.exception("command %s failed", url)
            return False

    def is_enabled(self, url, frame, app=None):
        """`app` is the active document kind ('writer'/'calc'/...), or None if unknown."""
        target = self.resolve(url)
        if target is None:
            return False
        feature, command = target
        if app is not None and app not in feature.manifest["apps"]:
            return False
        try:
            return bool(feature.is_enabled(command, frame))
        except Exception:
            self.log.exception("is_enabled failed for %s", url)
            return False
