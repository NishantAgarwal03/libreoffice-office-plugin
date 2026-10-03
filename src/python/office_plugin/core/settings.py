"""Per-feature settings persisted as JSON in the user profile.

Layout of the file:
    {"debug": false,
     "features": {"navigation": {"enabled": true, "options": {...}}, ...}}
"""

import json
import os


class Settings:
    def __init__(self, path=None):
        self._path = path
        self._data = {"debug": False, "features": {}}
        self.load()

    def load(self):
        if not self._path or not os.path.exists(self._path):
            return
        try:
            with open(self._path, "r", encoding="utf-8") as fh:
                loaded = json.load(fh)
            if isinstance(loaded, dict):
                self._data.update(loaded)
                self._data.setdefault("features", {})
        except (OSError, ValueError):
            pass  # corrupt/unreadable settings -> defaults

    def save(self):
        if not self._path:
            return
        try:
            os.makedirs(os.path.dirname(self._path), exist_ok=True)
            with open(self._path, "w", encoding="utf-8") as fh:
                json.dump(self._data, fh, indent=2)
        except OSError:
            pass

    @property
    def debug(self):
        return bool(self._data.get("debug", False))

    def feature_enabled(self, feature_id, default=True):
        return bool(self._data["features"].get(feature_id, {}).get("enabled", default))

    def set_feature_enabled(self, feature_id, enabled):
        self._data["features"].setdefault(feature_id, {})["enabled"] = bool(enabled)

    def option(self, feature_id, key, default=None):
        return self._data["features"].get(feature_id, {}).get("options", {}).get(key, default)

    def set_option(self, feature_id, key, value):
        self._data["features"].setdefault(feature_id, {}).setdefault("options", {})[key] = value
