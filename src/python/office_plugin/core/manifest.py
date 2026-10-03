"""feature.json loading + validation + helpers shared by the runtime and build_oxt.py.

A feature.json is the single source of truth for everything LibreOffice needs to
know about a feature (menu/toolbar entries, key bindings, which apps it applies to).
"""

import json
import os
import re

APPS = {
    "writer": "com.sun.star.text.TextDocument",
    "calc": "com.sun.star.sheet.SpreadsheetDocument",
    "impress": "com.sun.star.presentation.PresentationDocument",
    "draw": "com.sun.star.drawing.DrawingDocument",
}

_ID_RE = re.compile(r"^[a-z][a-z0-9_]*$")
_CMD_RE = re.compile(r"^[A-Za-z][A-Za-z0-9]*$")

_FEATURE_KEYS = {"id", "name", "version", "class", "apps", "requires", "default_enabled",
                 "description", "commands", "sidebar"}
_COMMAND_KEYS = {"name", "label", "tooltip", "icon", "group", "order", "toolbar", "menu",
                 "accelerator"}

# modifier name -> LibreOffice key-name suffix, in LibreOffice's canonical order
_MODS = (("shift", "SHIFT"), ("ctrl", "MOD1"), ("alt", "MOD2"))

# human key name -> LibreOffice key name (everything else is upper-cased as is)
_KEY_ALIASES = {
    ";": "SEMICOLON", ":": "COLON", ",": "COMMA", ".": "POINT", "-": "SUBTRACT",
    "+": "ADD", "=": "EQUAL", "/": "DIVIDE", "[": "LEFTBRACKET", "]": "RIGHTBRACKET",
    "`": "QUOTELEFT", "PGUP": "PAGEUP", "PGDN": "PAGEDOWN", "DEL": "DELETE",
    "ESC": "ESCAPE", "RETURN": "RETURN", "ENTER": "RETURN", "BKSP": "BACKSPACE",
}


class ManifestError(ValueError):
    pass


def accelerator_to_lo_key(text):
    """'Ctrl+Shift+Backspace' -> 'BACKSPACE_SHIFT_MOD1' (LibreOffice naming)."""
    if text == "+" or text.endswith("++"):          # the literal '+' key: 'Ctrl++'
        parts = [p.strip() for p in text[:-1].split("+") if p.strip()] + ["+"]
    else:
        parts = [p.strip() for p in text.split("+")]
    if not parts or not parts[-1]:
        raise ManifestError("bad accelerator %r" % text)
    key, mods = parts[-1], {p.lower() for p in parts[:-1]}
    unknown = mods - {"ctrl", "alt", "shift"}
    if unknown:
        raise ManifestError("unknown modifier(s) %s in %r" % (sorted(unknown), text))
    name = _KEY_ALIASES.get(key.upper(), _KEY_ALIASES.get(key, key.upper()))
    if not re.match(r"^[A-Z0-9_]+$", name):
        raise ManifestError("unsupported key %r in %r" % (key, text))
    return "_".join([name] + [lo for word, lo in _MODS if word in mods])


def load_manifest(path):
    with open(path, "r", encoding="utf-8") as fh:
        try:
            data = json.load(fh)
        except ValueError as exc:
            raise ManifestError("%s: invalid JSON: %s" % (path, exc))
    return validate(data, source=path)


def validate(data, source="<manifest>"):
    def fail(msg):
        raise ManifestError("%s: %s" % (source, msg))

    if not isinstance(data, dict):
        fail("top level must be an object")
    unknown = set(data) - _FEATURE_KEYS
    if unknown:
        fail("unknown key(s) %s" % sorted(unknown))
    for key in ("id", "name", "version", "class", "apps", "commands"):
        if key not in data:
            fail("missing required key %r" % key)
    if not _ID_RE.match(data["id"]):
        fail("id %r must match %s" % (data["id"], _ID_RE.pattern))
    if not isinstance(data["apps"], list) or not data["apps"]:
        fail("apps must be a non-empty list")
    bad = [a for a in data["apps"] if a not in APPS]
    if bad:
        fail("unknown app(s) %s (allowed: %s)" % (bad, sorted(APPS)))
    data.setdefault("requires", [])
    data.setdefault("default_enabled", True)
    data.setdefault("description", "")

    seen = set()
    for cmd in data["commands"]:
        if not isinstance(cmd, dict):
            fail("each command must be an object")
        unknown = set(cmd) - _COMMAND_KEYS
        if unknown:
            fail("command %r: unknown key(s) %s" % (cmd.get("name"), sorted(unknown)))
        for key in ("name", "label"):
            if key not in cmd:
                fail("command missing required key %r" % key)
        if not _CMD_RE.match(cmd["name"]):
            fail("command name %r must match %s" % (cmd["name"], _CMD_RE.pattern))
        if cmd["name"] in seen:
            fail("duplicate command %r" % cmd["name"])
        seen.add(cmd["name"])
        cmd.setdefault("tooltip", "")
        cmd.setdefault("icon", None)
        cmd.setdefault("group", data["id"])
        cmd.setdefault("order", 100)
        cmd.setdefault("toolbar", False)
        cmd.setdefault("menu", True)
        cmd.setdefault("accelerator", None)
        if cmd["toolbar"] and not cmd["icon"]:
            fail("command %r is on the toolbar but has no icon" % cmd["name"])
        acc = cmd["accelerator"]
        if acc is not None:
            if not isinstance(acc, dict):
                fail("command %r: accelerator must be an object {app: key}" % cmd["name"])
            for app, keytext in acc.items():
                if app != "all" and app not in data["apps"]:
                    fail("command %r: accelerator for app %r not in apps" % (cmd["name"], app))
                if keytext is not None:
                    accelerator_to_lo_key(keytext)  # raises ManifestError if bad
    return data


def command_url(feature_id, command_name):
    return "vnd.officeplugin:%s.%s" % (feature_id, command_name)


def parse_command_url(url):
    """'vnd.officeplugin:navigation.Back' -> ('navigation', 'Back') or None."""
    prefix = "vnd.officeplugin:"
    if not url.startswith(prefix):
        return None
    path = url[len(prefix):].split("?", 1)[0]
    if "." not in path:
        return None
    feature_id, command = path.split(".", 1)
    return feature_id, command


def accelerators(manifests):
    """Flatten all key bindings -> list of (module_service, lo_key, command_url).

    Raises ManifestError if two commands claim the same key in the same module.
    """
    claimed = {}
    out = []
    for man in manifests:
        for cmd in man["commands"]:
            acc = cmd["accelerator"] or {}
            for app, keytext in acc.items():
                if keytext is None:
                    continue
                for target in (man["apps"] if app == "all" else [app]):
                    if target not in man["apps"]:
                        continue
                    module = APPS[target]
                    lo_key = accelerator_to_lo_key(keytext)
                    url = command_url(man["id"], cmd["name"])
                    prior = claimed.get((module, lo_key))
                    if prior and prior != url:
                        raise ManifestError("key %s in %s claimed by both %s and %s"
                                            % (keytext, target, prior, url))
                    claimed[(module, lo_key)] = url
                    if (module, lo_key, url) not in out:
                        out.append((module, lo_key, url))
    return out


def discover_manifests(features_dir):
    """Load every features_dir/<name>/feature.json (skipping names starting '_')."""
    result = []
    for name in sorted(os.listdir(features_dir)):
        path = os.path.join(features_dir, name, "feature.json")
        if name.startswith("_") or not os.path.isfile(path):
            continue
        man = load_manifest(path)
        if man["id"] != name:
            raise ManifestError("%s: id %r must equal its folder name %r" % (path, man["id"], name))
        result.append(man)
    return result
