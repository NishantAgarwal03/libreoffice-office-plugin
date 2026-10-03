"""Guards every feature (present and future) against shadowing a built-in LibreOffice key.

Parses the registry files of the local LibreOffice installation and compares them with the
accelerators build_oxt.py would generate. Skipped when LibreOffice is not installed.
"""

import collections
import glob
import os
import unittest
import xml.etree.ElementTree as ET

import tests  # noqa: F401
import build_oxt
from office_plugin.core import manifest as M

OOR = "{http://openoffice.org/2001/registry}"
LO_REGISTRY = os.environ.get("LO_REGISTRY", r"C:\Program Files\LibreOffice\share\registry")
FEATURES_DIR = os.path.join(build_oxt.PYDIR, "office_plugin", "features")


def shipped_bindings():
    """-> {module_or_'Global': {lo_key: command}} from LibreOffice's own registry."""
    binds = collections.defaultdict(dict)
    for path in glob.glob(os.path.join(LO_REGISTRY, "*.xcd")):
        try:
            root = ET.parse(path).getroot()
        except ET.ParseError:
            continue
        for comp in root.iter(OOR + "component-data"):
            if comp.get(OOR + "name") != "Accelerators":
                continue
            for primary in comp.findall("node"):
                if primary.get(OOR + "name") != "PrimaryKeys":
                    continue
                for group in primary.findall("node"):
                    if group.get(OOR + "name") == "Global":
                        scopes = [("Global", group)]
                    else:
                        scopes = [(n.get(OOR + "name"), n) for n in group.findall("node")]
                    for scope, node in scopes:
                        for key in node.findall("node"):
                            command = key.findtext("prop/value")
                            if command:
                                binds[scope][key.get(OOR + "name")] = command
    return binds


@unittest.skipUnless(os.path.isdir(LO_REGISTRY), "LibreOffice registry not found")
class KeyConflictTests(unittest.TestCase):
    def test_no_feature_shadows_a_builtin_binding(self):
        shipped = shipped_bindings()
        self.assertTrue(shipped, "could not read LibreOffice's shipped accelerators")
        clashes = []
        for module, lo_key, url in M.accelerators(M.discover_manifests(FEATURES_DIR)):
            existing = shipped.get(module, {}).get(lo_key) or shipped.get("Global", {}).get(lo_key)
            if existing:
                clashes.append("%s in %s is already %s (wanted by %s)" % (lo_key, module, existing, url))
        self.assertEqual(clashes, [], "\n".join(clashes))

    def test_the_check_itself_detects_a_known_builtin(self):
        # Ctrl+Shift+Backspace really is "delete to start of sentence" in Writer: the guard must see it.
        shipped = shipped_bindings()
        self.assertIn("BACKSPACE_SHIFT_MOD1", shipped.get(M.APPS["writer"], {}))


if __name__ == "__main__":
    unittest.main()
