import os
import unittest

import tests  # noqa: F401  (sets sys.path)
from office_plugin.core import manifest as M

FEATURES_DIR = os.path.join(os.path.dirname(M.__file__), "..", "features")


def good():
    return {
        "id": "demo", "name": "Demo", "version": "1.0", "class": "Demo", "apps": ["calc"],
        "commands": [{"name": "Run", "label": "Run", "icon": "demo/run", "toolbar": True,
                      "accelerator": {"calc": "Alt+F5"}}],
    }


class AcceleratorKeyTests(unittest.TestCase):
    def test_lo_names(self):
        self.assertEqual(M.accelerator_to_lo_key("Alt+Left"), "LEFT_MOD2")
        self.assertEqual(M.accelerator_to_lo_key("Alt+;"), "SEMICOLON_MOD2")
        self.assertEqual(M.accelerator_to_lo_key("Ctrl+Shift+Backspace"), "BACKSPACE_SHIFT_MOD1")
        self.assertEqual(M.accelerator_to_lo_key("Alt+F5"), "F5_MOD2")
        self.assertEqual(M.accelerator_to_lo_key("Ctrl+Alt+Shift+F5"), "F5_SHIFT_MOD1_MOD2")
        self.assertEqual(M.accelerator_to_lo_key("F5"), "F5")
        self.assertEqual(M.accelerator_to_lo_key("Ctrl+A"), "A_MOD1")

    def test_bad_input(self):
        for bad in ("Hyper+A", "Ctrl+", ""):
            with self.assertRaises(M.ManifestError):
                M.accelerator_to_lo_key(bad)


class ValidateTests(unittest.TestCase):
    def test_good_manifest_gets_defaults(self):
        m = M.validate(good())
        self.assertTrue(m["default_enabled"])
        self.assertEqual(m["commands"][0]["group"], "demo")
        self.assertTrue(m["commands"][0]["menu"])

    def test_rejections(self):
        def broken(mutate):
            d = good()
            mutate(d)
            return d
        cases = {
            "missing id": broken(lambda d: d.pop("id")),
            "bad id": broken(lambda d: d.update(id="Bad-Id")),
            "unknown top key": broken(lambda d: d.update(oops=1)),
            "unknown app": broken(lambda d: d.update(apps=["math"])),
            "empty apps": broken(lambda d: d.update(apps=[])),
            "unknown command key": broken(lambda d: d["commands"][0].update(oops=1)),
            "toolbar without icon": broken(lambda d: d["commands"][0].update(icon=None)),
            "duplicate command": broken(lambda d: d["commands"].append(dict(d["commands"][0]))),
            "accelerator for app not in apps":
                broken(lambda d: d["commands"][0].update(accelerator={"writer": "Ctrl+B"})),
            "bad accelerator": broken(lambda d: d["commands"][0].update(accelerator={"calc": "Foo+B"})),
        }
        for label, data in cases.items():
            with self.subTest(label), self.assertRaises(M.ManifestError):
                M.validate(data)

    def test_url_roundtrip(self):
        url = M.command_url("navigation", "Back")
        self.assertEqual(M.parse_command_url(url), ("navigation", "Back"))
        self.assertIsNone(M.parse_command_url(".uno:Save"))
        self.assertIsNone(M.parse_command_url("vnd.officeplugin:nodot"))


class AcceleratorConflictTests(unittest.TestCase):
    def test_conflict_between_features_is_detected(self):
        a, b = M.validate(good()), M.validate(good())
        b["id"] = "other"
        with self.assertRaises(M.ManifestError):
            M.accelerators([a, b])

    def test_same_key_different_module_is_fine(self):
        a, b = M.validate(good()), M.validate(good())
        b["id"], b["apps"] = "other", ["writer"]
        b["commands"][0]["accelerator"] = {"writer": "Alt+F5"}
        self.assertEqual(len(M.accelerators([a, b])), 2)

    def test_all_expands_to_every_app(self):
        m = good()
        m["apps"] = ["writer", "calc"]
        m["commands"][0]["accelerator"] = {"all": "Alt+Left"}
        mods = {mod for mod, _, _ in M.accelerators([M.validate(m)])}
        self.assertEqual(mods, {M.APPS["writer"], M.APPS["calc"]})


class RealManifestTests(unittest.TestCase):
    def test_shipped_features_are_valid_and_conflict_free(self):
        manifests = M.discover_manifests(FEATURES_DIR)
        ids = {m["id"] for m in manifests}
        self.assertEqual(ids, {"navigation", "last_edit", "calc_visible", "calc_dedupe", "calc_setops", "calc_fill_blanks",
                         "calc_remove_blanks", "calc_highlight", "calc_transpose",
                         "calc_sequence", "calc_missing_seq", "calc_gantt",
                         "calc_decimals", "calc_numwords", "calc_units", "calc_functions"})
        M.accelerators(manifests)

    def test_template_is_a_valid_manifest_and_not_discovered(self):
        M.load_manifest(os.path.join(FEATURES_DIR, "_template", "feature.json"))
        self.assertNotIn("my_feature", {m["id"] for m in M.discover_manifests(FEATURES_DIR)})


if __name__ == "__main__":
    unittest.main()
