import copy
import os
import shutil
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

import tests  # noqa: F401
import build_oxt
from office_plugin.core import manifest as M

OOR = "{http://openoffice.org/2001/registry}"
FEATURES_DIR = os.path.join(build_oxt.PYDIR, "office_plugin", "features")


def shipped():
    return M.discover_manifests(FEATURES_DIR)


def hello():
    return M.validate({
        "id": "hello", "name": "Hello", "version": "1.0", "class": "Hello", "apps": ["writer"],
        "commands": [{"name": "Say", "label": "Say Hello", "icon": "hello/say", "toolbar": True,
                      "group": "fun", "order": 900,
                      "accelerator": {"writer": "Ctrl+Alt+H"}}]})


def items(xml_text, node_name):
    """-> [(url, contexts)] children of the Addons.xcu node `node_name` in document order."""
    root = ET.fromstring(xml_text)
    out = []
    for node in root.iter("node"):
        if node.get(OOR + "name") == node_name:
            for child in node.iter("node"):
                props = {p.get(OOR + "name"): (p.findtext("value") or "") for p in child.findall("prop")}
                if "URL" in props:
                    out.append((props["URL"], props.get("Context", "")))
            break
    return out


class GeneratedXmlTests(unittest.TestCase):
    def test_every_generated_file_is_well_formed(self):
        for name, text in build_oxt.generated_files(shipped()).items():
            if name.endswith((".xml", ".xcu")):
                ET.fromstring(text)

    def test_toolbar_groups_are_separated_and_ordered(self):
        tb = items(build_oxt.gen_addons(shipped()), "org.officeplugin.toolbar")
        urls = [u for u, _ in tb]
        self.assertEqual(urls, [
            "vnd.officeplugin:navigation.Back", "vnd.officeplugin:navigation.Forward",
            "vnd.officeplugin:last_edit.GoTo", "private:separator",
            "vnd.officeplugin:calc_visible.Select", "vnd.officeplugin:calc_dedupe.Run",
            "vnd.officeplugin:calc_setops.Run", "vnd.officeplugin:calc_transpose.Run",
            "vnd.officeplugin:calc_remove_blanks.Run",
            "vnd.officeplugin:calc_fill_blanks.Run", "vnd.officeplugin:calc_sequence.Run",
            "vnd.officeplugin:calc_missing_seq.Run", "private:separator",
            "vnd.officeplugin:calc_highlight.Run", "vnd.officeplugin:calc_highlight.Clear", "private:separator",
            "vnd.officeplugin:calc_gantt.Run", "vnd.officeplugin:calc_gantt.Remove", "private:separator",
            "vnd.officeplugin:calc_decimals.Run", "vnd.officeplugin:calc_numwords.ToWords",
            "vnd.officeplugin:calc_numwords.FromWords", "vnd.officeplugin:calc_units.Run",
            "vnd.officeplugin:calc_functions.Manage"])

    def test_separator_only_where_both_neighbours_are_visible(self):
        tb = dict()
        for url, ctx in items(build_oxt.gen_addons(shipped()), "org.officeplugin.toolbar"):
            tb[url] = ctx
        # navigation group is in all apps, calc group only in calc -> separator is calc-only
        self.assertEqual(tb["private:separator"], M.APPS["calc"])
        self.assertEqual(tb["vnd.officeplugin:calc_dedupe.Run"], M.APPS["calc"])
        self.assertNotIn(M.APPS["calc"], tb["vnd.officeplugin:last_edit.GoTo"].split(",")[2:])

    def test_menu_mirrors_toolbar_commands(self):
        addons = build_oxt.gen_addons(shipped())
        menu = [u for u, _ in items(addons, "org.officeplugin.menu") if u != "private:separator"]
        tool = [u for u, _ in items(addons, "org.officeplugin.toolbar") if u != "private:separator"]
        self.assertEqual(menu, tool)

    def test_accelerators_for_expected_keys(self):
        root = ET.fromstring(build_oxt.gen_accelerators(shipped()))
        found = {}
        for module in root.iter("node"):
            if module.get(OOR + "name", "").startswith("com.sun.star."):
                for key in module:
                    found[(module.get(OOR + "name"), key.get(OOR + "name"))] = key.findtext("prop/value")
        self.assertEqual(found[(M.APPS["calc"], "SEMICOLON_MOD2")], "vnd.officeplugin:calc_visible.Select")
        self.assertEqual(found[(M.APPS["writer"], "LEFT_MOD2")], "vnd.officeplugin:navigation.Back")
        self.assertEqual(found[(M.APPS["impress"], "RIGHT_MOD2")], "vnd.officeplugin:navigation.Forward")


class ModularityTests(unittest.TestCase):
    """Adding a feature must need no edit outside its own folder."""

    def test_new_feature_appears_in_all_generated_registration(self):
        manifests = shipped() + [hello()]
        addons = build_oxt.gen_addons(manifests)
        self.assertIn("vnd.officeplugin:hello.Say", [u for u, _ in items(addons, "org.officeplugin.toolbar")])
        self.assertIn("vnd.officeplugin:hello.Say", [u for u, _ in items(addons, "org.officeplugin.menu")])
        self.assertIn("hello/say_16.png", addons.replace("icons/", ""))
        self.assertIn("vnd.officeplugin:hello.Say", build_oxt.gen_accelerators(manifests))
        # the new group (order 900) sorts after the existing ones, separated from them
        urls = [u for u, _ in items(addons, "org.officeplugin.toolbar")]
        self.assertEqual(urls[-1], "vnd.officeplugin:hello.Say")

    def test_conflicting_new_feature_fails_the_build(self):
        clash = hello()
        clash["apps"] = ["calc"]
        clash["commands"][0]["accelerator"] = {"calc": "Alt+;"}
        with self.assertRaises(M.ManifestError):
            build_oxt.gen_accelerators(shipped() + [clash])

    def test_feature_folder_dropped_into_a_copy_is_built(self):
        with tempfile.TemporaryDirectory() as tmp:
            copy_dir = os.path.join(tmp, "features")
            shutil.copytree(FEATURES_DIR, copy_dir, ignore=shutil.ignore_patterns("__pycache__"))
            os.makedirs(os.path.join(copy_dir, "hello"))
            import json
            with open(os.path.join(copy_dir, "hello", "feature.json"), "w") as fh:
                json.dump({"id": "hello", "name": "Hello", "version": "1", "class": "Hello",
                           "apps": ["writer"],
                           "commands": [{"name": "Say", "label": "Say"}]}, fh)
            ids = [m["id"] for m in M.discover_manifests(copy_dir)]
            self.assertIn("hello", ids)
            self.assertNotIn("my_feature", ids)     # _template stays excluded


class WindowStateTests(unittest.TestCase):
    def test_toolbar_gets_a_title_and_is_docked_per_app(self):
        files = build_oxt.window_state_files(shipped())
        self.assertEqual(sorted(files), ["WindowState_calc.xcu", "WindowState_draw.xcu", "WindowState_impress.xcu",
                                         "WindowState_writer.xcu"])
        root = ET.fromstring(files["WindowState_calc.xcu"])
        self.assertEqual(root.get("{http://openoffice.org/2001/registry}name"), "CalcWindowState")
        text = files["WindowState_calc.xcu"]
        for needle in ("private:resource/toolbar/addon_org.officeplugin.toolbar", ">Office Plugin<",
                       'oor:name="Docked"', 'oor:name="DockingArea"'):
            self.assertIn(needle, text)

    def test_manifest_registers_every_window_state_file(self):
        manifest = build_oxt.gen_manifest_xml(shipped())
        for name in build_oxt.window_state_files(shipped()):
            self.assertIn(name, manifest)


class BuildOutputTests(unittest.TestCase):
    def test_oxt_contents(self):
        with tempfile.TemporaryDirectory() as tmp:
            out, manifests = build_oxt.build(os.path.join(tmp, "t.oxt"))
            names = set(zipfile.ZipFile(out).namelist())
        for required in ("META-INF/manifest.xml", "description.xml", "Addons.xcu",
                         "Accelerators.xcu", "ProtocolHandler.xcu", "Jobs.xcu",
                         "WindowState_calc.xcu", "python/officeplugin.py", "python/office_plugin/core/registry.py",
                         "icons/navigation/back_16.png", "icons/navigation/back_26.png"):
            self.assertIn(required, names)
        self.assertFalse([n for n in names if "_template" in n or "__pycache__" in n])
        self.assertEqual(len(manifests), 16)

    def test_missing_icon_fails_the_build(self):
        broken = hello()
        broken["commands"][0]["icon"] = "hello/does_not_exist"
        with self.assertRaises(SystemExit):
            build_oxt.check_icons([broken])


if __name__ == "__main__":
    unittest.main()
