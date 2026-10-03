"""End-to-end smoke test against a REAL (headless) LibreOffice with the .oxt installed.

Run with LibreOffice's own Python so `uno` is importable:

    "C:\\Program Files\\LibreOffice\\program\\python.exe" tests\\headless_smoke.py

What it does: builds the .oxt, installs it into a throw-away profile (your real profile is
never touched), starts soffice headless, then drives documents through the UNO API and
dispatches vnd.officeplugin:* commands exactly like the toolbar/keys do.

What it cannot do: real mouse clicks / key presses. Those are covered by the manual
checklist in README.md.
"""

import argparse
from datetime import date
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LO = r"C:\Program Files\LibreOffice\program"
PORT = 2099
EXT_ID = "org.officeplugin.office_plugin_lo"

import uno  # noqa: E402  (LibreOffice's python)
import unohelper  # noqa: E402
from com.sun.star.beans import PropertyValue  # noqa: E402
from com.sun.star.frame import XStatusListener  # noqa: E402


def pv(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


class StatusProbe(unohelper.Base, XStatusListener):
    def __init__(self):
        self.last = None

    def statusChanged(self, ev):
        self.last = ev

    def disposing(self, ev):
        pass


class Env:
    """Owns the scratch profile + soffice process for the whole test run."""
    profile = None
    proc = None
    desktop = None
    ctx = None


def _profile_url(path):
    return "file:///" + path.replace("\\", "/")


def setup_env(keep_profile=False):
    sys.path.insert(0, ROOT)
    import build_oxt
    oxt, _ = build_oxt.build(os.path.join(tempfile.gettempdir(), "opl_smoke", "office_plugin_lo.oxt"))
    # SHORT path on purpose: unopkg fails beyond Windows' 260-char MAX_PATH
    Env.profile = os.path.join(tempfile.gettempdir(), "oplp")
    shutil.rmtree(Env.profile, ignore_errors=True)
    env_arg = "-env:UserInstallation=" + _profile_url(Env.profile)
    r = subprocess.run([os.path.join(LO, "unopkg.exe"), "add", "--suppress-license", env_arg, oxt],
                       capture_output=True, text=True, timeout=300)
    if r.returncode != 0:
        raise RuntimeError("unopkg add failed:\n%s\n%s" % (r.stdout, r.stderr))
    settings_dir = os.path.join(Env.profile, "user", "office_plugin")
    os.makedirs(settings_dir, exist_ok=True)
    with open(os.path.join(settings_dir, "settings.json"), "w") as fh:
        json.dump({"debug": True}, fh)
    Env.proc = subprocess.Popen(
        [os.path.join(LO, "soffice.exe"), env_arg, "--headless", "--invisible", "--norestore",
         "--nologo", "--accept=socket,host=127.0.0.1,port=%d;urp;" % PORT],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    deadline = time.time() + 90
    while True:
        try:
            Env.ctx = resolver.resolve("uno:socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext" % PORT)
            break
        except Exception:
            if time.time() > deadline:
                raise RuntimeError("soffice did not start")
            time.sleep(0.5)
    Env.desktop = Env.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", Env.ctx)


def teardown_env():
    try:
        Env.desktop.terminate()
    except Exception:
        pass
    if Env.proc:
        try:
            Env.proc.wait(timeout=20)
        except Exception:
            Env.proc.kill()


def log_text():
    path = os.path.join(Env.profile, "user", "office_plugin", "office_plugin.log")
    return open(path, encoding="utf-8").read() if os.path.exists(path) else ""


class Base(unittest.TestCase):
    def new_doc(self, kind, hidden=True):
        props = (pv("Hidden", True),) if hidden else ()
        doc = Env.desktop.loadComponentFromURL("private:factory/" + kind, "_blank", 0, props)
        self.addCleanup(doc.close, True)
        return doc

    def run_cmd(self, doc, command):
        """Dispatch vnd.officeplugin:<command> on the doc's frame (like the toolbar does)."""
        frame = doc.getCurrentController().getFrame()
        helper = Env.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.DispatchHelper", Env.ctx)
        helper.executeDispatch(frame, "vnd.officeplugin:" + command, "", 0, ())

    def status(self, doc, command):
        frame = doc.getCurrentController().getFrame()
        url = uno.createUnoStruct("com.sun.star.util.URL")
        url.Complete = "vnd.officeplugin:" + command
        parser = Env.ctx.ServiceManager.createInstanceWithContext("com.sun.star.util.URLTransformer", Env.ctx)
        _, url = parser.parseStrict(url)
        dispatch = frame.queryDispatch(url, "_self", 0)
        self.assertIsNotNone(dispatch, "no dispatch for " + command)
        probe = StatusProbe()
        dispatch.addStatusListener(probe, url)
        dispatch.removeStatusListener(probe, url)
        return probe.last


class Calc(Base):
    def setUp(self):
        self.doc = self.new_doc("scalc")
        self.ctrl = self.doc.getCurrentController()
        self.doc.getSheets().insertNewByName("Second", 1)
        self.run_cmd(self.doc, "navigation.Back")      # first dispatch also starts the plugin
        time.sleep(0.2)

    def goto(self, ref, sheet=0):
        sh = self.doc.getSheets().getByIndex(sheet)
        self.ctrl.setActiveSheet(sh)
        self.ctrl.select(sh.getCellRangeByName(ref))

    def where(self):
        a = self.ctrl.getSelection().getRangeAddress()
        return (a.Sheet, a.StartRow, a.StartColumn)

    def test_plugin_started_all_features(self):
        text = log_text()
        self.assertIn("feature navigation started", text)
        self.assertIn("feature calc_dedupe started", text)
        self.assertNotIn("failed to load", text)
        self.assertNotIn("Traceback", text)

    def test_back_forward_on_far_jump(self):
        self.goto("A1")
        self.goto("A400")                 # far jump -> A1 recorded
        self.assertEqual(self.where(), (0, 399, 0))
        self.run_cmd(self.doc, "navigation.Back")
        self.assertEqual(self.where(), (0, 0, 0))
        self.run_cmd(self.doc, "navigation.Forward")
        self.assertEqual(self.where(), (0, 399, 0))

    def test_sheet_change_is_a_jump(self):
        self.goto("C5")
        self.goto("B2", sheet=1)
        self.run_cmd(self.doc, "navigation.Back")
        self.assertEqual(self.where(), (0, 4, 2))

    def test_small_moves_are_not_recorded(self):
        self.goto("A1")
        for ref in ("A2", "A3", "B3"):
            self.goto(ref)
        self.goto("D5")
        self.run_cmd(self.doc, "navigation.Back")      # nothing recorded -> stays put
        self.assertEqual(self.where(), (0, 4, 3))

    def test_multi_step_back_then_forward_chain(self):
        for ref in ("A1", "A300", "A700", "A1100"):
            self.goto(ref)
        self.run_cmd(self.doc, "navigation.Back")
        self.assertEqual(self.where(), (0, 699, 0))
        self.run_cmd(self.doc, "navigation.Back")
        self.assertEqual(self.where(), (0, 299, 0))
        self.run_cmd(self.doc, "navigation.Forward")
        self.assertEqual(self.where(), (0, 699, 0))

    def test_back_is_disabled_until_there_is_history(self):
        self.assertFalse(self.status(self.doc, "navigation.Back").IsEnabled)
        self.goto("A1")
        self.goto("A500")
        time.sleep(0.1)
        self.assertTrue(self.status(self.doc, "navigation.Back").IsEnabled)
        self.assertFalse(self.status(self.doc, "navigation.Forward").IsEnabled)

    def test_last_edit_cycles_through_edit_places(self):
        sh = self.doc.getSheets().getByIndex(0)
        for ref in ("A10", "A400", "A800"):
            self.goto(ref)
            sh.getCellRangeByName(ref).setString("edit")
        self.goto("Z1")
        self.run_cmd(self.doc, "last_edit.GoTo")
        self.assertEqual(self.where(), (0, 799, 0))
        self.run_cmd(self.doc, "last_edit.GoTo")
        self.assertEqual(self.where(), (0, 399, 0))
        self.run_cmd(self.doc, "last_edit.GoTo")
        self.assertEqual(self.where(), (0, 9, 0))

    def follow_link_on_cursor(self, doc):
        frame = doc.getCurrentController().getFrame()
        helper = Env.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.DispatchHelper", Env.ctx)
        helper.executeDispatch(frame, ".uno:OpenHyperlinkOnCursor", "", 0, ())
        time.sleep(0.3)

    def link_doc(self, target):
        """A *visible* (headless) Calc doc with a hyperlink in D4 -- link activation needs a view."""
        doc = self.new_doc("scalc", hidden=False)
        doc.getSheets().insertNewByName("Second", 1)
        ctrl = doc.getCurrentController()
        sh = doc.getSheets().getByIndex(0)
        cell = sh.getCellRangeByName("D4")
        field = doc.createInstance("com.sun.star.text.TextField.URL")
        field.URL, field.Representation = target, "link"
        cell.insertTextContent(cell.createTextCursor(), field, False)
        ctrl.setActiveSheet(sh)
        ctrl.select(cell)
        self.run_cmd(doc, "navigation.Back")      # make sure the plugin is attached to this doc
        time.sleep(0.2)
        ctrl.select(cell)
        return doc, ctrl

    @staticmethod
    def at(ctrl):
        a = ctrl.getSelection().getRangeAddress()
        return (a.Sheet, a.StartRow, a.StartColumn)

    def test_real_hyperlink_to_another_sheet_records_the_link_cell(self):
        """The 'hard part': the source (the link cell) must be known although Calc moves first."""
        doc, ctrl = self.link_doc("#Second.B2")
        self.follow_link_on_cursor(doc)
        self.assertEqual(self.at(ctrl), (1, 1, 1))              # landed on Second.B2
        self.run_cmd(doc, "navigation.Back")
        self.assertEqual(self.at(ctrl), (0, 3, 3))              # back on the link cell D4
        self.run_cmd(doc, "navigation.Forward")
        self.assertEqual(self.at(ctrl), (1, 1, 1))

    def test_real_hyperlink_to_a_far_cell_on_the_same_sheet(self):
        doc, ctrl = self.link_doc("#Sheet1.A900:B901")   # NB: LO ignores single-cell same-sheet marks
        self.follow_link_on_cursor(doc)
        self.assertEqual(self.at(ctrl), (0, 899, 0))
        self.run_cmd(doc, "navigation.Back")
        self.assertEqual(self.at(ctrl), (0, 3, 3))

    def test_select_visible_skips_hidden_rows(self):
        sh = self.doc.getSheets().getByIndex(0)
        sh.getRows().getByIndex(2).setPropertyValue("IsVisible", False)   # row 3 hidden
        sh.getRows().getByIndex(3).setPropertyValue("IsVisible", False)   # row 4 hidden
        self.goto("A1")
        self.ctrl.select(sh.getCellRangeByName("A1:A6"))
        self.run_cmd(self.doc, "calc_visible.Select")
        addrs = self.ctrl.getSelection().getRangeAddresses()
        rows = sorted(r for a in addrs for r in range(a.StartRow, a.EndRow + 1))
        self.assertEqual(rows, [0, 1, 4, 5])

    def test_calc_only_commands_disabled_in_writer(self):
        writer = self.new_doc("swriter")
        self.assertFalse(self.status(writer, "calc_dedupe.Run").IsEnabled)
        self.assertFalse(self.status(writer, "calc_visible.Select").IsEnabled)
        self.assertTrue(self.status(self.doc, "calc_dedupe.Run").IsEnabled)


class _FakeUi:
    """Stands in for services.ui so the dialog does not block a headless run."""

    def __init__(self, answer=None, form=None, confirm=True):
        # answer: result of choose_columns; form: dict of overrides applied on top of the form's
        # own defaults (None = user cancelled); confirm: result of confirm()
        self.answer, self.messages, self.prompt_items = answer, [], None
        self.form_overrides, self.confirm_answer = form, confirm
        self.form_calls, self.confirms = [], []

    def choose_columns(self, title, prompt, items, options, frame=None):
        self.prompt_items = items
        return self.answer

    def message(self, title, text, frame=None, kind="info"):
        self.messages.append(text)

    def confirm(self, title, text, frame=None):
        self.confirms.append(text)
        return self.confirm_answer

    def pick_range(self, controller, title, initial, callback):
        return False          # force the typed-address path (the picker needs a real GUI)

    def form(self, title, fields, frame=None, width=240):
        self.form_calls.append((title, fields))
        # a list answers successive dialogs in order (None entry = Cancel; exhausted = Cancel)
        if isinstance(self.form_overrides, list):
            overrides = self.form_overrides.pop(0) if self.form_overrides else None
        else:
            overrides = self.form_overrides
        if overrides is None:
            return None
        values = {}
        for spec in fields:
            kind = spec["type"]
            if kind == "text":
                values[spec["name"]] = str(spec.get("value", ""))
            elif kind == "checkbox":
                values[spec["name"]] = bool(spec.get("value"))
            elif kind == "dropdown":
                values[spec["name"]] = spec.get("value") or spec["items"][0]
            elif kind == "multilist":
                values[spec["name"]] = list(spec.get("selected", range(len(spec["items"]))))
        values.update(overrides)
        return values


def make_feature(feature_id, ui):
    """Instantiate a real feature against the remote LibreOffice (client side) with a fake ui."""
    sys.path.insert(0, os.path.join(ROOT, "src", "python"))
    import importlib
    import logging
    from office_plugin.core import manifest
    from office_plugin.core.events import EventBus
    from office_plugin.core.services import ServiceContainer
    from office_plugin.core.settings import Settings
    from office_plugin.services.doc_utils import DocUtils
    services = ServiceContainer()
    services.register_instance("doc_utils", DocUtils(Env.ctx))
    services.register_instance("ui", ui)
    base = os.path.join(ROOT, "src", "python", "office_plugin", "features", feature_id)
    man = manifest.load_manifest(os.path.join(base, "feature.json"))
    cls = getattr(importlib.import_module("office_plugin.features." + feature_id), man["class"])
    return cls(Env.ctx, services, EventBus(), Settings(None), man, logging.getLogger("smoke"))


class SharedServices(Base):
    """services/ui.build_form and services/calc_ranges against a real LibreOffice."""

    def real_ui(self):
        sys.path.insert(0, os.path.join(ROOT, "src", "python"))
        from office_plugin.services.doc_utils import DocUtils
        from office_plugin.services.ui import UiService
        docs = DocUtils(Env.ctx)
        return UiService(Env.ctx, docs), docs

    def test_form_builds_and_returns_defaults(self):
        ui, docs = self.real_ui()
        fields = [
            {"type": "label", "text": "line one\nline two"},
            {"type": "text", "name": "a", "label": "List A", "value": "$Sheet1.$A$1:$C$9"},
            {"type": "checkbox", "name": "hdr", "label": "Header", "value": True},
            {"type": "checkbox", "name": "off", "label": "Off", "value": False},
            {"type": "dropdown", "name": "op", "label": "Operation", "items": ["Union", "A - B"],
             "value": "A - B"},
            {"type": "multilist", "name": "cols", "label": "Columns", "items": ["A", "B", "C"],
             "selected": [0, 2], "rows": 3},
        ]
        dialog, named = ui.build_form("Test", fields)
        toolkit = docs.create("com.sun.star.awt.Toolkit")
        dialog.createPeer(toolkit, None)
        try:
            values = ui.read_form(dialog, named)
        finally:
            dialog.dispose()
        self.assertEqual(values, {"a": "$Sheet1.$A$1:$C$9", "hdr": True, "off": False,
                                  "op": "A - B", "cols": [0, 2]})

    def test_range_parsing_and_addresses(self):
        sys.path.insert(0, os.path.join(ROOT, "src", "python"))
        from office_plugin.services import calc_ranges as R
        doc = self.new_doc("scalc")
        doc.getSheets().insertNewByName("My Sheet", 1)
        ctrl = doc.getCurrentController()
        sh0, sh1 = doc.getSheets().getByIndex(0), doc.getSheets().getByIndex(1)
        sh1.getCellRangeByName("B2").setString("hello")
        rng = R.range_from_text(doc, ctrl, "'My Sheet'.A1:C3")
        self.assertEqual(rng.getRangeAddress().Sheet, 1)
        self.assertEqual(R.data_grid(rng)[1][1], "hello")
        self.assertEqual(R.range_from_text(doc, ctrl, "$A$1:$B$2").getRangeAddress().Sheet, 0)
        self.assertEqual(R.address_text(rng), "$'My Sheet'.$A$1:$C$3")
        self.assertEqual(R.range_from_text(doc, ctrl, R.address_text(rng)).getRangeAddress().EndColumn, 2)
        for bad in ("", "Nope.A1", "Sheet1.ZZZZZZ9999999", "not a range"):
            with self.assertRaises(ValueError, msg=bad):
                R.range_from_text(doc, ctrl, bad)

    def test_region_selection_undo_context_and_sheet_helpers(self):
        sys.path.insert(0, os.path.join(ROOT, "src", "python"))
        from office_plugin.services import calc_ranges as R
        doc = self.new_doc("scalc")
        ctrl = doc.getCurrentController()
        sh = doc.getSheets().getByIndex(0)
        for r in range(3):
            for c in range(2):
                sh.getCellByPosition(c, r).setString("x")
        ctrl.select(sh.getCellRangeByName("A1"))
        region = R.region_from_selection(ctrl)
        self.assertEqual(R.address_text(region), "$Sheet1.$A$1:$B$3")
        with R.undo_context(doc, "two edits"):
            sh.getCellRangeByName("D1").setString("1")
            sh.getCellRangeByName("D2").setString("2")
        doc.getUndoManager().undo()
        self.assertEqual(sh.getCellRangeByName("D1").getString() + sh.getCellRangeByName("D2").getString(), "")
        new = R.add_sheet_after(doc, "Union", 0)
        self.assertEqual(new.Name, "Union")
        self.assertEqual(R.add_sheet_after(doc, "Union", 0).Name, "Union_2")
        self.assertEqual(list(doc.getSheets().getElementNames())[:2], ["Sheet1", "Union_2"])

    def test_range_picker_is_available_on_a_calc_view(self):
        doc = self.new_doc("scalc", hidden=False)
        ctrl = doc.getCurrentController()
        # the click-to-pick wizard needs this API; report (not fail) if a build lacks it
        self.assertTrue(hasattr(ctrl, "startRangeSelection"), "no XRangeSelection on the Calc view")


class CalcSetops(Base):
    """Compare Lists on real Calc data: A on Sheet1 A1:C7, B on Sheet2 A1:C5."""

    def setUp(self):
        self.doc = self.new_doc("scalc", hidden=False)   # undo of sheet insertion needs a real view
        self.doc.getSheets().insertNewByName("Sheet2", 1)
        self.ctrl = self.doc.getCurrentController()
        self.s1, self.s2 = self.doc.getSheets().getByIndex(0), self.doc.getSheets().getByIndex(1)
        rows_a = [("Code", "Item", "Qty"), ("T-01", "Ceramic Tile", 10), ("T-02", "Paint", 5),
                  ("T-03", "Putty", 2), ("T-02", "Paint", 5), (101, "Bolt", 7), ("", "", "")]
        rows_b = [("Code", "Item", "Qty"), ("t-02 ", "paint", 5), ("T-04", "Cement", 1),
                  ("101", "bolt", 7), ("T-01", "ceramic tile", 9)]
        for sheet, rows in ((self.s1, rows_a), (self.s2, rows_b)):
            for r, row in enumerate(rows):
                for c, v in enumerate(row):
                    cell = sheet.getCellByPosition(c, r)
                    cell.setValue(v) if isinstance(v, (int, float)) else cell.setString(v)

    A = "$Sheet1.$A$1:$C$7"
    B = "$Sheet2.$A$1:$C$5"

    def run_op(self, op, cols=(0, 1, 2), confirm=True, **options):
        form1 = {"a": self.A, "b": self.B, "header": True, "op": op}
        form1.update(options)
        ui = _FakeUi(form=[form1, {"cols": list(cols)}], confirm=confirm)
        make_feature("calc_setops", ui).execute("Run", self.ctrl.getFrame(), ())
        return ui

    def sheet_rows(self, name, ncols=3, nrows=8):
        sh = self.doc.getSheets().getByName(name)
        rows = [tuple(sh.getCellByPosition(c, r).getString() for c in range(ncols)) for r in range(nrows)]
        return [r for r in rows if any(r)]

    def names(self):
        return list(self.doc.getSheets().getElementNames())

    def test_intersection_with_cleaning_on_the_code_column(self):
        ui = self.run_op("Intersection", cols=(0,))
        self.assertEqual(self.names(), ["Sheet1", "Intersection", "Sheet2"])      # right after A
        self.assertEqual(self.sheet_rows("Intersection"), [
            ("Code", "Item", "Qty"), ("T-01", "Ceramic Tile", "10"), ("T-02", "Paint", "5"),
            ("101", "Bolt", "7")])                                      # 101 (number) matched "101" (text)
        self.assertIn("Found in both: 3", ui.confirms[0])
        self.assertIn("3 row(s) written to sheet 'Intersection'", ui.messages[-1])
        self.assertEqual(self.ctrl.getActiveSheet().Name, "Intersection")

    def test_source_sheets_are_untouched(self):
        before = (self.sheet_rows("Sheet1"), self.sheet_rows("Sheet2"))
        self.run_op("Union", cols=(0,))
        self.assertEqual((self.sheet_rows("Sheet1"), self.sheet_rows("Sheet2")), before)

    def test_all_operations_on_the_code_column(self):
        expect = {
            "Union": ["T-01", "T-02", "T-03", "101", "T-04"],
            "A minus B": ["T-03"],
            "B minus A": ["T-04"],
            "Symmetric difference": ["T-03", "T-04"],
        }
        for op, codes in expect.items():
            self.run_op(op, cols=(0,))
            sheet = {"Symmetric difference": "Sym diff"}.get(op, op)
            self.assertEqual([r[0] for r in self.sheet_rows(sheet)][1:], codes, op)

    def test_exact_matching_when_cleaning_is_off(self):
        self.run_op("Intersection", cols=(0,), trim=False, ignore_case=False, numeric=False, nonprint=False)
        self.assertEqual([r[0] for r in self.sheet_rows("Intersection")][1:], ["T-01"])

    def test_multiple_key_columns(self):
        self.run_op("Intersection", cols=(0, 1))     # code AND item, cleaned: T-01, T-02, 101/bolt
        self.assertEqual([r[0] for r in self.sheet_rows("Intersection")][1:], ["T-01", "T-02", "101"])

    def test_no_header_option(self):
        ui = self.run_op("Intersection", cols=(0,), header=False)
        self.assertIn("List A: 7 rows", ui.confirms[0])            # header row counts as data now

    def test_result_is_one_undo_step_and_decline_writes_nothing(self):
        self.run_op("Intersection", cols=(0,))
        self.assertIn("Intersection", self.names())
        self.doc.getUndoManager().undo()
        self.assertEqual(self.names(), ["Sheet1", "Sheet2"])
        self.run_op("Intersection", cols=(0,), confirm=False)
        self.assertEqual(self.names(), ["Sheet1", "Sheet2"])

    def test_number_formats_text_and_formulas_are_kept_sensibly(self):
        self.s1.getCellRangeByName("C2").setFormula("=1+1")             # formula in the Qty column
        self.s1.getCellRangeByName("B2").setString("007")               # text that looks numeric
        self.s1.getCellRangeByName("C3").setPropertyValue("NumberFormat", 2)    # 0.00
        self.run_op("Intersection", cols=(0,))
        out = self.doc.getSheets().getByName("Intersection")
        c2 = out.getCellRangeByName("C2")
        self.assertFalse(c2.getFormula().startswith("="), "formula must be frozen to a value")
        self.assertEqual(c2.getValue(), 2.0)
        b2 = out.getCellRangeByName("B2")
        self.assertEqual(b2.getString(), "007")
        self.assertEqual(b2.getType().value, "TEXT")
        self.assertEqual(out.getCellRangeByName("C3").getString(), "5.00")      # format copied

    def test_clean_output_option(self):
        self.s1.getCellRangeByName("B3").setString("  Paint  ")
        self.run_op("Intersection", cols=(0,), clean_out=True)
        self.assertEqual(self.doc.getSheets().getByName("Intersection").getCellRangeByName("B3").getString(), "Paint")
        self.run_op("Intersection", cols=(0,), clean_out=False)
        self.assertEqual(self.doc.getSheets().getByName("Intersection_2").getCellRangeByName("B3").getString(),
                         "  Paint  ")

    def test_unique_sheet_names(self):
        self.run_op("Union", cols=(0,))
        self.run_op("Union", cols=(0,))
        self.assertEqual(self.names(), ["Sheet1", "Union_2", "Union", "Sheet2"])

    def test_nothing_to_write(self):
        ui = self.run_op("A minus B", cols=(0, 1, 2), a=self.B, b=self.B)
        self.assertIn("Nothing to write", ui.messages[-1])
        self.assertEqual(self.names(), ["Sheet1", "Sheet2"])

    def test_bad_input_shows_message_and_user_can_cancel(self):
        ui = _FakeUi(form=[{"a": "Nope.A1:B2", "b": self.B}, {"a": self.A, "b": "$Sheet2.$A$1:$B$5"}, None])
        make_feature("calc_setops", ui).execute("Run", self.ctrl.getFrame(), ())
        self.assertIn("no sheet named", ui.messages[0])
        self.assertIn("same number of columns", ui.messages[1])
        self.assertEqual(self.names(), ["Sheet1", "Sheet2"])

    def test_ctrl_multi_selection_prefills_a_and_b(self):
        ranges = self.doc.createInstance("com.sun.star.sheet.SheetCellRanges")
        ranges.addRangeAddress(self.s1.getCellRangeByName("A1:C7").getRangeAddress(), False)
        ranges.addRangeAddress(self.s1.getCellRangeByName("E1:G5").getRangeAddress(), False)
        self.ctrl.select(ranges)            # NB: a multi-selection can only span ONE sheet
        ui = _FakeUi(form=None)                                         # user cancels the dialog
        make_feature("calc_setops", ui).execute("Run", self.ctrl.getFrame(), ())
        fields = {f.get("name"): f for f in ui.form_calls[0][1]}
        self.assertEqual(fields["a"]["value"], self.A)
        self.assertEqual(fields["b"]["value"], "$Sheet1.$E$1:$G$5")
        self.assertIs(fields["header"]["value"], True)                 # header guessed from row 1

    def test_single_selection_prefills_region_and_falls_back_without_picker(self):
        self.ctrl.select(self.s1.getCellRangeByName("B3"))
        ui = _FakeUi(form=None)                                         # fake pick_range -> False
        make_feature("calc_setops", ui).execute("Run", self.ctrl.getFrame(), ())
        fields = {f.get("name"): f for f in ui.form_calls[0][1]}
        self.assertEqual(fields["a"]["value"], "$Sheet1.$A$1:$C$6")     # region around B3 (row 7 is blank)
        self.assertEqual(fields["b"]["value"], "")


class CalcFillBlanks(Base):
    def setUp(self):
        self.doc = self.new_doc("scalc")
        self.ctrl = self.doc.getCurrentController()
        self.sh = self.doc.getSheets().getByIndex(0)

    def put(self, rows, top=0):
        for r, row in enumerate(rows):
            for c, v in enumerate(row):
                cell = self.sh.getCellByPosition(c, top + r)
                cell.setValue(v) if isinstance(v, (int, float)) else cell.setString(v)

    def column(self, c, n):
        return [self.sh.getCellByPosition(c, r).getString() for r in range(n)]

    def go(self, rng, **form):
        self.ctrl.select(self.sh.getCellRangeByName(rng))
        ui = _FakeUi(form=form)
        make_feature("calc_fill_blanks", ui).execute("Run", self.ctrl.getFrame(), ())
        return ui

    REPORT = [("Region", "Item", "Qty"), ("North", "Tiles", 1), ("", "Paint", 2), ("", "Putty", 3),
              ("South", "Tiles", 4), ("", "", 5)]

    def test_fill_down_report_and_leading_blank(self):
        self.put(self.REPORT)
        self.put([("", "x")], top=7)                          # blank row at the bottom, nothing above in col C
        ui = self.go("A1")                                   # lone cell -> region A1:C6 (row 7 is blank)
        self.assertEqual(self.column(0, 6), ["Region", "North", "North", "North", "South", "South"])
        self.assertEqual(self.column(1, 6), ["Item", "Tiles", "Paint", "Putty", "Tiles", "Tiles"])
        self.assertIn("Filled 4 cell(s)", ui.messages[-1])

    def test_numbers_stay_numbers_and_formats_untouched(self):
        self.put([(5.0,), ("",), ("",)])
        self.sh.getCellRangeByName("A2").setPropertyValue("CharWeight", 150.0)     # bold blank cell
        self.go("A1:A3")
        self.assertEqual(self.sh.getCellRangeByName("A3").getValue(), 5.0)
        self.assertEqual(self.sh.getCellRangeByName("A2").getPropertyValue("CharWeight"), 150.0)

    def test_only_the_selection_is_filled(self):
        self.put(self.REPORT)
        self.go("A1:A4")
        self.assertEqual(self.column(0, 6), ["Region", "North", "North", "North", "South", ""])

    def test_fill_right_for_merged_looking_headers(self):
        self.put([("2026", "", "", "2027", "", ""), ("a", "b", "c", "d", "e", "f")])
        self.go("A1:F2", direction="Right")
        self.assertEqual([self.sh.getCellByPosition(c, 0).getString() for c in range(6)],
                         ["2026", "2026", "2026", "2027", "2027", "2027"])
        self.assertEqual(self.sh.getCellRangeByName("F2").getString(), "f")

    def test_whitespace_only_cells_option(self):
        self.put([("A",), (" ",), ("B",)])
        ui = self.go("A1:A3")                                # default: a space is a value
        self.assertIn("No blank cells", ui.messages[-1])
        self.go("A1:A3", whitespace=True)
        self.assertEqual(self.column(0, 3), ["A", "A", "B"])

    def test_real_merged_cells_are_unmerged_and_filled(self):
        self.put([("Region", "Item"), ("North", "Tiles"), ("", "Paint"), ("", "Putty"), ("South", "Tiles")])
        self.sh.getCellRangeByName("A2:A4").merge(True)
        self.assertTrue(self.sh.getCellRangeByName("A2:A4").getIsMerged())
        ui = self.go("A1:B5")
        self.assertFalse(self.sh.getCellRangeByName("A2:A4").getIsMerged())
        self.assertEqual(self.column(0, 5), ["Region", "North", "North", "North", "South"])
        self.assertIn("Unmerged 1 merged area(s)", ui.messages[-1])

    def test_merged_cells_left_alone_when_unmerge_is_off(self):
        self.put([("Region",), ("North",), ("",), ("",)])
        self.sh.getCellRangeByName("A2:A4").merge(True)
        self.go("A1:A4", unmerge=False)
        self.assertTrue(self.sh.getCellRangeByName("A2:A4").getIsMerged())

    def test_single_undo_step_and_cancel(self):
        self.put(self.REPORT)
        self.go("A1:C6")
        self.assertEqual(self.column(0, 6)[2], "North")
        self.doc.getUndoManager().undo()
        self.assertEqual(self.column(0, 6), ["Region", "North", "", "", "South", ""])
        self.ctrl.select(self.sh.getCellRangeByName("A1:C6"))
        cancelled = _FakeUi(form=None)
        make_feature("calc_fill_blanks", cancelled).execute("Run", self.ctrl.getFrame(), ())
        self.assertEqual(self.column(0, 6), ["Region", "North", "", "", "South", ""])

    def test_multi_range_selection(self):
        self.put([("A", "B"), ("", ""), ("x", "y"), ("", "")])
        ranges = self.doc.createInstance("com.sun.star.sheet.SheetCellRanges")
        ranges.addRangeAddress(self.sh.getCellRangeByName("A1:A2").getRangeAddress(), False)
        ranges.addRangeAddress(self.sh.getCellRangeByName("B3:B4").getRangeAddress(), False)
        self.ctrl.select(ranges)
        make_feature("calc_fill_blanks", _FakeUi(form={})).execute("Run", self.ctrl.getFrame(), ())
        self.assertEqual(self.column(0, 4), ["A", "A", "x", ""])        # only A1:A2 filled
        self.assertEqual(self.column(1, 4), ["B", "", "y", "y"])        # only B3:B4 filled


class CalcRemoveBlanks(Base):
    DATA = [("H1", "H2", "H3"), ("a", "b", "c"), ("", "", ""), ("d", "", "f"), ("", "", ""), ("g", "h", "i")]

    def setUp(self):
        self.doc = self.new_doc("scalc")
        self.ctrl = self.doc.getCurrentController()
        self.sh = self.doc.getSheets().getByIndex(0)
        for r, row in enumerate(self.DATA):
            for c, v in enumerate(row):
                if v:
                    self.sh.getCellByPosition(c, r).setString(v)

    def grid(self, rows=9, cols=6):
        return [[self.sh.getCellByPosition(c, r).getString() for c in range(cols)] for r in range(rows)]

    def go(self, rng, confirm=True, **form):
        self.ctrl.select(self.sh.getCellRangeByName(rng))
        ui = _FakeUi(form=form, confirm=confirm)
        make_feature("calc_remove_blanks", ui).execute("Run", self.ctrl.getFrame(), ())
        return ui

    def test_blank_rows_inside_selection_shift_cells_up(self):
        self.sh.getCellRangeByName("A8").setString("tail")               # below the selection
        ui = self.go("A1:D6", mode="Blank rows")
        self.assertIn("Remove 2 blank row(s)", ui.confirms[0])
        rows = self.grid()
        self.assertEqual([r[:3] for r in rows[:4]],
                         [["H1", "H2", "H3"], ["a", "b", "c"], ["d", "", "f"], ["g", "h", "i"]])
        self.assertEqual(rows[5][0], "tail")                              # moved up by 2 (8 -> 6)
        self.assertIn("Removed 2 blank row(s)", ui.messages[-1])

    def test_whole_sheet_rows_option_deletes_outside_data_too(self):
        self.sh.getCellRangeByName("F3").setString("outside")             # same row as a blank row
        self.go("A1:D6", mode="Blank rows", whole=True)
        self.assertNotIn("outside", [c for row in self.grid() for c in row])

    def test_blank_columns(self):
        for r in range(6):                                                # column D blank, E has data
            self.sh.getCellByPosition(4, r).setString("e%d" % r)
        ui = self.go("A1:D6", mode="Blank columns")
        self.assertIn("1 blank column(s)", ui.confirms[0])
        self.assertEqual(self.grid(6)[1][3], "e1")                        # E moved left into D

    def test_blank_cells_shift_up_per_column(self):
        self.go("A1:C6", mode="Blank cells", shift="Up")
        rows = self.grid(6, 3)
        self.assertEqual([r[0] for r in rows], ["H1", "a", "d", "g", "", ""])
        self.assertEqual([r[1] for r in rows], ["H2", "b", "h", "", "", ""])
        self.assertEqual([r[2] for r in rows], ["H3", "c", "f", "i", "", ""])

    def test_blank_cells_shift_left_per_row(self):
        self.go("A1:C6", mode="Blank cells", shift="Left")
        self.assertEqual(self.grid(6, 3)[3], ["d", "f", ""])
        self.assertEqual(self.grid(6, 3)[1], ["a", "b", "c"])

    def test_whitespace_option(self):
        self.sh.getCellRangeByName("A3").setString("   ")
        ui = self.go("A1:C6", mode="Blank rows")
        self.assertIn("Remove 1 blank row(s)", ui.confirms[0])            # only row 5; row 3 has spaces
        ui = self.go("A1:C6", mode="Blank rows", whitespace=True)
        self.assertIn("Remove 2 blank row(s)", ui.confirms[0])

    def test_formula_returning_empty_text_is_not_blank(self):
        self.sh.getCellRangeByName("B3").setFormula('=""')
        ui = self.go("A1:C6", mode="Blank rows")
        self.assertIn("Remove 1 blank row(s)", ui.confirms[0])            # row 3 kept
        self.assertEqual(self.sh.getCellRangeByName("B3").getFormula(), '=""')

    def test_lone_cell_means_the_used_area(self):
        ui = self.go("B2", mode="Blank rows")
        self.assertIn("$Sheet1.$A$1:$C$6", ui.confirms[0])
        self.assertIn("Remove 2 blank row(s)", ui.confirms[0])

    def test_decline_nothing_found_and_multi_range(self):
        before = self.grid()
        self.go("A1:C6", mode="Blank rows", confirm=False)
        self.assertEqual(self.grid(), before)
        ui = self.go("A1:C2", mode="Blank rows")
        self.assertIn("No blank rows found", ui.messages[-1])
        ranges = self.doc.createInstance("com.sun.star.sheet.SheetCellRanges")
        for rng in ("A1:A2", "C1:C2"):
            ranges.addRangeAddress(self.sh.getCellRangeByName(rng).getRangeAddress(), False)
        self.ctrl.select(ranges)
        ui = _FakeUi(form={})
        make_feature("calc_remove_blanks", ui).execute("Run", self.ctrl.getFrame(), ())
        self.assertIn("single range", ui.messages[-1])

    def test_one_undo_step_for_each_mode(self):
        before = self.grid()
        for form in ({"mode": "Blank rows"}, {"mode": "Blank cells", "shift": "Up"},
                     {"mode": "Blank cells", "shift": "Left"}):
            self.go("A1:C6", **form)
            self.assertNotEqual(self.grid(), before, form)
            self.doc.getUndoManager().undo()
            self.assertEqual(self.grid(), before, form)


class CalcHighlight(Base):
    LIST = [("Item", "Qty"), ("A", 1), ("b", 2), ("a", 3), ("C", 4), ("B", 5)]
    A = "$Sheet1.$A$1:$B$6"
    B = "$Sheet2.$A$1:$B$3"
    NONE = -1
    PAL = None

    def setUp(self):
        sys.path.insert(0, os.path.join(ROOT, "src", "python"))
        from office_plugin.features.calc_highlight import marks
        self.M = marks
        self.doc = self.new_doc("scalc")
        self.doc.getSheets().insertNewByName("Sheet2", 1)
        self.ctrl = self.doc.getCurrentController()
        self.s1, self.s2 = self.doc.getSheets().getByIndex(0), self.doc.getSheets().getByIndex(1)
        for sheet, rows in ((self.s1, self.LIST), (self.s2, [("Item", "Qty"), ("a", 1), ("x", 2)])):
            for r, row in enumerate(rows):
                for c, v in enumerate(row):
                    cell = sheet.getCellByPosition(c, r)
                    cell.setValue(v) if isinstance(v, (int, float)) else cell.setString(v)

    def back(self, ref, sheet=None):
        return (sheet or self.s1).getCellRangeByName(ref).getPropertyValue("CellBackColor")

    def painted(self, column="A", rows=7, sheet=None):
        """1-based rows of `column` that carry a plugin colour."""
        pal = set(self.M.PALETTE.values())
        return [r for r in range(1, rows + 1) if self.back("%s%d" % (column, r), sheet) in pal]

    def run_highlight(self, mode, colour="Yellow", options=None, cols=(0,), a=None, b=None, select="A1:B6",
                      form_extra=None):
        self.ctrl.select(self.s1.getCellRangeByName(select))
        o = {"a": a or self.A, "b": b or self.B, "header": True}
        o.update(options or {})
        forms = [{"mode": mode, "colour": colour}, o]
        if mode != self.M.CELL_DIFF:
            forms.append({"cols": list(cols)})
        ui = _FakeUi(form=forms)
        make_feature("calc_highlight", ui).execute("Run", self.ctrl.getFrame(), ())
        return ui

    def test_duplicates_every_occurrence_whole_rows(self):
        ui = self.run_highlight(self.M.DUP_ALL)
        self.assertEqual(self.painted("A"), [2, 3, 4, 6])                 # A,b,a,B ; header and C not
        self.assertEqual(self.painted("B"), [2, 3, 4, 6])                 # whole row coloured
        self.assertEqual(self.back("A2"), self.M.PALETTE["Yellow"])
        self.assertIn("Highlighted 8 cell(s) in yellow", ui.messages[-1])
        self.assertEqual(self.back("A1"), self.NONE)

    def test_only_matching_columns_option(self):
        self.run_highlight(self.M.DUP_ALL, options={"key_only": True})
        self.assertEqual(self.painted("A"), [2, 3, 4, 6])
        self.assertEqual(self.painted("B"), [])

    def test_repeats_only_and_unique_and_auto_colour(self):
        self.run_highlight(self.M.DUP_REPEATS, colour="Auto (suggested for the mode)")
        self.assertEqual(self.painted("A"), [4, 6])                       # 2nd occurrences only
        self.assertEqual(self.back("A4"), self.M.PALETTE["Orange"])        # auto colour of that mode
        self.run_highlight(self.M.UNIQUE, colour="Blue")
        self.assertEqual(self.back("A5"), self.M.PALETTE["Blue"])          # "C"

    def test_rows_in_and_not_in_b_with_b_side(self):
        self.run_highlight(self.M.IN_BOTH, colour="Green", options={"both": True})
        self.assertEqual(self.painted("A"), [2, 4])                       # A and a are in B
        self.assertEqual(self.painted("A", sheet=self.s2), [2])           # "a" row in B
        self.run_highlight(self.M.ONLY_IN_A, colour="Red")
        self.assertEqual(self.back("A3"), self.M.PALETTE["Red"])          # b is not in B (and wins over old colour)

    def test_cell_by_cell_differences(self):
        for r, row in enumerate([("Item", "Qty"), ("A", 1), ("b", 9)]):
            for c, v in enumerate(row):
                cell = self.s2.getCellByPosition(c, r)
                cell.setValue(v) if isinstance(v, (int, float)) else cell.setString(v)
        ui = self.run_highlight(self.M.CELL_DIFF, colour="Red", select="A1:B3",
                                a="$Sheet1.$A$1:$B$3", b="$Sheet2.$A$1:$B$3", options={"both": True})
        self.assertEqual(self.back("B3"), self.M.PALETTE["Red"])           # 2 vs 9
        self.assertEqual(self.back("A3"), self.NONE)                       # "b" == "b"
        self.assertEqual(self.back("B2", sheet=self.s2), self.NONE)
        self.assertEqual(self.back("B3", sheet=self.s2), self.M.PALETTE["Red"])
        self.assertIn("Highlighted 2 cell(s)", ui.messages[-1])            # B3 in A plus B3 in B

    def test_size_mismatch_message_then_cancel(self):
        ui = _FakeUi(form=[{"mode": self.M.CELL_DIFF, "colour": "Red"},
                           {"a": "$Sheet1.$A$1:$B$3", "b": "$Sheet2.$A$1:$B$2"}, None])
        self.ctrl.select(self.s1.getCellRangeByName("A1:B3"))
        make_feature("calc_highlight", ui).execute("Run", self.ctrl.getFrame(), ())
        self.assertIn("same size", ui.messages[0])
        self.assertEqual(self.painted("A"), [])

    def test_cleaning_options_matter(self):
        self.run_highlight(self.M.DUP_ALL, options={"ignore_case": False})
        self.assertEqual(self.painted("A"), [], "A vs a differ when case matters")

    def test_nothing_to_highlight(self):
        ui = self.run_highlight(self.M.UNIQUE, cols=(1,), options={"numeric": True})   # Qty is unique per row
        self.assertEqual(self.painted("A"), [2, 3, 4, 5, 6])                  # all unique -> all rows painted
        ui = self.run_highlight(self.M.DUP_ALL, cols=(1,))                    # no duplicate quantities
        self.assertIn("Nothing to highlight", ui.messages[-1])

    def test_one_undo_step(self):
        self.run_highlight(self.M.DUP_ALL)
        self.assertTrue(self.painted("A"))
        self.doc.getUndoManager().undo()
        self.assertEqual(self.painted("A"), [])

    def test_clear_removes_only_plugin_colours(self):
        self.run_highlight(self.M.DUP_ALL)
        self.s1.getCellRangeByName("D2").setPropertyValue("CellBackColor", 0x123456)     # user's own colour
        self.s1.getCellRangeByName("B4").setPropertyValue("CellBackColor", 0x123457)     # user colour inside a painted row
        self.ctrl.select(self.s1.getCellRangeByName("A1:D7"))
        ui = _FakeUi()
        make_feature("calc_highlight", ui).execute("Clear", self.ctrl.getFrame(), ())
        self.assertEqual(self.painted("A"), [])
        self.assertEqual(self.painted("B"), [])
        self.assertEqual(self.back("D2"), 0x123456)
        self.assertEqual(self.back("B4"), 0x123457)
        self.assertIn("Cleared 7 highlighted cell(s)", ui.messages[-1])               # 8 painted minus the user-coloured B4

    def test_clear_lone_cell_uses_used_area_and_undo(self):
        self.run_highlight(self.M.DUP_ALL)
        self.ctrl.select(self.s1.getCellRangeByName("F20"))
        ui = _FakeUi()
        make_feature("calc_highlight", ui).execute("Clear", self.ctrl.getFrame(), ())
        self.assertEqual(self.painted("A"), [])
        self.doc.getUndoManager().undo()
        self.assertEqual(self.painted("A"), [2, 3, 4, 6])
        ui2 = _FakeUi()
        self.ctrl.select(self.s1.getCellRangeByName("A1:B6"))
        self.doc.getUndoManager().undo()                                    # remove the highlight itself
        make_feature("calc_highlight", ui2).execute("Clear", self.ctrl.getFrame(), ())
        self.assertIn("No plugin highlights found", ui2.messages[-1])


class _TextOwner(unohelper.Base):
    """Plain-text clipboard owner living in the test process (see CalcTranspose.release_clipboard)."""


def _make_text_transferable():
    from com.sun.star.datatransfer import DataFlavor, XTransferable

    class Text(unohelper.Base, XTransferable):
        def _flavor(self):
            f = DataFlavor()
            f.MimeType, f.HumanPresentableName = "text/plain;charset=utf-16", "Unicode Text"
            f.DataType = uno.getTypeByName("string")
            return f

        def getTransferData(self, flavor):
            return "plugin test"

        def getTransferDataFlavors(self):
            return (self._flavor(),)

        def isDataFlavorSupported(self, flavor):
            return flavor.MimeType.startswith("text/plain")

    return Text()


class CalcTranspose(Base):
    def release_clipboard(self):
        """Headless LibreOffice can crash when a document that owns the clipboard is closed (and a
        later document then pastes): hand the clipboard to this process before the document closes."""
        clip = Env.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.datatransfer.clipboard.SystemClipboard", Env.ctx)
        self._owner = _make_text_transferable()
        clip.setContents(self._owner, None)

    def setUp(self):
        self.doc = self.new_doc("scalc", hidden=False)       # sheet-insertion undo needs a real view
        self.addCleanup(self.release_clipboard)              # runs before the document is closed
        self.ctrl = self.doc.getCurrentController()
        self.sh = self.doc.getSheets().getByIndex(0)
        for r, (t, n) in enumerate([("a", 1), ("b", 2), ("c", None)]):
            self.sh.getCellByPosition(0, r).setString(t)
            if n is not None:
                self.sh.getCellByPosition(1, r).setValue(n)
        self.sh.getCellRangeByName("B3").setFormula("=B1+B2+$D$1")
        self.sh.getCellRangeByName("D1").setValue(100)
        self.sh.getCellRangeByName("B1").setPropertyValue("NumberFormat", 2)      # 0.00

    def go(self, select="A1:B3", dest="In place (replace the table)", target="", values=False,
           confirm=True, engine=None):
        self.ctrl.select(self.sh.getCellRangeByName(select))
        ui = _FakeUi(form={"dest": dest, "target": target or "$Sheet1.$F$5", "values": values}, confirm=confirm)
        feature = make_feature("calc_transpose", ui)
        if engine:
            feature.settings.set_option("calc_transpose", "engine", engine)
        feature.execute("Run", self.ctrl.getFrame(), ())
        return ui

    def formulas(self, rng):
        r = self.sh.getCellRangeByName(rng)
        return [list(row) for row in r.getFormulaArray()]

    def strings(self, rng):
        r = self.sh.getCellRangeByName(rng)
        return [list(row) for row in r.getDataArray()]

    SRC_BEFORE = [["a", "1"], ["b", "2"], ["c", "=B1+B2+$D$1"]]

    def test_in_place_swaps_shape_adjusts_formulas_and_leaves_no_stale_cells(self):
        ui = self.go()
        self.assertEqual(self.formulas("A1:C2"), [["a", "b", "c"], ["1", "2", "=A2+B2+$D$1"]])
        self.assertEqual(self.sh.getCellRangeByName("A2").getString(), "1.00")        # number format kept
        self.assertEqual(self.formulas("A3:C3"), [["", "", ""]])                       # old row 3 is gone
        self.assertEqual(self.sh.getCellRangeByName("D1").getValue(), 100.0)           # outside data untouched
        self.assertIn("Transposed $Sheet1.$A$1:$B$3 to $Sheet1.$A$1:$C$2", ui.messages[-1])

    def test_in_place_is_one_undo_step(self):
        self.go()
        self.doc.getUndoManager().undo()
        self.assertEqual(self.formulas("A1:B3"), self.SRC_BEFORE)
        self.assertEqual(self.sh.getCellRangeByName("B1").getString(), "1.00")

    def test_new_sheet_leaves_the_source_alone_and_undo_removes_it(self):
        self.go(dest="New sheet")
        names = list(self.doc.getSheets().getElementNames())
        self.assertEqual(names, ["Sheet1", "Transposed"])
        self.assertEqual(self.formulas("A1:B3"), self.SRC_BEFORE)
        out = self.doc.getSheets().getByName("Transposed")
        self.assertEqual([list(r) for r in out.getCellRangeByName("A1:C2").getFormulaArray()],
                         [["a", "b", "c"], ["1", "2", "=A2+B2+$D$1"]])
        self.assertEqual(self.ctrl.getActiveSheet().Name, "Transposed")
        self.doc.getUndoManager().undo()
        self.assertEqual(list(self.doc.getSheets().getElementNames()), ["Sheet1"])

    def test_at_cell(self):
        self.go(dest="At cell...", target="$Sheet1.$F$5")
        self.assertEqual(self.formulas("F5:H6"), [["a", "b", "c"], ["1", "2", "=F6+G6+$D$1"]])
        self.assertEqual(self.formulas("A1:B3"), self.SRC_BEFORE)

    def test_overwrite_warning_decline_and_accept(self):
        self.sh.getCellRangeByName("C1").setString("keep me")           # inside A1:C2, outside the table
        before = self.formulas("A1:D3")
        ui = self.go(confirm=False)
        self.assertIn("1 non-empty cell(s)", ui.confirms[0])
        self.assertEqual(self.formulas("A1:D3"), before)
        self.go(confirm=True)
        self.assertEqual(self.sh.getCellRangeByName("C1").getString(), "c")

    def test_no_warning_when_the_extra_cells_are_empty(self):
        ui = self.go()
        self.assertEqual(ui.confirms, [])

    def test_convert_formulas_to_values(self):
        ui = self.go(values=True)
        self.assertEqual(self.formulas("A1:C2")[1], ["1", "2", "103"])
        self.assertIn("1 formula(s) converted", ui.messages[-1])

    def test_lone_cell_means_the_region(self):
        self.go(select="A2")
        self.assertEqual(self.formulas("A1:C2")[0], ["a", "b", "c"])

    def test_row_to_column_and_column_to_row(self):
        for c, v in enumerate(("x", "y", "z")):
            self.sh.getCellByPosition(c, 9).setString(v)                  # A10:C10
        self.go(select="A10:C10")
        self.assertEqual([r[0] for r in self.formulas("A10:A12")], ["x", "y", "z"])
        self.go(select="A10:A12")                                          # and back
        self.assertEqual(self.formulas("A10:C10")[0], ["x", "y", "z"])

    def test_merged_cells_are_carried_over(self):
        self.sh.getCellRangeByName("A1:B1").merge(True)
        self.go()
        self.assertTrue(self.sh.getCellRangeByName("A1:A2").getIsMerged())

    def test_too_wide_overlap_and_multi_range_are_refused(self):
        self.sh.getCellRangeByName("A20000").setValue(1)                   # the used area really is 20000 rows tall
        ui = self.go(select="A1:A20000")                                   # -> 20000 columns
        self.assertIn("does not fit", ui.messages[-1])
        ui = self.go(dest="At cell...", target="$Sheet1.$B$1")
        self.assertIn("overlaps the table", ui.messages[-1])
        ranges = self.doc.createInstance("com.sun.star.sheet.SheetCellRanges")
        for rng in ("A1:A2", "C1:C2"):
            ranges.addRangeAddress(self.sh.getCellRangeByName(rng).getRangeAddress(), False)
        self.ctrl.select(ranges)
        ui = _FakeUi(form={})
        make_feature("calc_transpose", ui).execute("Run", self.ctrl.getFrame(), ())
        self.assertIn("single table", ui.messages[-1])
        self.assertEqual(self.formulas("A1:B3"), self.SRC_BEFORE)

    def test_cancel_changes_nothing(self):
        self.ctrl.select(self.sh.getCellRangeByName("A1:B3"))
        make_feature("calc_transpose", _FakeUi(form=None)).execute("Run", self.ctrl.getFrame(), ())
        self.assertEqual(self.formulas("A1:B3"), self.SRC_BEFORE)

    def test_cell_by_cell_fallback_engine(self):
        self.go(engine="cells")
        self.assertEqual(self.formulas("A1:C2"), [["a", "b", "c"], ["1", "2", "103"]])    # formula -> value
        self.assertEqual(self.sh.getCellRangeByName("A2").getString(), "1.00")           # number format kept
        self.assertEqual(self.formulas("A3:C3"), [["", "", ""]])
        self.doc.getUndoManager().undo()
        self.assertEqual(self.formulas("A1:B3"), self.SRC_BEFORE)


class CalcSequence(Base):
    def setUp(self):
        self.doc = self.new_doc("scalc")
        self.ctrl = self.doc.getCurrentController()
        self.sh = self.doc.getSheets().getByIndex(0)

    def go(self, select, confirm=True, forms=None, **form):
        self.ctrl.select(self.sh.getCellRangeByName(select))
        ui = _FakeUi(form=forms if forms is not None else form, confirm=confirm)
        make_feature("calc_sequence", ui).execute("Run", self.ctrl.getFrame(), ())
        return ui

    def col(self, ref_col="A", n=8):
        return [self.sh.getCellRangeByName("%s%d" % (ref_col, r)).getString() for r in range(1, n + 1)]

    def test_selection_size_is_the_default_count(self):
        ui = self.go("A1:A5")
        self.assertEqual(self.col()[:6], ["1", "2", "3", "4", "5", ""])
        self.assertIn("Filled 5 cell(s)", ui.messages[-1])

    def test_count_end_and_step(self):
        self.go("A1:A5", step="2", count="3")
        self.assertEqual(self.col()[:5], ["1", "3", "5", "", ""])                # count wins over selection size
        self.go("C1:C5", start="1", step="2", end="9", count="")
        self.assertEqual(self.col("C")[:6], ["1", "3", "5", "7", "9", ""])
        self.go("E1:E2", start="0", step="", end="1", count="5")                  # step computed = 0.25
        self.assertEqual([float(v) for v in self.col("E")[:5]], [0, 0.25, 0.5, 0.75, 1.0])

    def test_across_and_descending_and_decimals(self):
        self.go("A1:E1", start="10", step="5")                                    # wide selection -> Across
        self.assertEqual([self.sh.getCellByPosition(c, 0).getValue() for c in range(5)], [10, 15, 20, 25, 30])
        self.go("A3:A6", start="10", step="-3", count="4")
        self.assertEqual(self.col()[2:6], ["10", "7", "4", "1"])
        self.go("C3:C6", start="0", step="0.1", count="4")
        self.assertEqual([self.sh.getCellByPosition(2, r).getValue() for r in range(2, 6)], [0, 0.1, 0.2, 0.3])

    def test_values_are_real_numbers(self):
        self.go("A1:A3")
        self.assertEqual(self.sh.getCellRangeByName("A2").getType().value, "VALUE")
        self.assertEqual(self.sh.getCellRangeByName("A2").getValue(), 2.0)

    def test_overwriting_inside_the_selection_needs_no_confirm_beyond_does(self):
        self.sh.getCellRangeByName("A2").setString("old")
        ui = self.go("A1:A3")
        self.assertEqual(ui.confirms, [])                                         # your own selection
        self.sh.getCellRangeByName("A5").setString("keep")
        ui = self.go("A1:A3", count="6", confirm=False)                           # runs to A6, hits A5
        self.assertIn("1 non-empty cell(s) beyond the selection", ui.confirms[0])
        self.assertEqual(self.col()[4], "keep")
        self.go("A1:A3", count="6", confirm=True)
        self.assertEqual(self.col()[4], "5")

    def test_errors_are_shown_and_the_user_can_retry_or_cancel(self):
        ui = self.go("A1:A3", forms=[{"step": "0"}, {"start": "abc"}, {"start": "1", "step": "1", "end": "-5"}, None])
        self.assertEqual(len(ui.messages), 3)
        self.assertIn("Step cannot be 0", ui.messages[0])
        self.assertIn("plain number", ui.messages[1])
        self.assertIn("away from End", ui.messages[2])
        self.assertEqual(self.col()[:3], ["", "", ""])

    def test_does_not_fit_on_the_sheet_and_multi_range(self):
        ui = self.go("A1048570", count="20")
        self.assertIn("does not fit", ui.messages[-1])
        ranges = self.doc.createInstance("com.sun.star.sheet.SheetCellRanges")
        for rng in ("A1:A2", "C1:C2"):
            ranges.addRangeAddress(self.sh.getCellRangeByName(rng).getRangeAddress(), False)
        self.ctrl.select(ranges)
        ui = _FakeUi(form={})
        make_feature("calc_sequence", ui).execute("Run", self.ctrl.getFrame(), ())
        self.assertIn("Select the cells to number", ui.messages[-1])

    def test_one_undo_step_and_result_is_selected(self):
        self.go("A1:A4")
        self.assertEqual(self.ctrl.getSelection().getRangeAddress().EndRow, 3)
        self.doc.getUndoManager().undo()
        self.assertEqual(self.col()[:4], ["", "", "", ""])


class CalcMissingSeq(Base):
    def setUp(self):
        self.doc = self.new_doc("scalc")
        self.ctrl = self.doc.getCurrentController()
        self.sh = self.doc.getSheets().getByIndex(0)

    def put(self, rows, col=0):
        for r, row in enumerate(rows):
            for c, v in enumerate(row):
                cell = self.sh.getCellByPosition(col + c, r)
                if v == "":
                    continue
                cell.setValue(v) if isinstance(v, (int, float)) else cell.setString(v)

    def column(self, col, n):
        return [self.sh.getCellByPosition(col, r).getString() for r in range(n)]

    def back(self, ref):
        return self.sh.getCellRangeByName(ref).getPropertyValue("CellBackColor")

    NUMBERS = [("No", "Item"), (1, "a"), (2, "b"), (5, "c"), (6, "d"), (9, "e")]

    def go(self, select="A1:B6", confirm=True, forms=None, **form):
        self.ctrl.select(self.sh.getCellRangeByName(select))
        ui = _FakeUi(form=forms if forms is not None else form, confirm=confirm)
        make_feature("calc_missing_seq", ui).execute("Run", self.ctrl.getFrame(), ())
        return ui

    def test_inserts_placeholder_rows_with_numbers_and_colour(self):
        self.put(self.NUMBERS)
        self.sh.getCellRangeByName("A10").setString("tail")                       # below the table
        ui = self.go()
        self.assertIn("4 missing number(s) in 2 gap(s):\n3-4, 7-8", ui.confirms[0])
        self.assertEqual(self.column(0, 10), ["No", "1", "2", "3", "4", "5", "6", "7", "8", "9"])
        self.assertEqual(self.column(1, 10), ["Item", "a", "b", "", "", "c", "d", "", "", "e"])
        yellow = 0xFFEB84
        self.assertEqual([self.back("A%d" % r) for r in (3, 4, 5, 6)], [-1, yellow, yellow, -1])
        self.assertEqual(self.back("B8"), yellow)
        self.assertEqual(self.sh.getCellRangeByName("A14").getString(), "tail")   # shifted down by 4
        self.assertIn("Inserted 4 placeholder row(s) for 2 gap(s)", ui.messages[-1])

    def test_whole_rows_option_moves_data_beside_the_table_too(self):
        self.put(self.NUMBERS)
        self.sh.getCellRangeByName("D4").setString("beside")                      # same row as the number 5
        self.go(whole=True)
        self.assertEqual(self.sh.getCellRangeByName("D6").getString(), "beside")
        for ref in ("D4", "D5"):
            self.assertEqual(self.sh.getCellRangeByName(ref).getString(), "")

    def test_selection_columns_only_leaves_data_beside_in_place(self):
        self.put(self.NUMBERS)
        self.sh.getCellRangeByName("D4").setString("beside")
        self.go()
        self.assertEqual(self.sh.getCellRangeByName("D4").getString(), "beside")

    def test_text_numbers_keep_prefix_and_padding(self):
        self.put([("Doc",), ("INV-0001",), ("INV-0002",), ("INV-0005",)])
        self.go(select="A1:A4")
        self.assertEqual(self.column(0, 6), ["Doc", "INV-0001", "INV-0002", "INV-0003", "INV-0004", "INV-0005"])

    def test_descending_and_step_and_no_colour(self):
        self.put([("No",), (10,), (9,), (6,), (5,)])
        self.go(select="A1:A5", direction="Descending", colour="No colour")
        self.assertEqual(self.column(0, 7), ["No", "10", "9", "8", "7", "6", "5"])
        self.assertEqual(self.back("A4"), -1)
        self.put([("No",), (5,), (10,), (20,)], col=3)
        self.go(select="D1:D4", step="5")
        self.assertEqual(self.column(3, 5), ["No", "5", "10", "15", "20"])

    def test_no_gap_decline_and_duplicate_report(self):
        self.put([("No",), (1,), (2,), (2,), (3,)])
        ui = self.go(select="A1:A5")
        self.assertIn("No missing numbers", ui.messages[-1])
        self.assertIn("1 duplicate(s)", ui.messages[-1])
        self.put([("No",), (1,), (4,)], col=2)
        before = self.column(2, 4)
        self.go(select="C1:C3", confirm=False)
        self.assertEqual(self.column(2, 4), before)

    def test_one_undo_step(self):
        self.put(self.NUMBERS)
        self.go()
        self.doc.getUndoManager().undo()
        self.assertEqual(self.column(0, 7), ["No", "1", "2", "5", "6", "9", ""])
        self.assertEqual(self.back("A4"), -1)

    def test_limits_and_bad_step(self):
        self.put([("No",), (1,), (200000,)])
        ui = self.go(select="A1:A3")
        self.assertIn("too many to insert", ui.messages[-1])
        self.put([("No",), (1,), (20000,)], col=2)
        ui = self.go(select="C1:C3", confirm=False)
        self.assertIn("That is a LOT of rows", ui.confirms[0])
        ui = self.go(select="A1:A3", forms=[{"step": "abc"}, {"step": "1.5"}, None])
        self.assertEqual(len(ui.messages), 2)

    def test_multi_range_is_refused(self):
        self.put(self.NUMBERS)
        ranges = self.doc.createInstance("com.sun.star.sheet.SheetCellRanges")
        for rng in ("A1:A2", "C1:C2"):
            ranges.addRangeAddress(self.sh.getCellRangeByName(rng).getRangeAddress(), False)
        self.ctrl.select(ranges)
        ui = _FakeUi(form={})
        make_feature("calc_missing_seq", ui).execute("Run", self.ctrl.getFrame(), ())
        self.assertIn("Select the table", ui.messages[-1])


def render_colours(doc, sheet, refs):
    """What LibreOffice actually draws: export the sheet to PNG on one big borderless page and
    sample the centre of each cell -> {ref: (r, g, b)}. Needs a visible (non-hidden) document."""
    import pngsample
    page_styles = doc.getStyleFamilies().getByName("PageStyles")
    style = page_styles.getByName(sheet.getPropertyValue("PageStyle"))
    for key, value in (("IsLandscape", True), ("Width", 100000), ("Height", 40000), ("LeftMargin", 0),
                       ("RightMargin", 0), ("TopMargin", 0), ("BottomMargin", 0), ("HeaderIsOn", False),
                       ("FooterIsOn", False), ("PrintGrid", False), ("PrintHeaders", False)):
        style.setPropertyValue(key, value)
    path = os.path.join(tempfile.gettempdir(), "opl_render.png")
    if os.path.exists(path):
        os.remove(path)
    doc.storeToURL(uno.systemPathToFileUrl(path), (pv("FilterName", "calc_png_Export"),))
    png = pngsample.read_png(path)
    kx, ky = png.width / 100000.0, png.height / 40000.0
    out = {}
    for ref in refs:
        cell = sheet.getCellRangeByName(ref)
        pos, size = cell.Position, cell.Size
        out[ref] = png.pixel(int((pos.X + size.Width / 2) * kx), int((pos.Y + size.Height / 2) * ky))
    return out


class CalcFormatServices(Base):
    BAR, DONE, MILE = (0x9D, 0xC3, 0xFF), (0x1F, 0x4E, 0x9E), (0x6A, 0x3D, 0x9A)
    WHITE = (255, 255, 255)

    def setUp(self):
        sys.path.insert(0, os.path.join(ROOT, "src", "python"))
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from office_plugin.services import calc_format as F
        self.F = F
        self.doc = self.new_doc("scalc", hidden=False)
        self.sh = self.doc.getSheets().getByIndex(0)
        F.ensure_cell_style(self.doc, "Gantt Bar", 0x9DC3FF)
        F.ensure_cell_style(self.doc, "Gantt Done", 0x1F4E9E)
        F.ensure_cell_style(self.doc, "Gantt Milestone", 0x6A3D9A)
        base = 46300
        for r, (s, d, p) in enumerate([(base, 3, 0.5), (base + 2, 0, 0), (base + 1, 4, 100)], start=1):
            for c, v in ((2, s), (3, d), (4, p)):
                self.sh.getCellByPosition(c, r).setValue(v)
        for i in range(6):
            self.sh.getCellByPosition(6 + i, 0).setValue(base + i)
        self.rng = self.sh.getCellRangeByName("G2:L4")

    DONE_DAYS = "ROUND($D2*IF($E2>1;$E2/100;$E2);0)"

    def entries(self):
        d = self.DONE_DAYS
        return [("AND(ISNUMBER($C2);$D2=0;G$1<=$C2;G$1>=$C2)", "Gantt Milestone"),
                ("AND(ISNUMBER($C2);$D2>0;%s>0;G$1<=$C2+%s-1;G$1>=$C2)" % (d, d), "Gantt Done"),
                ("AND(ISNUMBER($C2);$D2>0;G$1<=$C2+$D2-1;G$1>=$C2)", "Gantt Bar")]

    def grid(self):
        refs = ["%s%d" % (col, r) for r in (2, 3, 4) for col in "GHIJKL"]
        px = render_colours(self.doc, self.sh, refs)
        return [[px["%s%d" % (col, r)] for col in "GHIJKL"] for r in (2, 3, 4)]

    def test_conditional_formats_render_bars_progress_and_milestones(self):
        self.F.set_conditional_formats(self.rng, self.entries())
        W, B, D, M = self.WHITE, self.BAR, self.DONE, self.MILE
        self.assertEqual(self.grid(), [
            [D, D, B, W, W, W],          # T1: 3 days, 50% -> round(1.5) = 2 done days, then the light bar
            [W, W, M, W, W, W],          # T2: milestone on day 3
            [W, D, D, D, D, W]])         # T3: 4 days at "100" (percent points) -> all done

    def test_stored_entries_are_readable_and_replaceable(self):
        self.F.set_conditional_formats(self.rng, self.entries())
        self.assertEqual(self.F.conditional_entries(self.rng), self.entries())
        self.F.set_conditional_formats(self.rng, self.entries()[2:])                # replace, not append
        self.assertEqual(len(self.F.conditional_entries(self.rng)), 1)

    def test_clear_removes_the_bars_and_first_set_is_undoable(self):
        self.F.set_conditional_formats(self.rng, self.entries())
        self.doc.getUndoManager().undo()                       # the first set IS one undo step
        self.assertEqual(self.F.conditional_entries(self.rng), [])
        # NOTE (LibreOffice): redo, and clearing/replacing conditional formats through the API, are not
        # reliably recorded, so tools must not promise that Ctrl+Z / Ctrl+Y bring formatting back.
        self.F.set_conditional_formats(self.rng, self.entries())
        self.F.clear_conditional_formats(self.rng)
        self.assertEqual(self.F.conditional_entries(self.rng), [])
        self.assertTrue(all(c == self.WHITE for row in self.grid() for c in row))      # nothing drawn any more

    def test_style_is_idempotent_and_recoloured(self):
        self.F.ensure_cell_style(self.doc, "Gantt Bar", 0x112233)
        styles = self.doc.getStyleFamilies().getByName("CellStyles")
        self.assertEqual(styles.getByName("Gantt Bar").getPropertyValue("CellBackColor"), 0x112233)
        self.assertEqual([n for n in styles.getElementNames()].count("Gantt Bar"), 1)

    def test_date_formats(self):
        cell = self.sh.getCellRangeByName("A1")
        cell.setValue(46300)
        for pattern, shown in (("DD-MMM", "05-Oct"), ("MMM-YY", "Oct-26")):
            cell.setPropertyValue("NumberFormat", self.F.ensure_date_format(self.doc, pattern))
            self.assertEqual(cell.getString(), shown)
        self.assertEqual(self.F.ensure_date_format(self.doc, "DD-MMM"), self.F.ensure_date_format(self.doc, "DD-MMM"))


class CalcGantt(Base):
    BASE = 46300                                                    # Monday 5 Oct 2026
    HEAD = ("ID", "Task", "Start", "Duration", "% Complete", "Predecessors", "Lag")
    TASKS = [(1, "Excavation", BASE, 3, 100, "", ""),
             (2, "Foundation", "", 4, 50, "1", ""),
             (3, "Walls", "", 5, 0, "2", ""),
             (4, "Handover", "", 0, 0, "3", ""),
             (5, "Independent", BASE + 1, 2, 0, "", "")]
    NAMES = {(0x9D, 0xC3, 0xFF): "bar", (0x1F, 0x4E, 0x9E): "done", (0x6A, 0x3D, 0x9A): "milestone",
             (255, 255, 255): ""}

    def setUp(self):
        sys.path.insert(0, os.path.join(ROOT, "src", "python"))
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from office_plugin.features.calc_gantt import gantt
        from office_plugin.services import calc_format, doc_store
        self.G, self.F, self.store = gantt, calc_format, doc_store
        self.doc = self.new_doc("scalc", hidden=False)
        self.ctrl = self.doc.getCurrentController()
        self.sh = self.doc.getSheets().getByIndex(0)

    # ---- helpers ---------------------------------------------------------------------
    def put_table(self, head=None, tasks=None, top=0):
        for c, h in enumerate(head or self.HEAD):
            self.sh.getCellByPosition(c, top).setString(h)
        for r, row in enumerate(tasks if tasks is not None else self.TASKS, start=top + 1):
            for c, v in enumerate(row):
                if v == "":
                    continue
                cell = self.sh.getCellByPosition(c, r)
                cell.setValue(v) if isinstance(v, (int, float)) else cell.setString(v)

    def go(self, select="A1", confirm=True, forms=None, **form):
        self.ctrl.select(self.sh.getCellRangeByName(select))
        ui = _FakeUi(form=forms if forms is not None else form, confirm=confirm)
        make_feature("calc_gantt", ui).execute("Run", self.ctrl.getFrame(), ())
        return ui

    def remove(self, confirm=True):
        ui = _FakeUi(confirm=confirm)
        make_feature("calc_gantt", ui).execute("Remove", self.ctrl.getFrame(), ())
        return ui

    def value(self, ref):
        return self.sh.getCellRangeByName(ref).getValue()

    def header(self, row=0, first_col=8):
        out, c = [], first_col
        while self.sh.getCellByPosition(c, row).getType().value == "VALUE":
            out.append(int(self.sh.getCellByPosition(c, row).getValue()))
            c += 1
        return out

    def period_end(self, start, scale):
        import calendar
        if scale == "Day":
            return start
        if scale == "Week":
            return start + 6
        d = date.fromordinal(date(1899, 12, 30).toordinal() + start)
        return start + (calendar.monthrange(d.year, d.month)[1] - d.day)

    def rendered(self, rows, first_col=8, scale="Day"):
        """{(row, period index): name} actually drawn by LibreOffice."""
        starts = self.header(0, first_col)
        refs = ["%s%d" % (self.G.col_name(first_col + i), r + 1) for r in rows for i in range(len(starts))]
        px = render_colours(self.doc, self.sh, refs)
        return starts, {(r, i): self.NAMES.get(px["%s%d" % (self.G.col_name(first_col + i), r + 1)], str(px["%s%d" % (self.G.col_name(first_col + i), r + 1)]))
                        for r in rows for i in range(len(starts))}

    def assert_matches_model(self, scale="Day", mode="duration", first_col=8, n_tasks=5):
        starts, drawn = self.rendered(range(1, n_tasks + 1), first_col, scale)
        for r in range(1, n_tasks + 1):
            start = self.sh.getCellByPosition(2, r).getValue() if self.sh.getCellByPosition(2, r).getType().value == "VALUE" else None
            if mode == "duration":
                dur_cell = self.sh.getCellByPosition(3, r)
                days = int(dur_cell.getValue()) if dur_cell.getType().value == "VALUE" else None
            else:
                end_cell = self.sh.getCellByPosition(3, r)
                days = (int(end_cell.getValue()) - int(start) + 1) if end_cell.getType().value == "VALUE" else 0
            pct = self.sh.getCellByPosition(4, r).getValue() if self.sh.getCellByPosition(4, r).getType().value == "VALUE" else 0
            for i, p in enumerate(starts):
                want = "" if start is None or days is None else self.G.expected_cell(int(start), days, pct, p, self.period_end(p, scale))
                self.assertEqual(drawn[(r, i)], want, "row %d, period %d (%s)" % (r + 1, i, p))

    # ---- creation ---------------------------------------------------------------------
    def test_creates_timeline_schedules_dependencies_and_draws_the_bars(self):
        self.put_table()
        ui = self.go()
        # dependencies written back into the Start column
        self.assertEqual([self.value("C%d" % r) for r in (2, 3, 4, 5, 6)],
                         [self.BASE, self.BASE + 3, self.BASE + 7, self.BASE + 12, self.BASE + 1])
        starts = self.header()
        self.assertEqual(starts, list(range(self.BASE, self.BASE + 14)))        # 13 days + one extra period
        self.assertEqual(self.sh.getCellRangeByName("I1").getString(), "05-Oct")
        self.assertEqual(self.sh.getCellRangeByName("H1").getString(), "")      # the spare column
        self.assertEqual(self.sh.getCellRangeByName("I1").getPropertyValue("RotateAngle"), 9000)
        self.assertIn("Rescheduled 3 task(s)", ui.messages[-1])
        self.assert_matches_model()                                              # what LibreOffice DRAWS

    def test_bars_follow_edited_dates_live(self):
        self.put_table()
        self.go()
        self.sh.getCellRangeByName("D6").setValue(4)                             # lengthen the independent task
        self.sh.getCellRangeByName("E6").setValue(50)
        self.assert_matches_model()                                              # no re-run needed

    def test_rerun_recalculates_and_replaces_the_old_timeline(self):
        self.put_table()
        self.go()
        self.sh.getCellRangeByName("D2").setValue(6)                             # Excavation 3 -> 6 days
        ui = self.go(select="Z50")                                               # any cell: saved table is used
        self.assertEqual([self.value("C%d" % r) for r in (3, 4, 5)],
                         [self.BASE + 6, self.BASE + 10, self.BASE + 15])
        self.assertEqual(self.header()[-1], self.BASE + 16)                      # timeline grew by 3 days
        self.assertEqual(ui.form_calls[0][1][2]["value"], "C - Start")           # dialog prefilled from the saved mapping
        self.assert_matches_model()
        self.go(select="Z50")                                                    # idempotent
        self.assertEqual([self.value("C%d" % r) for r in (3, 4, 5)], [self.BASE + 6, self.BASE + 10, self.BASE + 15])

    def test_week_and_month_scales(self):
        self.put_table(tasks=[(1, "A", self.BASE + 2, 20, 50, "", ""), (2, "B", self.BASE + 25, 0, 0, "", "")])
        self.go(scale="Week")
        starts = self.header()
        self.assertEqual(starts[0], self.BASE)                                   # Monday
        self.assertTrue(all(b - a == 7 for a, b in zip(starts, starts[1:])))
        self.assert_matches_model("Week", n_tasks=2)
        self.go(scale="Month")
        self.assertEqual(self.sh.getCellRangeByName("I1").getString(), "Oct-26")
        self.assertEqual(len(self.header()), 2)                                  # Oct (the tasks) + one extra period (Nov)
        self.assert_matches_model("Month", n_tasks=2)

    def test_end_date_mode_with_blank_end_as_milestone(self):
        head = ("ID", "Task", "Start", "End", "% Complete", "Predecessors", "Lag")
        rows = [(1, "A", self.BASE, self.BASE + 2, 50, "", ""), (2, "B", self.BASE + 10, self.BASE + 13, 0, "1", ""),
                (3, "M", self.BASE + 4, "", 0, "", "")]
        self.put_table(head, rows)
        self.go(mode="Start + End date")
        self.assertEqual(self.value("C3"), self.BASE + 3)                        # starts the day after A ends
        self.assertEqual(self.value("D3"), self.BASE + 6)                       # its 4-day length is kept: end = start + 3
        self.assertEqual(self.sh.getCellRangeByName("D4").getType().value, "EMPTY")   # milestone keeps a blank end
        self.assert_matches_model(mode="end", n_tasks=3)

    # ---- problems -------------------------------------------------------------------------
    def test_problems_change_nothing(self):
        cases = {
            "circular": [(1, "A", self.BASE, 1, 0, "2", ""), (2, "B", "", 1, 0, "1", "")],
            "unknown predecessor": [(1, "A", self.BASE, 1, 0, "9", "")],
            "start is text": [(1, "A", "soon", 1, 0, "", "")],
            "negative duration": [(1, "A", self.BASE, -2, 0, "", "")],
        }
        for label, rows in cases.items():
            with self.subTest(label):
                self.sh.clearContents(1023)
                self.put_table(tasks=rows)
                ui = self.go()
                self.assertIn("Nothing was changed", ui.messages[-1], label)
                self.assertEqual(self.sh.getCellRangeByName("I1").getType().value, "EMPTY")
        self.sh.clearContents(1023)
        self.put_table(tasks=[(1, "A", self.BASE, 1, 0, "", ""), (2, "B", "", 1, 0, "1", "")])
        before = self.value("C3")
        ui = self.go(forms=[{"start": "(none)"}, {"second": "D - Duration", "preds": "(none)"}, None])
        self.assertIn("Choose the Start date column", ui.messages[0])           # first answer rejected, user retries
        self.assertEqual(self.value("C3"), before)

    def test_validation_messages(self):
        self.put_table()
        ui = self.go(forms=[{"id": "(none)"}, {"preds": "(none)", "start": "D - Duration"}, None])
        self.assertIn("Predecessors need a Task ID", ui.messages[0])
        self.assertIn("one purpose only", ui.messages[1])

    def test_refusals(self):
        ranges = self.doc.createInstance("com.sun.star.sheet.SheetCellRanges")
        for rng in ("A1:A2", "C1:C2"):
            ranges.addRangeAddress(self.sh.getCellRangeByName(rng).getRangeAddress(), False)
        self.ctrl.select(ranges)
        ui = _FakeUi(form={})
        make_feature("calc_gantt", ui).execute("Run", self.ctrl.getFrame(), ())
        self.assertIn("Select the task table", ui.messages[-1])
        self.put_table(tasks=[])
        ui = self.go(select="A1:G1")
        self.assertIn("header row and at least one task", ui.messages[-1])

    def test_too_many_columns_asks_for_a_coarser_scale(self):
        self.put_table(tasks=[(1, "A", self.BASE, 3, 0, "", ""), (2, "B", self.BASE + 600, 3, 0, "", "")])
        ui = self.go(scale="Day")
        self.assertIn("coarser scale", ui.messages[-1])
        self.assertEqual(self.sh.getCellRangeByName("I1").getType().value, "EMPTY")

    def test_occupied_area_needs_confirmation(self):
        self.put_table()
        self.sh.getCellRangeByName("J3").setString("mine")
        ui = self.go(confirm=False)
        self.assertIn("not empty", ui.confirms[0])
        self.assertEqual(self.sh.getCellRangeByName("J3").getString(), "mine")
        self.assertEqual(self.sh.getCellRangeByName("C3").getType().value, "EMPTY")      # nothing written either

    # ---- undo / remove / persistence ---------------------------------------------------------
    def test_undo_reverts_the_cells(self):
        self.put_table()
        self.go()
        self.doc.getUndoManager().undo()
        self.assertEqual(self.sh.getCellRangeByName("C3").getType().value, "EMPTY")      # computed dates gone
        self.assertEqual(self.sh.getCellRangeByName("I1").getType().value, "EMPTY")      # header dates gone
        self.assertEqual(self.sh.getCellRangeByName("A2").getValue(), 1.0)               # the table is intact

    def test_remove_clears_the_timeline_but_keeps_the_table(self):
        self.put_table()
        self.go()
        ui = self.remove(confirm=False)
        self.assertIn("columns I to", ui.confirms[0])
        self.assertEqual(self.sh.getCellRangeByName("I1").getType().value, "VALUE")      # declined: still there
        ui = self.remove()
        self.assertIn("removed", ui.messages[-1])
        self.assertEqual(self.sh.getCellRangeByName("I1").getType().value, "EMPTY")
        self.assertEqual(self.value("C3"), self.BASE + 3)                                # computed dates stay
        self.assertEqual(self.F.conditional_entries(self.sh.getCellRangeByName("I2:V6")), [])
        self.assertIsNone(self.store.get(self.doc, "Gantt.Sheet1"))
        self.assertIn("No Gantt chart", self.remove().messages[-1])

    def test_survives_save_and_reopen_as_ods_and_xlsx(self):
        self.put_table()
        self.go()
        for ext, flt in (("ods", "calc8"), ("xlsx", "Calc MS Excel 2007 XML")):
            path = os.path.join(tempfile.gettempdir(), "opl_gantt." + ext)
            if os.path.exists(path):
                os.remove(path)
            self.doc.storeToURL(uno.systemPathToFileUrl(path), (pv("FilterName", flt),))
            again = Env.desktop.loadComponentFromURL(uno.systemPathToFileUrl(path), "_blank", 0, (pv("Hidden", True),))
            try:
                sh = again.getSheets().getByIndex(0)
                entries = self.F.conditional_entries(sh.getCellRangeByName("I2:V6"))
                self.assertEqual(len(entries), 3, ext)
                self.assertEqual([f for f, _ in entries], [f for f, _ in self.F.conditional_entries(self.sh.getCellRangeByName("I2:V6"))], ext)
                if ext == "ods":                         # xlsx re-creates the styles under generated names
                    self.assertEqual([style for _, style in entries], ["Gantt Milestone", "Gantt Done", "Gantt Bar"])
                self.assertEqual(self.store.get(again, "Gantt.Sheet1")["mode"], "duration", ext)
            finally:
                again.close(True)


class CalcDecimals(Base):
    WRAP = "Wrap in ROUND (formulas stay formulas)"
    LEAVE = "Leave formulas as they are"
    REMOVE = "Remove the ROUND wrapper"

    def setUp(self):
        self.doc = self.new_doc("scalc")
        self.ctrl = self.doc.getCurrentController()
        self.sh = self.doc.getSheets().getByIndex(0)

    def fmt(self, ref, code):
        nf = self.doc.getNumberFormats()
        loc = uno.createUnoStruct("com.sun.star.lang.Locale", "en", "US", "")
        key = nf.queryKey(code, loc, False)
        self.sh.getCellRangeByName(ref).setPropertyValue("NumberFormat", nf.addNew(code, loc) if key == -1 else key)

    def fill(self):
        cell = lambda ref: self.sh.getCellRangeByName(ref)
        cell("A1").setValue(3.14159)
        cell("A2").setValue(2.675)
        cell("A3").setString("text")
        cell("A4").setFormula("=A1*A2")
        cell("A5").setFormula("=ROUND(A1;4)")
        cell("A6").setValue(46300)
        self.fmt("A6", "DD-MMM-YY")
        cell("A7").setFormula('=TEXT(A1;"0")')
        cell("A8").setValue(5)
        cell("A9").setFormula("=1/0")

    def go(self, select="A1:A9", **form):
        self.ctrl.select(self.sh.getCellRangeByName(select))
        ui = _FakeUi(form=form)
        make_feature("calc_decimals", ui).execute("Run", self.ctrl.getFrame(), ())
        return ui

    def f(self, ref):
        return self.sh.getCellRangeByName(ref).getFormula()

    def s(self, ref):
        return self.sh.getCellRangeByName(ref).getString()

    def test_values_rounded_formulas_wrapped_everything_displays_two_decimals(self):
        self.fill()
        ui = self.go()
        self.assertEqual(self.sh.getCellRangeByName("A1").getValue(), 3.14)
        self.assertEqual(self.sh.getCellRangeByName("A2").getValue(), 2.68)               # 2.675 rounds half up like Calc
        self.assertEqual((self.s("A1"), self.s("A2"), self.s("A8")), ("3.14", "2.68", "5.00"))
        self.assertEqual(self.f("A4"), "=ROUND(A1*A2;2)")                                  # formula stays a formula
        self.assertEqual(self.f("A5"), "=ROUND(A1;2)")                                     # existing wrapper replaced
        self.assertEqual(self.s("A4"), "8.42")                                             # 3.14 * 2.68 = 8.4152
        self.assertEqual(self.s("A3"), "text")
        self.assertEqual(self.s("A6"), "05-Oct-26")                                        # date untouched
        self.assertEqual(self.f("A7"), '=TEXT(A1;"0")')                                    # text-valued formula untouched
        self.assertEqual(self.f("A9"), "=1/0")                                             # error untouched
        self.assertIn("2 value(s) rounded", ui.messages[-1])
        self.assertIn("1 formula(s) wrapped", ui.messages[-1])
        self.assertIn("1 existing ROUND wrapper(s) updated", ui.messages[-1])
        self.assertIn("Skipped:", ui.messages[-1])

    def test_idempotent_never_nests(self):
        self.fill()
        self.go()
        ui = self.go()
        self.assertEqual(self.f("A4"), "=ROUND(A1*A2;2)")
        self.assertIn("Nothing needed changing", ui.messages[-1])
        self.go(places="3")
        self.assertEqual(self.f("A4"), "=ROUND(A1*A2;3)")
        self.assertEqual(self.s("A1"), "3.140")

    def test_modes_up_and_down(self):
        self.sh.getCellRangeByName("A1").setValue(2.341)
        self.sh.getCellRangeByName("A2").setValue(2.349)
        self.sh.getCellRangeByName("A3").setFormula("=A1")
        self.go(select="A1:A3", mode="Round up (away from zero)")
        self.assertEqual(self.sh.getCellRangeByName("A1").getValue(), 2.35)
        self.assertEqual(self.f("A3"), "=ROUNDUP(A1;2)")
        self.sh.getCellRangeByName("A2").setValue(2.349)
        self.go(select="A2", mode="Round down (toward zero)")
        self.assertEqual(self.sh.getCellRangeByName("A2").getValue(), 2.34)

    def test_leave_formulas_and_values_off(self):
        self.fill()
        self.go(formulas=self.LEAVE)
        self.assertEqual(self.f("A4"), "=A1*A2")
        self.assertEqual(self.s("A4"), "8.42")                                             # displayed with 2 decimals
        self.fill()
        self.go(values=False)
        self.assertEqual(self.sh.getCellRangeByName("A1").getValue(), 3.14159)             # value untouched
        self.assertEqual(self.s("A1"), "3.14")                                              # but displayed rounded

    def test_remove_wrapper(self):
        self.fill()
        self.go()
        ui = self.go(formulas=self.REMOVE)
        self.assertEqual((self.f("A4"), self.f("A5")), ("=A1*A2", "=A1"))
        self.assertIn("2 ROUND wrapper(s) removed", ui.messages[-1])

    def test_zero_decimals_and_preserved_formats(self):
        self.sh.getCellRangeByName("A1").setValue(1234567.891)
        self.fmt("A1", "#,##0.00")
        self.sh.getCellRangeByName("A2").setValue(1234.5)
        self.fmt("A2", "[$₹-4009] #,##0.00")
        self.sh.getCellRangeByName("A3").setValue(0.12345)
        self.fmt("A3", "0.0%")
        self.go(select="A1:A3", places="0")
        self.assertEqual(self.s("A1"), "1,234,568")                                        # thousands separator kept
        self.assertEqual(self.s("A2"), "₹ 1,235")                                          # currency kept
        self.assertEqual(self.s("A3"), "12%")                                              # percent shows 0 decimals
        self.assertEqual(self.sh.getCellRangeByName("A3").getValue(), 0.12)                # = ROUND to whole percent

    def test_percent_rounds_the_displayed_percentage(self):
        self.sh.getCellRangeByName("A1").setValue(0.12345)
        self.fmt("A1", "0.0%")
        self.go(select="A1", places="1")
        self.assertEqual(self.sh.getCellRangeByName("A1").getValue(), 0.123)
        self.assertEqual(self.s("A1"), "12.3%")

    def test_array_formulas_and_multiple_ranges(self):
        self.sh.getCellRangeByName("A1").setValue(1.234)
        self.sh.getCellRangeByName("A2").setValue(2.345)
        self.sh.getCellRangeByName("C1:C2").setArrayFormula("=A1:A2*2")
        self.sh.getCellRangeByName("E1").setValue(9.999)
        ranges = self.doc.createInstance("com.sun.star.sheet.SheetCellRanges")
        for rng in ("A1:C2", "E1:E1"):
            ranges.addRangeAddress(self.sh.getCellRangeByName(rng).getRangeAddress(), False)
        self.ctrl.select(ranges)
        ui = _FakeUi(form={})
        make_feature("calc_decimals", ui).execute("Run", self.ctrl.getFrame(), ())
        self.assertEqual(self.sh.getCellRangeByName("E1").getValue(), 10.0)
        self.assertEqual(self.sh.getCellRangeByName("A1").getValue(), 1.23)
        self.assertIn("2 array formula(s)", ui.messages[-1])
        self.assertEqual(self.sh.getCellRangeByName("C1").getArrayFormula(), "{=A1:A2*2}")      # untouched

    def test_one_undo_step(self):
        self.fill()
        before = [(self.f(r), self.s(r)) for r in ("A1", "A2", "A4", "A5", "A8")]
        self.go()
        self.assertNotEqual([(self.f(r), self.s(r)) for r in ("A1", "A2", "A4", "A5", "A8")], before)
        self.doc.getUndoManager().undo()
        self.assertEqual([(self.f(r), self.s(r)) for r in ("A1", "A2", "A4", "A5", "A8")], before)

    def test_lone_cell_and_cancel(self):
        self.sh.getCellRangeByName("A1").setValue(1.23456)
        self.go(select="A1", places="3")
        self.assertEqual(self.sh.getCellRangeByName("A1").getValue(), 1.235)
        self.ctrl.select(self.sh.getCellRangeByName("A1"))
        make_feature("calc_decimals", _FakeUi(form=None)).execute("Run", self.ctrl.getFrame(), ())
        self.assertEqual(self.sh.getCellRangeByName("A1").getValue(), 1.235)


class CalcNumWords(Base):
    def setUp(self):
        self.doc = self.new_doc("scalc")
        self.ctrl = self.doc.getCurrentController()
        self.sh = self.doc.getSheets().getByIndex(0)

    def go(self, command, select, confirm=True, **form):
        self.ctrl.select(self.sh.getCellRangeByName(select))
        ui = _FakeUi(form=form, confirm=confirm)
        make_feature("calc_numwords", ui).execute(command, self.ctrl.getFrame(), ())
        return ui

    def s(self, ref):
        return self.sh.getCellRangeByName(ref).getString()

    def test_number_to_words_next_column(self):
        for ref, v in (("A1", 123456), ("A2", 12.45), ("A3", -42), ("A4", 0)):
            self.sh.getCellRangeByName(ref).setValue(v)
        self.sh.getCellRangeByName("A5").setString("text")
        self.sh.getCellRangeByName("A6").setFormula("=A1+1")
        ui = self.go("ToWords", "A1:A6")
        self.assertEqual([self.s("B%d" % r) for r in (1, 2, 3, 4, 5, 6)], [
            "One lakh twenty-three thousand four hundred fifty-six", "Twelve point four five", "Minus forty-two",
            "Zero", "", "One lakh twenty-three thousand four hundred fifty-seven"])
        self.assertEqual(self.sh.getCellRangeByName("A1").getValue(), 123456.0)           # source untouched
        self.assertIn("5 number(s) written in words", ui.messages[-1])
        self.assertIn("1 cell(s) skipped", ui.messages[-1])

    def test_dates_times_and_booleans_are_skipped_numbers_and_percent_are_not(self):
        nf = self.doc.getNumberFormats()
        loc = uno.createUnoStruct("com.sun.star.lang.Locale", "en", "US", "")

        def put(ref, value, code=None, formula=None):
            cell = self.sh.getCellRangeByName(ref)
            if formula:
                cell.setFormula(formula)
            else:
                cell.setValue(value)
            if code:
                key = nf.queryKey(code, loc, False)
                cell.setPropertyValue("NumberFormat", nf.addNew(code, loc) if key == -1 else key)

        put("A1", 46300, "DD-MMM-YY")                         # a date
        put("A2", 0.5, "HH:MM")                               # a time
        put("A3", 1, "BOOLEAN")                               # TRUE
        put("A4", 0, None, formula="=TODAY()")
        self.sh.getCellRangeByName("A4").setPropertyValue("NumberFormat", nf.addNew("DD-MMM-YY", loc) if nf.queryKey("DD-MMM-YY", loc, False) == -1 else nf.queryKey("DD-MMM-YY", loc, False))
        put("A5", 42)                                         # an ordinary number
        put("A6", 0.5, "0%")                                  # a percentage is still a number
        put("A7", 0, None, formula="=1.1+2.2")                # floating-point noise
        ui = self.go("ToWords", "A1:A7")
        self.assertEqual([self.s("B%d" % r) for r in range(1, 8)],
                         ["", "", "", "", "Forty-two", "Zero point five", "Three point three"])
        self.assertIn("number(s) written in words", ui.messages[-1])
        self.assertIn("4 cell(s) skipped", ui.messages[-1])
        self.assertEqual(self.sh.getCellRangeByName("A1").getValue(), 46300.0)             # the date itself is untouched

    def test_international_currency_and_case(self):
        self.sh.getCellRangeByName("A1").setValue(1234567.5)
        self.go("ToWords", "A1", system="International (million, billion)", format="Currency: Dollars / Cents",
                case="Title Case")
        self.assertEqual(self.s("B1"), "Dollars One Million Two Hundred Thirty-Four Thousand Five Hundred Sixty-Seven "
                                        "And Cents Fifty Only")
        self.sh.getCellRangeByName("A2").setValue(1234.5)
        self.go("ToWords", "A2", format="Currency: Rupees / Paise", only=False, use_and=True)
        self.assertEqual(self.s("B2"), "Rupees one thousand two hundred and thirty-four and Paise fifty")

    def test_overwrite_confirmation(self):
        self.sh.getCellRangeByName("A1").setValue(5)
        self.sh.getCellRangeByName("B1").setString("mine")
        ui = self.go("ToWords", "A1", confirm=False)
        self.assertIn("1 non-empty cell(s) to the right", ui.confirms[0])
        self.assertEqual(self.s("B1"), "mine")
        self.go("ToWords", "A1")
        self.assertEqual(self.s("B1"), "Five")

    def test_in_place_asks_and_is_one_undo_step(self):
        self.sh.getCellRangeByName("A1").setValue(7)
        self.sh.getCellRangeByName("A2").setValue(8)
        ui = self.go("ToWords", "A1:A2", confirm=False, delivery="In place (replace the cells)")
        self.assertIn("Replace 2 cell(s) in place?", ui.confirms[0])
        self.assertEqual(self.sh.getCellRangeByName("A1").getValue(), 7.0)
        self.go("ToWords", "A1:A2", delivery="In place (replace the cells)")
        self.assertEqual((self.s("A1"), self.s("A2")), ("Seven", "Eight"))
        self.doc.getUndoManager().undo()
        self.assertEqual((self.sh.getCellRangeByName("A1").getValue(), self.sh.getCellRangeByName("A2").getValue()), (7.0, 8.0))

    def test_words_to_number_and_error_report(self):
        rows = ["one lakh twenty-three thousand four hundred fifty-six", "Rupees twelve and Paise fifty Only",
                "banana", "minus two crore"]
        for i, text in enumerate(rows, start=1):
            self.sh.getCellRangeByName("A%d" % i).setString(text)
        self.sh.getCellRangeByName("A5").setValue(99)                                      # a number: skipped
        ui = self.go("FromWords", "A1:A5")
        self.assertEqual([self.sh.getCellRangeByName("B%d" % r).getValue() for r in (1, 2, 4)],
                         [123456.0, 12.5, -20000000.0])
        self.assertEqual(self.sh.getCellRangeByName("B3").getType().value, "EMPTY")        # unreadable: left empty
        self.assertIn("1 could not be converted", ui.messages[-1])
        self.assertIn("A3: unrecognised word 'banana'", ui.messages[-1])
        self.assertIn("1 cell(s) skipped", ui.messages[-1])
        self.assertEqual(self.sh.getCellRangeByName("B1").getType().value, "VALUE")

    def test_round_trip_through_the_sheet(self):
        values = [0, 7, 99, 100, 123456, 12345678, 9876543210, 12.5]
        for i, v in enumerate(values, start=1):
            self.sh.getCellRangeByName("A%d" % i).setValue(float(v))
        n = len(values)
        self.go("ToWords", "A1:A%d" % n)
        self.go("FromWords", "B1:B%d" % n)                                                  # words -> column C
        self.assertEqual([self.sh.getCellRangeByName("C%d" % r).getValue() for r in range(1, n + 1)],
                         [float(v) for v in values])

    def test_multi_range_and_nothing_to_do(self):
        self.sh.getCellRangeByName("A1").setValue(1)
        self.sh.getCellRangeByName("C1").setValue(2)
        ranges = self.doc.createInstance("com.sun.star.sheet.SheetCellRanges")
        for rng in ("A1:A1", "C1:C1"):
            ranges.addRangeAddress(self.sh.getCellRangeByName(rng).getRangeAddress(), False)
        self.ctrl.select(ranges)
        make_feature("calc_numwords", _FakeUi(form={})).execute("ToWords", self.ctrl.getFrame(), ())
        self.assertEqual((self.s("B1"), self.s("D1")), ("One", "Two"))
        ui = self.go("FromWords", "A1:A1")                                                  # only a number: nothing to read
        self.assertIn("no text to read", ui.messages[-1])

    def test_whole_column_selection_only_reads_the_used_rows(self):
        for r, v in enumerate((11, 22, 33), start=1):
            self.sh.getCellRangeByName("A%d" % r).setValue(v)
        started = time.time()
        ui = self.go("ToWords", "A:A")
        self.assertLess(time.time() - started, 15)                                          # not a million rows
        self.assertEqual([self.s("B%d" % r) for r in (1, 2, 3, 4)], ["Eleven", "Twenty-two", "Thirty-three", ""])
        self.assertIn("3 number(s) written in words", ui.messages[-1])

    def test_lakh_crore_round_trip_through_the_sheet(self):
        self.sh.getCellRangeByName("A1").setString("two lakh crore")
        self.sh.getCellRangeByName("A2").setString("1 lakh crore")
        self.go("FromWords", "A1")
        self.assertEqual(self.sh.getCellRangeByName("B1").getValue(), 2e12)
        ui = _FakeUi(form={})
        self.ctrl.select(self.sh.getCellRangeByName("A2"))
        make_feature("calc_units", ui).execute("Run", self.ctrl.getFrame(), ())
        self.assertEqual(self.sh.getCellRangeByName("B2").getValue(), 1e12)

    def test_cancel(self):
        self.sh.getCellRangeByName("A1").setValue(1)
        self.ctrl.select(self.sh.getCellRangeByName("A1"))
        make_feature("calc_numwords", _FakeUi(form=None)).execute("ToWords", self.ctrl.getFrame(), ())
        self.assertEqual(self.s("B1"), "")


class CalcUnits(Base):
    def setUp(self):
        self.doc = self.new_doc("scalc")
        self.ctrl = self.doc.getCurrentController()
        self.sh = self.doc.getSheets().getByIndex(0)

    def go(self, select, confirm=True, **form):
        self.ctrl.select(self.sh.getCellRangeByName(select))
        ui = _FakeUi(form=form, confirm=confirm)
        make_feature("calc_units", ui).execute("Run", self.ctrl.getFrame(), ())
        return ui

    def test_the_users_examples_and_more(self):
        texts = ["1 lakh", "1L", "2 cr", "2cr", "1.5k", "Rs. 2 crore", "12,34,567", "1 crore 25 lakh", "3M"]
        for i, t in enumerate(texts, start=1):
            self.sh.getCellRangeByName("A%d" % i).setString(t)
        self.go("A1:A%d" % len(texts))
        got = [self.sh.getCellRangeByName("B%d" % r).getValue() for r in range(1, len(texts) + 1)]
        self.assertEqual(got, [100000.0, 100000.0, 20000000.0, 20000000.0, 1500.0, 20000000.0, 1234567.0,
                               12500000.0, 3000000.0])
        self.assertEqual(self.sh.getCellRangeByName("B1").getType().value, "VALUE")

    def test_unreadable_and_numbers_are_reported_or_skipped(self):
        self.sh.getCellRangeByName("A1").setString("2 cr")
        self.sh.getCellRangeByName("A2").setString("abc")
        self.sh.getCellRangeByName("A3").setValue(5)
        ui = self.go("A1:A3")
        self.assertEqual(self.sh.getCellRangeByName("B1").getValue(), 20000000.0)
        self.assertIn("A2: cannot read 'abc'", ui.messages[-1])
        self.assertIn("1 cell(s) skipped", ui.messages[-1])
        self.assertIn("1 could not be converted", ui.messages[-1])

    def test_in_place_confirm_and_undo(self):
        self.sh.getCellRangeByName("A1").setString("1 lakh")
        ui = self.go("A1", confirm=False, delivery="In place (replace the cells)")
        self.assertIn("Replace 1 cell(s) in place?", ui.confirms[0])
        self.assertEqual(self.sh.getCellRangeByName("A1").getString(), "1 lakh")
        self.go("A1", delivery="In place (replace the cells)")
        self.assertEqual(self.sh.getCellRangeByName("A1").getValue(), 100000.0)
        self.doc.getUndoManager().undo()
        self.assertEqual(self.sh.getCellRangeByName("A1").getString(), "1 lakh")


class CalcWholeColumn(Base):
    """Selecting an entire column (or the whole sheet) must stay fast: tools read the used area only."""
    LIMIT = 20                                                              # seconds; a million-row read takes longer

    def setUp(self):
        self.doc = self.new_doc("scalc")
        self.ctrl = self.doc.getCurrentController()
        self.sh = self.doc.getSheets().getByIndex(0)

    def put(self, col, values, top=0):
        for r, v in enumerate(values, start=top):
            if v == "":
                continue
            cell = self.sh.getCellByPosition(col, r)
            cell.setValue(v) if isinstance(v, (int, float)) else cell.setString(v)

    def col(self, c, n):
        return [self.sh.getCellByPosition(c, r).getString() for r in range(n)]

    def run_tool(self, feature_id, command, select, ui):
        self.ctrl.select(self.sh.getCellRangeByName(select))
        started = time.time()
        make_feature(feature_id, ui).execute(command, self.ctrl.getFrame(), ())
        self.assertLess(time.time() - started, self.LIMIT, "%s on %s took too long" % (feature_id, select))
        return ui

    def test_remove_duplicates(self):
        self.put(0, ["Name", "a", "b", "A", "c"])
        on = {"header": True, "trim": True, "ignore_case": True, "numeric": True, "nonprint": True}
        ui = self.run_tool("calc_dedupe", "Run", "A:A", _FakeUi(([0], on)))
        self.assertEqual(self.col(0, 6), ["Name", "a", "b", "c", "", ""])
        self.assertIn("1 duplicate row(s) removed", ui.messages[-1])

    def test_remove_duplicates_shared_cleaning(self):
        self.put(0, ["ID", "101", "101 ", " 101", "102"])
        self.sh.getCellByPosition(0, 2).setString("101 ")      # text with NBSP
        on = {"header": True, "trim": True, "ignore_case": True, "numeric": True, "nonprint": True}
        ui = self.run_tool("calc_dedupe", "Run", "A:A", _FakeUi(([0], on)))
        self.assertEqual(self.col(0, 4), ["ID", "101", "102", ""])
        self.assertIn("2 duplicate row(s) removed", ui.messages[-1])
        self.ctrl.getModel().getUndoManager().undo()                 # one step restores everything
        self.assertEqual(self.col(0, 5), ["ID", "101", "101 ", " 101", "102"])
        exact = dict(on, trim=False, numeric=False, nonprint=False)
        self.run_tool("calc_dedupe", "Run", "A:A", _FakeUi(([0], exact)))
        self.assertEqual(self.col(0, 5), ["ID", "101", "101 ", " 101", "102"])

    def test_fill_blanks(self):
        self.put(0, ["Region", "North", "", "", "South", ""])
        ui = self.run_tool("calc_fill_blanks", "Run", "A:A", _FakeUi(form={}))
        # the blank after the last value is OUTSIDE the used area: not filled (a whole column used to
        # fill "South" into the million rows below the data)
        self.assertEqual(self.col(0, 7), ["Region", "North", "North", "North", "South", "", ""])
        self.assertIn("Filled 2 cell(s)", ui.messages[-1])

    def test_remove_blanks(self):
        self.put(0, ["a", "", "b", "", "", "c"])
        ui = self.run_tool("calc_remove_blanks", "Run", "A:A", _FakeUi(form={"mode": "Blank rows"}))
        self.assertEqual(self.col(0, 4), ["a", "b", "c", ""])
        self.assertIn("Remove 3 blank row(s)", ui.confirms[0])

    def test_highlight_and_clear(self):
        self.put(0, ["Item", "x", "y", "x", "z"])
        self.run_tool("calc_highlight", "Run", "A:A", _FakeUi(form=[
            {"mode": "Duplicates (every occurrence)", "colour": "Yellow"}, {"header": True}, {"cols": [0]}]))
        yellow = 0xFFEB84
        self.assertEqual([self.sh.getCellByPosition(0, r).getPropertyValue("CellBackColor") for r in range(5)],
                         [-1, yellow, -1, yellow, -1])
        ui = self.run_tool("calc_highlight", "Clear", "A:A", _FakeUi())
        self.assertIn("Cleared 2 highlighted cell(s)", ui.messages[-1])

    def test_compare_lists_with_typed_whole_columns(self):
        self.put(0, ["Code", "a", "b", "c"])
        self.put(2, ["Code", "b", "c", "d"])
        ui = _FakeUi(form=[{"a": "A:A", "b": "C:C", "header": True, "op": "Intersection"}, {"cols": [0]}])
        self.run_tool("calc_setops", "Run", "E1", ui)
        out = self.doc.getSheets().getByName("Intersection")
        self.assertEqual([out.getCellByPosition(0, r).getString() for r in range(4)], ["Code", "b", "c", ""])
        self.assertIn("Found in both: 2", ui.confirms[0])

    def test_transpose_in_place(self):
        self.put(0, ["a", "b", "c"])
        self.run_tool("calc_transpose", "Run", "A:A", _FakeUi(form={"dest": "In place (replace the table)"}))
        self.assertEqual([self.sh.getCellByPosition(c, 0).getString() for c in range(4)], ["a", "b", "c", ""])
        self.assertEqual(self.sh.getCellByPosition(0, 1).getString(), "")

    def test_find_missing_numbers(self):
        self.put(0, ["No", 1, 2, 5])
        ui = self.run_tool("calc_missing_seq", "Run", "A:A", _FakeUi(form={}))
        self.assertEqual(self.col(0, 7), ["No", "1", "2", "3", "4", "5", ""])
        self.assertIn("2 missing number(s)", ui.confirms[0])

    def test_sequence_keeps_its_exact_selection(self):
        ui = self.run_tool("calc_sequence", "Run", "A1:A5", _FakeUi(form={}))        # NOT clipped: it writes into it
        self.assertEqual(self.col(0, 6), ["1", "2", "3", "4", "5", ""])
        self.assertIn("Filled 5 cell(s)", ui.messages[-1])

    def test_whole_sheet_selection(self):
        self.put(0, ["Region", "North", "", "South", ""])
        self.put(2, ["x", "", "y"])
        ui = self.run_tool("calc_fill_blanks", "Run", "A1:XFD1048576", _FakeUi(form={}))
        self.assertEqual(self.col(0, 5), ["Region", "North", "North", "South", ""])
        self.assertEqual(self.col(2, 3), ["x", "x", "y"])
        self.assertIn("Filled", ui.messages[-1])

    def test_a_drawn_range_beyond_the_data_is_not_clipped(self):
        self.put(0, ["North"])
        ui = self.run_tool("calc_fill_blanks", "Run", "A1:A10", _FakeUi(form={}))      # user drew A1:A10: fill it all
        self.assertEqual(self.col(0, 11), ["North"] * 10 + [""])
        self.assertIn("Filled 9 cell(s)", ui.messages[-1])

    def test_empty_whole_column_does_nothing_quickly(self):
        ui = self.run_tool("calc_fill_blanks", "Run", "F:F", _FakeUi(form={}))
        self.assertIn("No blank cells", ui.messages[-1])

    def test_a_deliberate_wide_selection_keeps_its_width(self):
        self.put(0, [1, 2])
        self.run_tool("calc_numwords", "ToWords", "A1:C2", _FakeUi(form={}))        # A:C selected, only A used
        self.assertEqual([self.sh.getCellByPosition(3, r).getString() for r in range(2)], ["One", "Two"])  # results right of C


class CalcFunctions(Base):
    ENABLE, DISABLE = "Enable the functions", "Disable (remove them)"

    def setUp(self):
        sys.path.insert(0, os.path.join(ROOT, "src", "python"))
        from office_plugin.features.calc_functions import basic_source
        self.B = basic_source
        self.addCleanup(self.cleanup)
        self.doc = self.new_doc("scalc")
        self.sh = self.doc.getSheets().getByIndex(0)

    def cleanup(self):
        make_feature("calc_functions", _FakeUi()).uninstall()

    def manage(self, action, feature=None, ui=None):
        ui = ui or _FakeUi(form={"action": action})
        feature = feature or make_feature("calc_functions", ui)
        feature.services.register_instance("ui", ui)
        feature.execute("Manage", None, ())
        return ui, feature

    def calc(self, formula, ref="A1"):
        cell = self.sh.getCellRangeByName(ref)
        cell.setFormula(formula)
        self.doc.calculateAll()
        return cell

    def test_enable_gives_working_live_functions(self):
        ui, _ = self.manage(self.ENABLE)
        self.assertIn("now available", ui.messages[-1])
        c = self.calc("=NUMTOWORDS(123456)")
        self.assertEqual(c.getString(), "One lakh twenty-three thousand four hundred fifty-six")
        self.assertEqual(self.calc('=NUMTOWORDS(1234567;"International")').getString(),
                         "One million two hundred thirty-four thousand five hundred sixty-seven")
        self.assertEqual(self.calc('=NUMTOWORDS(1234.5;"";"INR")').getString(),
                         "Rupees one thousand two hundred thirty-four and Paise fifty only")
        self.assertEqual(self.calc('=NUMTOWORDS(5;;"usd";"upper")').getString(), "DOLLARS FIVE ONLY")
        self.assertEqual(self.calc('=WORDSTONUM("one lakh crore")').getValue(), 1e12)
        self.assertEqual(self.calc('=UNITSTONUM("2cr")').getValue(), 20000000.0)
        self.assertEqual(self.calc('=WORDSTONUM("two lakh")+1').getValue(), 200001.0)         # a real number

    def test_functions_follow_the_cells_live(self):
        self.manage(self.ENABLE)
        self.sh.getCellRangeByName("A1").setValue(21)
        self.sh.getCellRangeByName("B1").setFormula("=NUMTOWORDS(A1)")
        self.sh.getCellRangeByName("C1").setFormula("=WORDSTONUM(B1)")
        self.doc.calculateAll()
        self.assertEqual((self.sh.getCellRangeByName("B1").getString(), self.sh.getCellRangeByName("C1").getValue()),
                         ("Twenty-one", 21.0))
        self.sh.getCellRangeByName("A1").setValue(100000)                                      # no re-run needed
        self.assertEqual((self.sh.getCellRangeByName("B1").getString(), self.sh.getCellRangeByName("C1").getValue()),
                         ("One lakh", 100000.0))

    def test_bad_input_shows_value_error_text_not_a_blank_or_crash(self):
        # LibreOffice blanks a cell when Basic/Python raises and cannot raise a real cell error from Basic,
        # so bad input returns the visible text "#VALUE!" (not catchable by IFERROR; documented)
        self.manage(self.ENABLE)
        for formula in ('=NUMTOWORDS("abc")', '=WORDSTONUM("banana")', "=UNITSTONUM(\"1,5L\")", "=NUMTOWORDS(B9)",
                        '=NUMTOWORDS(1;"klingon")', "=NUMTOWORDS(1E+20)"):
            self.assertEqual(self.calc(formula).getString(), "#VALUE!", formula)
        self.assertEqual(self.calc('=IF(ISNUMBER(WORDSTONUM("banana"));1;0)').getValue(), 0.0)

    def test_enable_twice_update_and_disable(self):
        ui, feature = self.manage(self.ENABLE)
        ui2, _ = self.manage(self.ENABLE)
        self.assertIn("already available", ui2.messages[-1])
        libs = self.docs_libs()
        libs.getByName("Standard").replaceByName(self.B.MODULE_NAME, "REM old version" + chr(10))     # an outdated module
        ui3, _ = self.manage(self.ENABLE)
        self.assertIn("updated and available", ui3.messages[-1])
        self.assertEqual(self.calc("=NUMTOWORDS(7)").getString(), "Seven")
        ui4, _ = self.manage(self.DISABLE)
        self.assertIn("were removed", ui4.messages[-1])
        self.assertNotEqual(self.calc("=NUMTOWORDS(7)").getError(), 0)                         # gone
        ui5, _ = self.manage(self.DISABLE)
        self.assertIn("were not installed", ui5.messages[-1])

    def docs_libs(self):
        return Env.ctx.ServiceManager.createInstanceWithContext("com.sun.star.script.ApplicationScriptLibraryContainer", Env.ctx)

    def test_startup_resync_only_when_enabled(self):
        ui = _FakeUi()
        feature = make_feature("calc_functions", ui)
        feature.start()                                                  # not enabled: nothing installed
        self.assertFalse(feature.is_installed())
        self.manage(self.ENABLE, feature=feature)
        self.assertTrue(feature.settings.option("calc_functions", "installed"))
        feature.uninstall()                                              # e.g. the user's Standard library was reset
        self.assertFalse(feature.is_installed())
        feature.start()                                                  # enabled: the module comes back
        self.assertTrue(feature.is_installed())
        self.assertEqual(self.calc("=UNITSTONUM(\"1L\")").getValue(), 100000.0)

    def test_cancel_changes_nothing(self):
        ui, feature = self.manage(None, ui=_FakeUi(form=None))
        self.assertFalse(feature.is_installed())

    def test_generated_basic_is_well_formed(self):
        text = self.B.source()
        self.assertEqual(text.count(chr(10) + "End Function"), 4)
        for name in ("NUMTOWORDS", "WORDSTONUM", "UNITSTONUM", "OfficePluginCall"):
            self.assertIn("Function " + name, text)


class DocStore(Base):
    def setUp(self):
        sys.path.insert(0, os.path.join(ROOT, "src", "python"))
        from office_plugin.services import doc_store
        self.store = doc_store
        self.doc = self.new_doc("scalc")

    def test_round_trip_update_remove_and_defaults(self):
        self.assertIsNone(self.store.get(self.doc, "Gantt.Sheet1"))
        self.assertEqual(self.store.get(self.doc, "Gantt.Sheet1", {"d": 1}), {"d": 1})
        value = {"table": "A1:E9", "map": {"start": 2, "name": None}, "txt": "ünï \"q\" ; , ."}
        self.store.put(self.doc, "Gantt.Sheet1", value)
        self.assertEqual(self.store.get(self.doc, "Gantt.Sheet1"), value)
        self.store.put(self.doc, "Gantt.Sheet1", {"v": 2})                    # overwrite
        self.assertEqual(self.store.get(self.doc, "Gantt.Sheet1"), {"v": 2})
        self.store.remove(self.doc, "Gantt.Sheet1")
        self.assertIsNone(self.store.get(self.doc, "Gantt.Sheet1"))
        self.store.remove(self.doc, "Gantt.Sheet1")                           # removing twice is fine

    def test_foreign_or_corrupt_property_reads_as_default(self):
        from com.sun.star.beans.PropertyAttribute import REMOVEABLE
        self.doc.getDocumentProperties().getUserDefinedProperties().addProperty(
            "OfficePlugin.Broken", REMOVEABLE, "{not json")
        self.assertEqual(self.store.get(self.doc, "Broken", "fallback"), "fallback")

    def test_survives_saving_as_ods_and_xlsx(self):
        value = {"table": "A1:E9", "txt": "x" * 3000}
        self.store.put(self.doc, "Gantt.Sheet1", value)
        for ext, flt in (("ods", "calc8"), ("xlsx", "Calc MS Excel 2007 XML")):
            path = os.path.join(tempfile.gettempdir(), "opl_store." + ext)
            if os.path.exists(path):
                os.remove(path)
            self.doc.storeToURL(uno.systemPathToFileUrl(path), (pv("FilterName", flt),))
            again = Env.desktop.loadComponentFromURL(uno.systemPathToFileUrl(path), "_blank", 0, (pv("Hidden", True),))
            try:
                self.assertEqual(self.store.get(again, "Gantt.Sheet1"), value, ext)
            finally:
                again.close(True)


class CalcDedupe(Base):
    """Runs the real feature code against a real Calc document (client side, remote ctx)."""

    TABLE = [("Name", "City", "Note"), ("Ann", "Pune", "x"), ("ann ", "Pune", "y"),
             ("Bob", "Pune", "z"), ("Ann", "Delhi", "w"), ("Bob", "Pune", "v")]

    def make_feature(self, ui):
        return make_feature("calc_dedupe", ui)

    def fill(self, doc):
        sh = doc.getSheets().getByIndex(0)
        for r, row in enumerate(self.TABLE):
            for c, value in enumerate(row):
                sh.getCellByPosition(c, r).setString(value)
        return sh

    def rows(self, sh, n=8):
        return [tuple(sh.getCellByPosition(c, r).getString() for c in range(3)) for r in range(n)]

    def test_removes_duplicates_in_one_undoable_step(self):
        doc = self.new_doc("scalc")
        sh = self.fill(doc)
        ctrl = doc.getCurrentController()
        ctrl.select(sh.getCellRangeByName("A1"))           # single cell -> region is auto-expanded
        ui = _FakeUi(([0, 1], {"header": True, "trim": True, "ignore_case": True, "numeric": True, "nonprint": True}))   # compare Name+City
        self.make_feature(ui).execute("Run", ctrl.getFrame(), ())
        self.assertEqual(self.rows(sh), [
            ("Name", "City", "Note"), ("Ann", "Pune", "x"), ("Bob", "Pune", "z"),
            ("Ann", "Delhi", "w"), ("", "", ""), ("", "", ""), ("", "", ""), ("", "", "")])
        self.assertIn("2 duplicate row(s) removed", ui.messages[-1])
        self.assertTrue(ui.prompt_items[0].endswith("Name"))        # header text shown in the list
        doc.getUndoManager().undo()                                 # ONE undo restores everything
        self.assertEqual(self.rows(sh, 6), [tuple(r) if len(r) == 3 else r for r in
                                            [("Name", "City", "Note"), ("Ann", "Pune", "x"),
                                             ("ann ", "Pune", "y"), ("Bob", "Pune", "z"),
                                             ("Ann", "Delhi", "w"), ("Bob", "Pune", "v")]])

    def test_only_selected_columns_shift(self):
        doc = self.new_doc("scalc")
        sh = self.fill(doc)
        ctrl = doc.getCurrentController()
        ctrl.select(sh.getCellRangeByName("A1:B6"))        # Note column C is outside the range
        ui = _FakeUi(([0], {"header": True, "trim": True, "ignore_case": True, "numeric": True, "nonprint": True}))      # Name only
        self.make_feature(ui).execute("Run", ctrl.getFrame(), ())
        self.assertEqual([sh.getCellByPosition(0, r).getString() for r in range(6)],
                         ["Name", "Ann", "Bob", "", "", ""])
        self.assertEqual([sh.getCellByPosition(2, r).getString() for r in range(6)],
                         ["Note", "x", "y", "z", "w", "v"])         # untouched

    def test_no_duplicates_and_cancel(self):
        doc = self.new_doc("scalc")
        sh = self.fill(doc)
        ctrl = doc.getCurrentController()
        ctrl.select(sh.getCellRangeByName("A1:C6"))
        ui = _FakeUi(([0, 1, 2], {"header": True, "trim": True, "ignore_case": False, "numeric": True, "nonprint": True}))
        self.make_feature(ui).execute("Run", ctrl.getFrame(), ())
        self.assertIn("No duplicate rows found", ui.messages[-1])
        self.assertEqual(self.rows(sh, 6)[5], ("Bob", "Pune", "v"))
        cancelled = _FakeUi(None)
        self.make_feature(cancelled).execute("Run", ctrl.getFrame(), ())
        self.assertEqual(cancelled.messages, [])
        self.assertEqual(self.rows(sh, 6)[5], ("Bob", "Pune", "v"))


class Impress(Base):
    def test_slide_change_is_a_jump(self):
        doc = self.new_doc("simpress")
        ctrl = doc.getCurrentController()
        pages = doc.getDrawPages()
        pages.insertNewByIndex(1)
        pages.insertNewByIndex(2)
        self.run_cmd(doc, "navigation.Back")
        time.sleep(0.2)
        # NB: controller.setCurrentPage() fires no events; the GUI's slide switching fires the
        # "CurrentPage" property change, which is what setPropertyValue reproduces.
        ctrl.setPropertyValue("CurrentPage", pages.getByIndex(0))
        time.sleep(0.3)
        ctrl.setPropertyValue("CurrentPage", pages.getByIndex(2))
        time.sleep(0.3)
        self.assertEqual(ctrl.getCurrentPage().Number, 3)
        self.run_cmd(doc, "navigation.Back")
        self.assertEqual(ctrl.getCurrentPage().Number, 1)
        self.run_cmd(doc, "navigation.Forward")
        self.assertEqual(ctrl.getCurrentPage().Number, 3)


class Writer(Base):
    def test_far_move_back_forward_and_edit_tracking(self):
        doc = self.new_doc("swriter")
        ctrl = doc.getCurrentController()
        text = doc.getText()
        for i in range(400):
            text.insertString(text.getEnd(), "Line %d of the test document\n" % i, False)
        self.run_cmd(doc, "navigation.Back")
        time.sleep(0.2)
        vc = ctrl.getViewCursor()
        vc.gotoStart(False)
        start_y = vc.getPosition().Y
        vc.gotoEnd(False)
        end_y = vc.getPosition().Y
        if end_y - start_y < 40000:
            self.skipTest("hidden Writer layout not computed (y=%s..%s)" % (start_y, end_y))
        self.run_cmd(doc, "navigation.Back")
        self.assertLess(abs(vc.getPosition().Y - start_y), 1000)
        self.run_cmd(doc, "navigation.Forward")
        self.assertLess(abs(vc.getPosition().Y - end_y), 1000)


class WriterLink(Base):
    def test_real_bookmark_hyperlink_records_the_link_position(self):
        doc = self.new_doc("swriter", hidden=False)
        ctrl = doc.getCurrentController()
        text = doc.getText()
        text.insertString(text.getEnd(), "Jump to the end: ", False)
        link_end = text.getEnd()
        cur = text.createTextCursorByRange(link_end)
        cur.goLeft(0, False)
        text.insertString(text.getEnd(), "CLICK-ME" + chr(10), False)
        for i in range(400):
            text.insertString(text.getEnd(), "Line %d of the test document" % i + chr(10), False)
        mark = doc.createInstance("com.sun.star.text.Bookmark")
        mark.setName("endmark")
        text.insertTextContent(text.getEnd(), mark, False)
        # make "CLICK-ME" a hyperlink to the bookmark
        search = doc.createSearchDescriptor()
        search.SearchString = "CLICK-ME"
        found = doc.findFirst(search)
        found.setPropertyValue("HyperLinkURL", "#endmark")
        vc = ctrl.getViewCursor()
        vc.gotoRange(found.getStart(), False)
        vc.goRight(2, False)
        self.run_cmd(doc, "navigation.Back")           # attach plugin
        time.sleep(0.2)
        vc.gotoRange(found.getStart(), False)
        vc.goRight(2, False)
        start_y = vc.getPosition().Y
        helper = Env.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.DispatchHelper", Env.ctx)
        helper.executeDispatch(ctrl.getFrame(), ".uno:OpenHyperlinkOnCursor", "", 0, ())
        time.sleep(0.4)
        end_y = vc.getPosition().Y
        if end_y - start_y < 40000:
            self.skipTest("Writer did not follow the link headlessly (y %s -> %s)" % (start_y, end_y))
        self.run_cmd(doc, "navigation.Back")
        self.assertLess(abs(vc.getPosition().Y - start_y), 1500)
        self.run_cmd(doc, "navigation.Forward")
        self.assertLess(abs(vc.getPosition().Y - end_y), 1500)


class Registration(Base):
    """The generated XML was actually merged into LibreOffice's configuration."""

    def cfg(self, path):
        provider = Env.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", Env.ctx)
        return provider.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationAccess", (pv("nodepath", path),))

    def test_protocol_handler_registered(self):
        node = self.cfg("/org.openoffice.Office.ProtocolHandler/HandlerSet")
        self.assertIn("org.officeplugin.ProtocolHandler", node.getElementNames())

    def test_toolbar_and_menu_registered_with_separator(self):
        tb = self.cfg("/org.openoffice.Office.Addons/AddonUI/OfficeToolBar/org.officeplugin.toolbar")
        urls = [tb.getByName(n).getByName("URL") for n in sorted(tb.getElementNames())]
        self.assertEqual(urls.count("private:separator"), 4)    # navigation | data tools | compare | planning | numbers
        self.assertIn("vnd.officeplugin:navigation.Back", urls)
        menu = self.cfg("/org.openoffice.Office.Addons/AddonUI/OfficeMenuBar/org.officeplugin.menu")
        self.assertEqual(menu.getByName("Title"), "Office Plugin")

    def test_toolbar_is_titled_and_docked_not_floating_add_on_n(self):
        url = "private:resource/toolbar/addon_org.officeplugin.toolbar"
        states = self.cfg("/org.openoffice.Office.UI.CalcWindowState/UIElements/States")
        self.assertEqual(states.getByName(url).getByName("UIName"), "Office Plugin")
        doc = self.new_doc("scalc", hidden=False)
        layout = doc.getCurrentController().getFrame().LayoutManager
        self.assertTrue(layout.isElementDocked(url))
        self.assertFalse(layout.isElementFloating(url))

    def test_accelerators_registered(self):
        # Ask LibreOffice's real accelerator manager which command a key press resolves to.
        from com.sun.star.awt import Key, KeyEvent
        from com.sun.star.awt.KeyModifier import SHIFT, MOD1, MOD2
        supplier = Env.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.ui.ModuleUIConfigurationManagerSupplier", Env.ctx)
        calc, writer = "com.sun.star.sheet.SpreadsheetDocument", "com.sun.star.text.TextDocument"
        impress = "com.sun.star.presentation.PresentationDocument"
        wanted = [
            (calc, Key.LEFT, MOD2, "navigation.Back"),
            (calc, Key.SEMICOLON, MOD2, "calc_visible.Select"),
            (calc, Key.F5, MOD2, "last_edit.GoTo"),
            (writer, Key.RIGHT, MOD2, "navigation.Forward"),
            (writer, Key.F5, MOD2, "last_edit.GoTo"),
            (impress, Key.LEFT, MOD2, "navigation.Back"),
        ]
        for module, code, mods, command in wanted:
            mgr = supplier.getUIConfigurationManager(module).getShortCutManager()
            ev = KeyEvent()
            ev.KeyCode, ev.Modifiers = code, mods
            self.assertEqual(mgr.getCommandByKeyEvent(ev), "vnd.officeplugin:" + command,
                             "%s key %s/%s" % (module, code, mods))

    def test_our_keys_do_not_shadow_existing_bindings(self):
        """Every default LibreOffice binding must still resolve to its original command."""
        from com.sun.star.awt import Key, KeyEvent
        from com.sun.star.awt.KeyModifier import MOD1, SHIFT
        supplier = Env.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.ui.ModuleUIConfigurationManagerSupplier", Env.ctx)
        mgr = supplier.getUIConfigurationManager("com.sun.star.text.TextDocument").getShortCutManager()
        ev = KeyEvent()
        ev.KeyCode, ev.Modifiers = Key.F5, SHIFT          # Writer's own Shift+F5
        self.assertEqual(mgr.getCommandByKeyEvent(ev), ".uno:RestoreEditingView")
        ev.KeyCode, ev.Modifiers = Key.LEFT, MOD1         # Ctrl+Left unchanged
        self.assertTrue(mgr.getCommandByKeyEvent(ev).startswith(".uno:"))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("tests", nargs="*")
    args, rest = ap.parse_known_args()
    setup_env()
    try:
        suite = unittest.defaultTestLoader.loadTestsFromNames(args.tests, sys.modules[__name__]) \
            if args.tests else unittest.defaultTestLoader.loadTestsFromModule(sys.modules[__name__])
        result = unittest.TextTestRunner(verbosity=2).run(suite)
    finally:
        print("\n--- plugin log (tail) ---")
        print("\n".join(log_text().splitlines()[-40:]))
        teardown_env()
    sys.exit(0 if result.wasSuccessful() else 1)


if __name__ == "__main__":
    main()
