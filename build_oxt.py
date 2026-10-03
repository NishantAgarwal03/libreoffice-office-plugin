"""Build dist/office_plugin_lo.oxt.

Discovers every src/python/office_plugin/features/<id>/feature.json and GENERATES all
LibreOffice registration XML from them (ProtocolHandler.xcu, Addons.xcu,
Accelerators.xcu, Jobs.xcu, META-INF/manifest.xml, description.xml). Adding a feature
therefore never requires editing this file or any hand-written XML.

Usage:  python build_oxt.py [--out PATH]      (stdlib only; any Python 3.8+)
"""

import argparse
import os
import sys
import zipfile
from xml.sax.saxutils import escape

ROOT = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(ROOT, "src")
PYDIR = os.path.join(SRC, "python")
sys.path.insert(0, PYDIR)

from office_plugin import __version__  # noqa: E402
from office_plugin.core import manifest as M  # noqa: E402

EXT_ID = "org.officeplugin.office_plugin_lo"
PROTOCOL_IMPL = "org.officeplugin.ProtocolHandler"
JOB_IMPL = "org.officeplugin.StartupJob"
MENU_NODE = "org.officeplugin.menu"
TOOLBAR_NODE = "org.officeplugin.toolbar"
MENU_TITLE = "Office Plugin"

_NS = ('xmlns:oor="http://openoffice.org/2001/registry" '
       'xmlns:xs="http://www.w3.org/2001/XMLSchema"')
_HEAD = '<?xml version="1.0" encoding="UTF-8"?>\n'


# ======================================================================
# helpers
# ======================================================================
def _prop(name, value, typ="xs:string"):
    return '<prop oor:name="%s" oor:type="%s"><value>%s</value></prop>' % (name, typ, escape(value))


def _title(text):
    return ('<prop oor:name="Title" oor:type="xs:string"><value/>'
            '<value xml:lang="en-US">%s</value></prop>' % escape(text))


def _contexts(apps):
    return ",".join(M.APPS[a] for a in apps)


def _ordered_commands(manifests, key):
    """[(manifest, command)] having cmd[key] true, sorted by group then order.

    Groups are ordered by their smallest `order`, so a group stays contiguous.
    """
    items = [(m, c) for m in manifests for c in m["commands"] if c[key]]
    group_rank = {}
    for _, c in items:
        group_rank[c["group"]] = min(group_rank.get(c["group"], 10 ** 9), c["order"])
    items.sort(key=lambda mc: (group_rank[mc[1]["group"]], mc[1]["group"], mc[1]["order"]))
    return items


def _with_separators(items):
    """Yield ('sep', apps) between groups, else ('cmd', manifest, command).

    A separator is only shown where both neighbouring groups are visible, so its context
    is the intersection of the two groups' apps (avoids dangling separators, e.g. Calc-only
    buttons absent in Writer).
    """
    previous = None
    for m, c in items:
        if previous is not None and previous[1]["group"] != c["group"]:
            shared = [a for a in previous[0]["apps"] if a in m["apps"]]
            if shared:
                yield ("sep", shared)
        yield ("cmd", m, c)
        previous = (m, c)


def _icon_paths(cmd):
    base = cmd["icon"]
    return "icons/%s_16.png" % base, "icons/%s_26.png" % base


# ======================================================================
# generators (each returns the file text)
# ======================================================================
def gen_protocol_handler():
    return _HEAD + (
        '<oor:component-data %s oor:name="ProtocolHandler" oor:package="org.openoffice.Office">\n'
        '  <node oor:name="HandlerSet">\n'
        '    <node oor:name="%s" oor:op="replace">\n'
        '      <prop oor:name="Protocols" oor:type="oor:string-list">'
        '<value>vnd.officeplugin:*</value></prop>\n'
        '    </node>\n  </node>\n</oor:component-data>\n') % (_NS, PROTOCOL_IMPL)


def gen_jobs():
    return _HEAD + (
        '<oor:component-data %s oor:name="Jobs" oor:package="org.openoffice.Office">\n'
        '  <node oor:name="Jobs">\n'
        '    <node oor:name="%s" oor:op="replace">\n'
        '      <prop oor:name="Service" oor:type="xs:string"><value>%s</value></prop>\n'
        '    </node>\n  </node>\n'
        '  <node oor:name="Events">\n'
        '    <node oor:name="onFirstVisibleTask" oor:op="fuse">\n'
        '      <node oor:name="JobList">\n'
        '        <node oor:name="%s" oor:op="replace"/>\n'
        '      </node>\n    </node>\n  </node>\n</oor:component-data>\n'
    ) % (_NS, JOB_IMPL, JOB_IMPL, JOB_IMPL)


def gen_accelerators(manifests):
    by_module = {}
    for module, lo_key, url in M.accelerators(manifests):   # raises on key conflicts
        by_module.setdefault(module, []).append((lo_key, url))
    out = [_HEAD, '<oor:component-data %s oor:name="Accelerators" '
                  'oor:package="org.openoffice.Office">\n' % _NS,
           '  <node oor:name="PrimaryKeys">\n    <node oor:name="Modules">\n']
    for module in sorted(by_module):
        out.append('      <node oor:name="%s">\n' % module)
        for lo_key, url in sorted(by_module[module]):
            out.append('        <node oor:name="%s" oor:op="replace">'
                       '<prop oor:name="Command"><value xml:lang="en-US">%s</value></prop></node>\n'
                       % (lo_key, escape(url)))
        out.append('      </node>\n')
    out.append('    </node>\n  </node>\n</oor:component-data>\n')
    return "".join(out)


def _item(index, url, title, contexts, target="_self"):
    return ('<node oor:name="m%03d" oor:op="replace">%s%s%s%s</node>\n' % (
        index, _prop("URL", url), _title(title), _prop("Target", target),
        _prop("Context", contexts) if contexts else ""))


def gen_addons(manifests):
    all_apps = sorted({a for m in manifests for a in m["apps"]}, key=list(M.APPS).index)

    # ---- menu --------------------------------------------------------------
    menu = []
    for i, entry in enumerate(_with_separators(_ordered_commands(manifests, "menu")), 1):
        if entry[0] == "sep":
            menu.append(_item(i, "private:separator", "", _contexts(entry[1])))
        else:
            _, m, c = entry
            menu.append(_item(i, M.command_url(m["id"], c["name"]), c["label"], _contexts(m["apps"])))

    # ---- toolbar -----------------------------------------------------------
    tool_items = _ordered_commands(manifests, "toolbar")
    toolbar, images = [], []
    for i, entry in enumerate(_with_separators(tool_items), 1):
        if entry[0] == "sep":
            toolbar.append(_item(i, "private:separator", "", _contexts(entry[1])))
        else:
            _, m, c = entry
            url = M.command_url(m["id"], c["name"])
            toolbar.append(_item(i, url, c["tooltip"] or c["label"], _contexts(m["apps"])))
            small, big = _icon_paths(c)
            images.append(
                '<node oor:name="org.officeplugin.image.%s_%s" oor:op="replace">%s'
                '<node oor:name="UserDefinedImages">%s%s</node></node>\n' % (
                    m["id"], c["name"], _prop("URL", url),
                    _prop("ImageSmallURL", "%origin%/" + small),
                    _prop("ImageBigURL", "%origin%/" + big)))

    return _HEAD + (
        '<oor:component-data %s oor:name="Addons" oor:package="org.openoffice.Office">\n'
        '<node oor:name="AddonUI">\n'
        '<node oor:name="OfficeMenuBar"><node oor:name="%s" oor:op="replace">%s%s%s'
        '<node oor:name="Submenu">\n%s</node></node></node>\n'
        '<node oor:name="OfficeToolBar"><node oor:name="%s" oor:op="replace">\n%s</node></node>\n'
        '<node oor:name="Images">\n%s</node>\n'
        '</node>\n</oor:component-data>\n'
    ) % (_NS, MENU_NODE, _title(MENU_TITLE), _prop("Target", "_self"),
         _prop("Context", _contexts(all_apps)), "".join(menu),
         TOOLBAR_NODE, "".join(toolbar), "".join(images))


_WINDOW_STATE = {"writer": "WriterWindowState", "calc": "CalcWindowState",
                 "impress": "ImpressWindowState", "draw": "DrawWindowState"}


def window_state_files(manifests):
    """{file name: xml}: gives the add-on toolbar a real title (LibreOffice otherwise shows 'Add-On 2') and
    docks it in the top toolbar area the first time (the user's own placement overrides this afterwards)."""
    apps = {a for m, c in _ordered_commands(manifests, "toolbar") for a in m["apps"]}
    files = {}
    for app in sorted(apps, key=list(M.APPS).index):
        component = _WINDOW_STATE[app]
        files["WindowState_%s.xcu" % app] = _HEAD + (
            '<oor:component-data %s oor:name="%s" oor:package="org.openoffice.Office.UI">\n'
            '<node oor:name="UIElements"><node oor:name="States">\n'
            '<node oor:name="private:resource/toolbar/addon_%s" oor:op="replace">'
            '<prop oor:name="UIName" oor:type="xs:string"><value/><value xml:lang="en-US">%s</value></prop>'
            '<prop oor:name="Docked" oor:type="xs:boolean"><value>true</value></prop>'
            '<prop oor:name="DockingArea" oor:type="xs:int"><value>0</value></prop>'
            '<prop oor:name="Visible" oor:type="xs:boolean"><value>true</value></prop>'
            '</node>\n</node></node>\n</oor:component-data>\n') % (
                _NS, component, TOOLBAR_NODE, escape(MENU_TITLE))
    return files


def gen_manifest_xml(manifests=()):
    entries = [
        ("application/vnd.sun.star.uno-component;type=Python", "python/officeplugin.py"),
        ("application/vnd.sun.star.configuration-data", "ProtocolHandler.xcu"),
        ("application/vnd.sun.star.configuration-data", "Jobs.xcu"),
        ("application/vnd.sun.star.configuration-data", "Addons.xcu"),
        ("application/vnd.sun.star.configuration-data", "Accelerators.xcu"),
    ] + [("application/vnd.sun.star.configuration-data", name) for name in sorted(window_state_files(manifests))]
    rows = "".join('  <manifest:file-entry manifest:media-type="%s" manifest:full-path="%s"/>\n'
                   % e for e in entries)
    return _HEAD + ('<manifest:manifest xmlns:manifest="http://openoffice.org/2001/manifest">\n'
                    '%s</manifest:manifest>\n') % rows


def gen_description_xml():
    return _HEAD + (
        '<description xmlns="http://openoffice.org/extensions/description/2006" '
        'xmlns:d="http://openoffice.org/extensions/description/2006" '
        'xmlns:xlink="http://www.w3.org/1999/xlink">\n'
        '  <identifier value="%s"/>\n  <version value="%s"/>\n'
        '  <dependencies><d:OpenOffice.org-minimal-version value="4.1" d:name="LibreOffice 4.1"/>'
        '</dependencies>\n'
        '  <display-name><name lang="en">%s</name></display-name>\n'
        '  <extension-description><src xlink:href="description/desc_en.txt" lang="en"/>'
        '</extension-description>\n</description>\n') % (EXT_ID, __version__, MENU_TITLE)


def gen_description_text(manifests):
    lines = ["Office Plugin %s - features LibreOffice is missing compared to MS Office." % __version__,
             "", "Included features:"]
    for m in manifests:
        lines.append("* %s (%s): %s" % (m["name"], ", ".join(m["apps"]), m["description"]))
    lines += ["", "Find everything on the 'Office Plugin' toolbar (View > Toolbars) and menu."]
    return "\n".join(lines) + "\n"


# ======================================================================
# build
# ======================================================================
def check_icons(manifests):
    missing = []
    for m in manifests:
        for c in m["commands"]:
            if c["toolbar"]:
                for rel in _icon_paths(c):
                    if not os.path.isfile(os.path.join(SRC, rel)):
                        missing.append(rel)
    if missing:
        raise SystemExit("missing icon file(s): %s\n(run tools/make_icons.py)" % ", ".join(missing))


def generated_files(manifests):
    return {
        **window_state_files(manifests),
        "META-INF/manifest.xml": gen_manifest_xml(manifests),
        "description.xml": gen_description_xml(),
        "description/desc_en.txt": gen_description_text(manifests),
        "ProtocolHandler.xcu": gen_protocol_handler(),
        "Jobs.xcu": gen_jobs(),
        "Addons.xcu": gen_addons(manifests),
        "Accelerators.xcu": gen_accelerators(manifests),
    }


def build(out_path):
    features_dir = os.path.join(PYDIR, "office_plugin", "features")
    manifests = M.discover_manifests(features_dir)
    check_icons(manifests)
    files = generated_files(manifests)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("META-INF/manifest.xml", files.pop("META-INF/manifest.xml"))
        for name, text in files.items():
            zf.writestr(name, text)
        for base in (PYDIR, os.path.join(SRC, "icons")):
            for dirpath, dirnames, filenames in os.walk(base):
                dirnames[:] = [d for d in dirnames if d != "__pycache__" and d != "_template"]
                for fn in filenames:
                    if fn.endswith((".pyc", ".pyo")):
                        continue
                    full = os.path.join(dirpath, fn)
                    zf.write(full, os.path.relpath(full, SRC).replace(os.sep, "/"))
    return out_path, manifests


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--out", default=os.path.join(ROOT, "dist", "office_plugin_lo.oxt"))
    args = ap.parse_args(argv)
    path, manifests = build(args.out)
    print("built %s  (%d features: %s)" % (path, len(manifests), ", ".join(m["id"] for m in manifests)))


if __name__ == "__main__":
    main()
