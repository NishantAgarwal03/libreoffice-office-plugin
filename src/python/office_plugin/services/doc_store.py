"""Small JSON settings stored inside the document itself.

Purpose: let a tool remember how it was used on THIS document (e.g. the Gantt column mapping)
         across save / close / reopen, in ods and xlsx (user-defined document properties).
Inputs:  document, key, any JSON-serialisable value.
Outputs: the stored value or the default.
Must not: import uno at module level; keep keys under the 'OfficePlugin.' prefix.
"""

import json

_PREFIX = "OfficePlugin."


def _props(doc):
    return doc.getDocumentProperties().getUserDefinedProperties()


def get(doc, key, default=None):
    name = _PREFIX + key
    try:
        props = _props(doc)
        if not props.getPropertySetInfo().hasPropertyByName(name):
            return default
        return json.loads(props.getPropertyValue(name))
    except Exception:          # missing / not our JSON / unreadable: behave as 'nothing stored'
        return default


def put(doc, key, value):
    from com.sun.star.beans.PropertyAttribute import REMOVEABLE
    name, text = _PREFIX + key, json.dumps(value, separators=(",", ":"))
    props = _props(doc)
    if props.getPropertySetInfo().hasPropertyByName(name):
        props.setPropertyValue(name, text)
    else:
        props.addProperty(name, REMOVEABLE, text)


def remove(doc, key):
    name = _PREFIX + key
    props = _props(doc)
    if props.getPropertySetInfo().hasPropertyByName(name):
        props.removeProperty(name)
