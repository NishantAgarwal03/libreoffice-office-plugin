"""Calc formatting helpers: cell styles, conditional formats, date number formats.

Purpose: reusable building blocks for tools that format cells live (Gantt bars, future banding...).
Inputs:  UNO document / ranges, style names, formula strings (API grammar: English function names,
         ';' separators, A1 references relative to the top-left cell of the range).
Outputs: style names, number-format keys; conditional formats set on / cleared from a range.
Must not: import uno at module level.
"""


def ensure_cell_style(doc, name, back_color, font_color=None, bold=False):
    """Create the cell style if missing and (re)apply the colours; returns `name`."""
    styles = doc.getStyleFamilies().getByName("CellStyles")
    if not styles.hasByName(name):
        styles.insertByName(name, doc.createInstance("com.sun.star.style.CellStyle"))
    style = styles.getByName(name)
    style.setPropertyValue("CellBackColor", back_color)
    if font_color is not None:
        style.setPropertyValue("CharColor", font_color)
    if bold:
        style.setPropertyValue("CharWeight", 150.0)
    return name


def ensure_date_format(doc, pattern, language="en", country="US"):
    """Number-format key for a date/number pattern such as 'DD-MMM' (created if missing)."""
    import uno
    formats = doc.getNumberFormats()
    locale = uno.createUnoStruct("com.sun.star.lang.Locale", language, country, "")
    key = formats.queryKey(pattern, locale, False)
    return formats.addNew(pattern, locale) if key == -1 else key


def set_conditional_formats(rng, entries):
    """Replace the range's conditional format by `entries` = [(formula, style_name)], first match
    wins (list order is priority). Formulas are relative to the range's top-left cell."""
    import uno
    addr = rng.getRangeAddress()
    origin = uno.createUnoStruct("com.sun.star.table.CellAddress", addr.Sheet, addr.StartColumn, addr.StartRow)
    operator = uno.Enum("com.sun.star.sheet.ConditionOperator", "FORMULA")
    conditional = rng.getPropertyValue("ConditionalFormat")
    conditional.clear()
    for formula, style in entries:
        conditional.addNew((_pv("Operator", operator), _pv("Formula1", formula),
                            _pv("StyleName", style), _pv("SourcePosition", origin)))
    rng.setPropertyValue("ConditionalFormat", conditional)


def clear_conditional_formats(rng):
    conditional = rng.getPropertyValue("ConditionalFormat")
    if conditional.getCount():
        conditional.clear()
        rng.setPropertyValue("ConditionalFormat", conditional)


def conditional_entries(rng):
    """[(formula, style_name)] currently set on the range (for tests / diagnostics)."""
    conditional = rng.getPropertyValue("ConditionalFormat")
    return [(conditional.getByIndex(i).getFormula1(), conditional.getByIndex(i).getStyleName())
            for i in range(conditional.getCount())]


def _pv(name, value):
    import uno
    return uno.createUnoStruct("com.sun.star.beans.PropertyValue", name, 0, value, 0)
