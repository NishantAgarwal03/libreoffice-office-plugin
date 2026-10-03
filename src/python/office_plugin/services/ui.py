"""Small UI helpers: message/confirm boxes, code-built dialogs, range picking."""


def _strings(items):
    import uno
    return uno.Any("[]string", tuple(items))


def _shorts(items):
    import uno
    return uno.Any("[]short", tuple(int(i) for i in items))


class UiService:
    def __init__(self, ctx, doc_utils):
        self.ctx = ctx
        self.docs = doc_utils
        self._pickers = []   # keeps range-selection listeners alive until they report back

    def _parent_window(self, frame):
        frame = frame or self.docs.active_frame()
        return frame.getContainerWindow() if frame is not None else None

    def message(self, title, text, frame=None, kind="info"):
        """Modal message box. kind: info | warning | error."""
        from com.sun.star.awt.MessageBoxType import INFOBOX, WARNINGBOX, ERRORBOX
        from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
        box_type = {"info": INFOBOX, "warning": WARNINGBOX, "error": ERRORBOX}[kind]
        toolkit = self.docs.create("com.sun.star.awt.Toolkit")
        parent = self._parent_window(frame)
        if parent is None:
            return
        box = toolkit.createMessageBox(parent, box_type, BUTTONS_OK, title, text)
        box.execute()
        box.dispose()

    def choose_columns(self, title, prompt, items, options, frame=None):
        """Dialog with a multi-select list (all preselected) plus checkbox options.

        items:   list of label strings
        options: list of (key, label, default_bool)
        Returns (selected_indices, {key: bool}) or None if cancelled.
        """
        fields = [{"type": "label", "text": prompt},
                  {"type": "multilist", "name": "items", "label": "Columns", "items": list(items),
                   "selected": list(range(len(items))), "rows": 8}]
        fields += [{"type": "checkbox", "name": key, "label": label, "value": default}
                   for key, label, default in options]
        answer = self.form(title, fields, frame, width=200)
        if answer is None:
            return None
        return answer["items"], {key: answer[key] for key, _, _ in options}

    # ======================================================================
    # confirm / generic form / range picking
    # ======================================================================
    def confirm(self, title, text, frame=None):
        """OK / Cancel question box. True when OK was pressed."""
        from com.sun.star.awt.MessageBoxType import QUERYBOX
        from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK_CANCEL
        parent = self._parent_window(frame)
        if parent is None:
            return False
        toolkit = self.docs.create("com.sun.star.awt.Toolkit")
        box = toolkit.createMessageBox(parent, QUERYBOX, BUTTONS_OK_CANCEL, title, text)
        try:
            return box.execute() == 1
        finally:
            box.dispose()

    def build_form(self, title, fields, width=240):
        """Create (but do not show) a dialog from field specs -> (dialog, [(name, spec)]).

        Field specs (dicts):
          {"type": "label",    "text": "..."}                                  (multi-line ok)
          {"type": "text",     "name": n, "label": l, "value": "A1:B9"}
          {"type": "checkbox", "name": n, "label": l, "value": True}
          {"type": "dropdown", "name": n, "label": l, "items": [...], "value": item_text}
          {"type": "multilist","name": n, "label": l, "items": [...], "selected": [idx], "rows": 6}
        """
        dlg_model = self.docs.create("com.sun.star.awt.UnoControlDialogModel")

        def add(kind, name, **props):
            model = dlg_model.createInstance("com.sun.star.awt.UnoControl%sModel" % kind)
            for key, value in props.items():
                model.setPropertyValue(key, value)
            dlg_model.insertByName(name, model)
            return model

        inner = width - 16
        y = 6
        named = []
        for i, spec in enumerate(fields):
            kind = spec["type"]
            if kind == "label":
                lines = max(1, spec["text"].count("\n") + 1)
                add("FixedText", "f%d" % i, PositionX=8, PositionY=y, Width=inner,
                    Height=9 * lines + 2, Label=spec["text"], MultiLine=True)
                y += 9 * lines + 6
            elif kind == "checkbox":
                add("CheckBox", spec["name"], PositionX=8, PositionY=y, Width=inner, Height=10,
                    Label=spec["label"], State=1 if spec.get("value") else 0)
                y += 13
                named.append((spec["name"], spec))
            elif kind == "text":
                add("FixedText", "l_" + spec["name"], PositionX=8, PositionY=y, Width=inner,
                    Height=9, Label=spec["label"])
                add("Edit", spec["name"], PositionX=8, PositionY=y + 10, Width=inner, Height=12,
                    Text=str(spec.get("value", "")))
                y += 26
                named.append((spec["name"], spec))
            elif kind == "dropdown":
                items = list(spec["items"])
                add("FixedText", "l_" + spec["name"], PositionX=8, PositionY=y, Width=inner,
                    Height=9, Label=spec["label"])
                chosen = items.index(spec["value"]) if spec.get("value") in items else 0
                add("ListBox", spec["name"], PositionX=8, PositionY=y + 10, Width=inner, Height=12,
                    Dropdown=True, LineCount=min(len(items), 10),
                    StringItemList=_strings(items), SelectedItems=_shorts([chosen]))
                y += 26
                named.append((spec["name"], spec))
            elif kind == "multilist":
                items = list(spec["items"])
                rows = spec.get("rows", 6)
                add("FixedText", "l_" + spec["name"], PositionX=8, PositionY=y, Width=inner,
                    Height=9, Label=spec["label"])
                add("ListBox", spec["name"], PositionX=8, PositionY=y + 10, Width=inner,
                    Height=9 * rows, MultiSelection=True, Dropdown=False,
                    StringItemList=_strings(items),
                    SelectedItems=_shorts(spec.get("selected", range(len(items)))))
                y += 14 + 9 * rows
                named.append((spec["name"], spec))
            else:
                raise ValueError("unknown field type %r" % kind)
        add("Button", "ok", PositionX=width - 100, PositionY=y + 4, Width=44, Height=14,
            Label="OK", PushButtonType=1, DefaultButton=True)
        add("Button", "cancel", PositionX=width - 52, PositionY=y + 4, Width=44, Height=14,
            Label="Cancel", PushButtonType=2)
        dlg_model.setPropertyValue("Title", title)
        dlg_model.setPropertyValue("Width", width)
        dlg_model.setPropertyValue("Height", y + 24)
        dialog = self.docs.create("com.sun.star.awt.UnoControlDialog")
        dialog.setModel(dlg_model)
        return dialog, named

    @staticmethod
    def read_form(dialog, named):
        """Current values: text -> str, checkbox -> bool, dropdown -> item text,
        multilist -> list of selected indices."""
        values = {}
        for name, spec in named:
            control = dialog.getControl(name)
            kind = spec["type"]
            if kind == "text":
                values[name] = control.getText()
            elif kind == "checkbox":
                values[name] = bool(control.getState())
            elif kind == "dropdown":
                pos = control.getSelectedItemPos()
                values[name] = spec["items"][pos] if pos >= 0 else None
            elif kind == "multilist":
                values[name] = list(control.getSelectedItemsPos())
        return values

    def form(self, title, fields, frame=None, width=240):
        """Show a modal dialog; returns {name: value} on OK, None on Cancel."""
        dialog, named = self.build_form(title, fields, width)
        toolkit = self.docs.create("com.sun.star.awt.Toolkit")
        dialog.createPeer(toolkit, self._parent_window(frame))
        try:
            if dialog.execute() != 1:
                return None
            return self.read_form(dialog, named)
        finally:
            dialog.dispose()

    def pick_range(self, controller, title, initial, callback):
        """Let the user mark a range with the mouse (LibreOffice's range-selection popup).

        Non-modal: returns immediately and later calls callback(address_text) when the user
        confirms, or callback(None) when aborted. Returns False (callback never called) when
        the view cannot do range selection, so the caller can fall back to typed addresses.
        """
        if not hasattr(controller, "startRangeSelection"):
            return False
        import uno
        import unohelper
        from com.sun.star.sheet import XRangeSelectionListener

        service = self

        class Listener(unohelper.Base, XRangeSelectionListener):
            def done(self, ev):
                service._finish_pick(controller, self, callback, ev.RangeDescriptor)

            def aborted(self, ev):
                service._finish_pick(controller, self, callback, None)

            def disposing(self, ev):
                pass

        def prop(name, value):
            return uno.createUnoStruct("com.sun.star.beans.PropertyValue", name, 0, value, 0)

        listener = Listener()
        self._pickers.append(listener)
        try:
            controller.addRangeSelectionListener(listener)
            controller.startRangeSelection((prop("Title", title), prop("InitialValue", initial or ""),
                                            prop("CloseOnMouseRelease", False)))
        except Exception:
            self._pickers.remove(listener)
            return False
        return True

    def _finish_pick(self, controller, listener, callback, text):
        try:
            controller.removeRangeSelectionListener(listener)
        except Exception:
            pass
        if listener in self._pickers:
            self._pickers.remove(listener)
        callback(text)
