"""Capture / restore 'where the user is' in Writer, Calc and Impress/Draw.

capture() is cheap enough to run on every selection change. Writer positions keep a
*live* text cursor, which LibreOffice keeps pointing at the same text while the
document is edited (and which does not modify the document).
"""

from office_plugin.core.position import Position


class PositionService:
    def __init__(self, ctx, doc_utils, log):
        self.ctx = ctx
        self.docs = doc_utils
        self.log = log

    # ======================================================================
    # capture
    # ======================================================================
    def capture(self, controller, page=None):
        """-> Position, or None when the current selection is not a trackable place
        (e.g. a drawing object selected in Calc).

        `page` (Impress/Draw only) overrides controller.getCurrentPage(): page-change
        events fire before the controller reflects the new page."""
        try:
            model = controller.getModel()
            kind = self.docs.kind(model)
            if kind == "writer":
                return self._capture_writer(controller, model)
            if kind == "calc":
                return self._capture_calc(controller, model)
            if kind in ("impress", "draw"):
                return self._capture_draw(controller, model, kind, page)
        except Exception:
            self.log.debug("capture failed", exc_info=True)
        return None

    def _capture_writer(self, controller, model):
        vc = controller.getViewCursor()
        y = vc.getPosition().Y
        start = vc.getStart()
        cursor = start.getText().createTextCursorByRange(start)
        return Position("writer", self.docs.doc_id(model), container=0, coords=(y,),
                        handle=cursor)

    def _capture_calc(self, controller, model):
        sel = controller.getSelection()
        addr = None
        if sel.supportsService("com.sun.star.sheet.SheetCellRanges"):
            addrs = sel.getRangeAddresses()
            if addrs:
                addr = addrs[0]
        elif sel.supportsService("com.sun.star.sheet.SheetCellRange"):
            addr = sel.getRangeAddress()
        if addr is None:
            return None
        return Position(
            "calc", self.docs.doc_id(model), container=addr.Sheet,
            coords=(addr.StartRow, addr.StartColumn),
            handle=(addr.Sheet, addr.StartColumn, addr.StartRow, addr.EndColumn, addr.EndRow))

    def _capture_draw(self, controller, model, kind, page=None):
        page = page if page is not None else controller.getCurrentPage()
        number = page.Number if page is not None else 0
        names = ()
        try:
            sel = controller.getSelection()
            if sel is not None and sel.supportsService("com.sun.star.drawing.ShapeCollection"):
                names = tuple(sel.getByIndex(i).Name for i in range(sel.getCount()))
        except Exception:
            names = ()
        return Position(kind, self.docs.doc_id(model), container=number, coords=(),
                        handle=(number, names))

    def capture_for(self, pos):
        """Current live Position of the document `pos` belongs to (None if it is gone)."""
        model = self.docs.find_model(pos.doc_id)
        controller = model.getCurrentController() if model is not None else None
        return self.capture(controller) if controller is not None else None

    # ======================================================================
    # restore
    # ======================================================================
    def restore(self, pos):
        """Move the UI back to `pos`. Returns True on success."""
        model = self.docs.find_model(pos.doc_id)
        if model is None:
            self.log.info("restore: document %s is no longer open", pos.doc_id)
            return False
        controller = model.getCurrentController()
        if controller is None:
            return False
        try:
            self.docs.bring_to_front(controller)
            if pos.kind == "writer":
                ok = self._restore_writer(controller, pos)
            elif pos.kind == "calc":
                ok = self._restore_calc(controller, model, pos)
            elif pos.kind in ("impress", "draw"):
                ok = self._restore_draw(controller, model, pos)
            else:
                ok = False
            return ok
        except Exception:
            self.log.info("restore failed for %r", pos, exc_info=True)
            return False

    def _restore_writer(self, controller, pos):
        controller.getViewCursor().gotoRange(pos.handle, False)
        return True

    def _restore_calc(self, controller, model, pos):
        sheet_idx, c1, r1, c2, r2 = pos.handle
        sheets = model.getSheets()
        if sheet_idx >= sheets.getCount():
            return False
        sheet = sheets.getByIndex(sheet_idx)
        rng = sheet.getCellRangeByPosition(c1, r1, c2, r2)
        controller.setActiveSheet(sheet)
        # GoToCell is what the Name Box uses: it selects *and* scrolls the cell into view.
        target = rng.getPropertyValue("AbsoluteName") if rng.getPropertySetInfo() \
            .hasPropertyByName("AbsoluteName") else None
        if target:
            self._dispatch(controller, ".uno:GoToCell", ToPoint=target)
        else:
            controller.select(rng)
        return True

    def _restore_draw(self, controller, model, pos):
        number, names = pos.handle
        pages = model.getDrawPages()
        if number < 1 or number > pages.getCount():
            return False
        page = pages.getByIndex(number - 1)
        controller.setCurrentPage(page)
        if names:
            try:
                coll = self.docs.create("com.sun.star.drawing.ShapeCollection")
                for i in range(page.getCount()):
                    shape = page.getByIndex(i)
                    if shape.Name and shape.Name in names:
                        coll.add(shape)
                if coll.getCount():
                    controller.select(coll)
            except Exception:
                self.log.debug("shape reselect failed", exc_info=True)
        return True

    # ======================================================================
    def _dispatch(self, controller, command, **props):
        import uno
        helper = self.docs.create("com.sun.star.frame.DispatchHelper")
        args = tuple(uno.createUnoStruct("com.sun.star.beans.PropertyValue", k, 0, v, 0)
                     for k, v in props.items())
        helper.executeDispatch(controller.getFrame(), command, "", 0, args)

    def describe(self, pos):
        """Human-readable label for logs/messages."""
        try:
            model = self.docs.find_model(pos.doc_id)
            title = model.getTitle() if model is not None else "?"
        except Exception:
            title = "?"
        return "%s %s %s" % (title, pos.container, pos.coords)
