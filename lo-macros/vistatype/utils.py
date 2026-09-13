"""
vistatype.utils
================
Shared helpers for talking to a LibreOffice Writer document over UNO.

Every macro function in this package takes an explicit `doc` argument
(a com.sun.star.text.TextDocument) rather than reaching for a global. That
makes each function:
  - directly unit-testable from an external UNO connection (see
    lo-macros/tests/), and
  - usable from a thin "entry point" wrapper for LibreOffice's own
    Tools > Macros > Run Macro / toolbar-button invocation, which is the
    only place `XSCRIPTCONTEXT` actually exists.

See lo-macros/vistatype/entrypoints.py for the XSCRIPTCONTEXT wrappers that
get registered as the actual user-facing macros.
"""
from __future__ import annotations

import uno
from com.sun.star.style.BreakType import PAGE_BEFORE, NONE as BREAK_NONE
from com.sun.star.beans import PropertyValue


def get_active_document():
    """Only works when called from inside a script LibreOffice itself
    invoked (Tools > Macros, a toolbar button, etc). Not usable from an
    external UNO socket connection -- pass `doc` explicitly there instead.
    """
    try:
        desktop = XSCRIPTCONTEXT.getDesktop()  # noqa: F821
        return XSCRIPTCONTEXT.getDocument()  # noqa: F821
    except NameError as exc:
        raise RuntimeError(
            "get_active_document() only works when invoked by LibreOffice's "
            "own scripting provider (XSCRIPTCONTEXT is not defined here). "
            "Call the underlying function with an explicit `doc` argument "
            "instead."
        ) from exc


def make_prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def iter_paragraphs(doc):
    """Yield every top-level paragraph object (com.sun.star.text.Paragraph)
    in the document body, in order. Tables/frames are not descended into --
    callers that need those should walk doc.TextTables / doc.TextFrames
    separately.
    """
    enum = doc.Text.createEnumeration()
    while enum.hasMoreElements():
        el = enum.nextElement()
        if el.supportsService("com.sun.star.text.Paragraph"):
            yield el


def get_current_paragraph(doc):
    """The paragraph containing the view cursor (LibreOffice's caret), or
    None if there's no controller (e.g. a doc opened Hidden with no view).
    Mirrors the original macros' "place the cursor in the paragraph before
    running" convention.
    """
    controller = doc.CurrentController
    if controller is None:
        return None
    vc = controller.ViewCursor
    text = vc.Text
    cursor = text.createTextCursorByRange(vc.Start)
    cursor.gotoStartOfParagraph(False)
    cursor.gotoEndOfParagraph(True)
    # Walk paragraphs to find the one whose range matches -- UNO doesn't
    # give us the Paragraph object directly from a cursor, only a
    # TextRange, so we compare by cursor position.
    for para in iter_paragraphs(doc):
        para_cursor = text.createTextCursorByRange(para.Start)
        para_cursor.gotoEndOfParagraph(True)
        if (text.compareRegionEnds(para_cursor, cursor) == 0
                and text.compareRegionStarts(para_cursor, cursor) == 0):
            return para
    return None


def has_selection(doc):
    controller = doc.CurrentController
    if controller is None:
        return False
    vc = controller.ViewCursor
    return vc.getString() != ""


def get_selected_paragraphs(doc):
    """Paragraphs touched by the current selection, or -- matching the
    original macros' UX ("place the cursor in the paragraph; no need to
    select it") -- just the current paragraph if nothing is selected.
    """
    controller = doc.CurrentController
    if controller is None:
        return list(iter_paragraphs(doc))
    vc = controller.ViewCursor
    if vc.getString() == "":
        cur = get_current_paragraph(doc)
        return [cur] if cur else []

    text = vc.Text
    # IMPORTANT: compareRegionStarts/compareRegionEnds each compare the
    # SAME endpoint kind (start-vs-start, or end-vs-end) of two ranges --
    # there's no cross "end-of-A vs start-of-B" comparison in the UNO API.
    # To compare arbitrary points against each other we build *collapsed*
    # (point) cursors at exactly the position we care about, then always
    # compare start-vs-start.
    sel_start_pt = text.createTextCursorByRange(vc.Start)   # collapsed
    sel_end_pt = text.createTextCursorByRange(vc.End)       # collapsed

    matched = []
    for para in iter_paragraphs(doc):
        p_start_pt = text.createTextCursorByRange(para.Start)  # collapsed
        p_end_pt = text.createTextCursorByRange(para.Start)
        p_end_pt.gotoEndOfParagraph(False)  # False: move, don't select -> stays collapsed

        # NOTE on sign convention (verified empirically against a real LO
        # instance -- this is the opposite of typical compareTo()
        # semantics): compareRegionStarts(r1, r2) returns +1 if r1 is
        # POSITIONED BEFORE r2, -1 if r1 is AFTER r2, 0 if equal.
        ends_before_sel = text.compareRegionStarts(p_end_pt, sel_start_pt) > 0
        starts_after_sel = text.compareRegionStarts(p_start_pt, sel_end_pt) < 0
        if not ends_before_sel and not starts_after_sel:
            matched.append(para)

    if matched:
        return matched
    cur = get_current_paragraph(doc)
    return [cur] if cur else []


def set_break_before_page(paragraph, on: bool):
    paragraph.BreakType = PAGE_BEFORE if on else BREAK_NONE


def ensure_paragraph_style(doc, name: str, based_on: str = "Standard", props: dict | None = None):
    """Get-or-create a paragraph style by display name. Used so that macros
    like `format_page_numbers` work even on a document that never had
    LargePrintTemplate.dotx (or an equivalent .ott) attached -- the styles
    this package needs are created on demand with the same property values
    extracted from the real template, so the *visual result* matches even
    when the *style catalog* started out empty. If you HAVE attached the
    real template/.ott, this just returns the existing style untouched.
    """
    families = doc.StyleFamilies
    para_styles = families.getByName("ParagraphStyles")
    if para_styles.hasByName(name):
        return para_styles.getByName(name)

    style = doc.createInstance("com.sun.star.style.ParagraphStyle")
    para_styles.insertByName(name, style)
    if based_on and para_styles.hasByName(based_on):
        style.ParentStyle = based_on
    if props:
        for k, v in props.items():
            style.setPropertyValue(k, v)
    return style


def show_message_box(doc, message: str, title: str = "VistaType LP", buttons="ok"):
    """Interactive confirmation dialog -- used where the original macro
    pauses to ask the user something (e.g. "Found N images. Delete all?").
    `buttons` is "ok", "yesno", or "yesnocancel". Returns "OK"/"CANCEL"/
    "YES"/"NO".

    Requires a real window toolkit; will raise in a plain `--headless
    --invisible` soffice instance with no window server. Callers that need
    to run unattended (batch jobs, automated tests) should catch that and
    fall back to an explicit yes/no argument -- every function in this
    package that would otherwise pop this dialog accepts a `confirm`
    keyword for exactly that reason; pass True/False to skip the dialog
    entirely.
    """
    from com.sun.star.awt.MessageBoxButtons import (
        BUTTONS_OK, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL,
    )
    from com.sun.star.awt.MessageBoxType import QUERYBOX

    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    frame = doc.CurrentController.Frame if doc.CurrentController else None
    parent = frame.ContainerWindow if frame else None

    button_map = {
        "ok": BUTTONS_OK,
        "yesno": BUTTONS_YES_NO,
        "yesnocancel": BUTTONS_YES_NO_CANCEL,
    }
    box = toolkit.createMessageBox(
        parent, QUERYBOX, button_map.get(buttons, BUTTONS_OK), title, message
    )
    result = box.execute()
    names = {1: "OK", 2: "CANCEL", 3: "YES", 4: "NO"}
    return names.get(result, str(result))
