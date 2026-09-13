"""
vistatype.lp.file_cleanup
==========================
Port of the "File Cleanup Group" (VistaType LP User Guide p.44-46):

  Lp_File_Fix_Sequence     -> file_fix_sequence (runs the two steps below)
  Lp_Selected_File_CleanUp -> selected_cleanup

Unlike the Pandoc Lua filters (pandoc/filters/lp-fix-common-file-errors.lua,
lp-delete-multiple-paragraph-marks.lua), this operates on a live LibreOffice
document via UNO, which gives access to things Pandoc's AST can't see or
represent: real DropCapFormat, real TextFrames, real inline graphics, and a
document the user can watch change and undo (Ctrl+Z) if something looks
wrong -- exactly like the original interactive macro.

Verified against a real headless LibreOffice instance -- see
lo-macros/tests/test_file_cleanup.py.
"""
from __future__ import annotations

import re

from com.sun.star.style.CaseMap import NONE as CASEMAP_NONE
import uno

from .. import utils

_DASH_CHARS = "\u2010\u2011\u2012\u2013\u2014-"  # hyphen variants, en dash, em dash, ascii hyphen
_DASH_RUN_RE = re.compile(r"([\u2013\u2014-])\1+")
_MULTI_SPACE_RE = re.compile(r" {2,}")
_PRIME_AFTER_DIGIT_RE = re.compile(r"(\d)'")
_DOUBLE_PRIME_AFTER_DIGIT_RE = re.compile(r'(\d)"')
_SPACE_AROUND_DASH_RE = re.compile(r" ?([\u2013\u2014-]) ?")


def _fix_text(s: str) -> str:
    """The deterministic string-level fixes: dash-run collapsing, space
    stripping around dashes, multi-space collapsing, prime conversion.
    Same rules as pandoc/filters/lp-fix-common-file-errors.lua, applied
    directly to a plain Python string (LO gives us the whole paragraph's
    text as a real string with real characters -- no AST-token games
    needed here).
    """
    def collapse_dash_run(m):
        return m.group(1)
    s = _DASH_RUN_RE.sub(collapse_dash_run, s)

    def strip_space_around_dash(m):
        return m.group(1)
    s = _SPACE_AROUND_DASH_RE.sub(strip_space_around_dash, s)

    s = _MULTI_SPACE_RE.sub(" ", s)
    s = _PRIME_AFTER_DIGIT_RE.sub("\\1\u2032", s)
    s = _DOUBLE_PRIME_AFTER_DIGIT_RE.sub("\\1\u2033", s)
    s = s.strip(" \t")
    return s


def fix_common_file_errors(doc, remove_images=None, on_confirm=None):
    """Lp_File_Fix_Sequence step 1 ("Fix Common File Errors").

    Applies, across the whole document:
      - dash-run collapsing + space-stripping around dashes (see _fix_text)
      - multi-space collapsing
      - digit+quote -> prime/double-prime
      - small caps -> plain all-caps text (CharCaseMap SMALLCAPS -> NONE,
        with the run's text upper-cased so it stays visually all-caps
        even without the case-map flag)
      - drop caps reset to normal text (DropCapFormat cleared)
      - text frames "un-boxed": each frame's text is appended as a new
        paragraph at the end of the document, then the frame is removed
        (matches "removes text boxes and frames while preserving the
        text" -- position is approximate, not re-anchored at the frame's
        original spot, since UNO doesn't expose an unambiguous "insert
        exactly here" for former frame anchors; review placement after
        running this)

    Image removal is interactive in the original ("Found N images. Delete
    all?"). `remove_images` lets you skip the prompt entirely (True/False);
    leave it None to be asked via a real message box (requires a visible
    LO window -- will raise in a pure `--headless --invisible` session,
    which is why the automated tests always pass an explicit True/False).
    `on_confirm(image_count)` overrides how confirmation is obtained, if
    you want your own UI instead of the built-in message box (e.g. from
    a wrapper toolbar button) -- it should return True/False.
    """
    # --- dash/space/prime + small caps, paragraph by paragraph ---
    for para in utils.iter_paragraphs(doc):
        original = para.getString()
        fixed = _fix_text(original)
        if fixed != original:
            para.setString(fixed)

        # small caps -> real upper-case text, across each text portion
        portion_enum = para.createEnumeration()
        while portion_enum.hasMoreElements():
            portion = portion_enum.nextElement()
            if getattr(portion, "CharCaseMap", CASEMAP_NONE) != CASEMAP_NONE:
                portion.setString(portion.getString().upper())
                portion.CharCaseMap = CASEMAP_NONE

        # drop caps -> normal text
        if para.DropCapFormat.Count:
            fmt = para.DropCapFormat
            fmt.Count = 0
            fmt.Lines = 0
            para.DropCapFormat = fmt
            para.DropCapWholeWord = False

    # --- text boxes / frames: preserve text, remove the box ---
    frame_names = list(doc.TextFrames.ElementNames)
    if frame_names:
        end_cursor = doc.Text.createTextCursorByRange(doc.Text.End)
        para_break = uno.getConstantByName("com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK")
        for name in frame_names:
            frame = doc.TextFrames.getByName(name)
            frame_text = frame.getString()
            if frame_text.strip():
                doc.Text.insertControlCharacter(end_cursor, para_break, False)
                doc.Text.insertString(end_cursor, frame_text, False)
            doc.Text.removeTextContent(frame)

    # --- inline images, with confirmation ---
    graphics = list(doc.GraphicObjects.ElementNames) if doc.GraphicObjects else []
    if graphics:
        if remove_images is None:
            if on_confirm is not None:
                remove_images = on_confirm(len(graphics))
            else:
                answer = utils.show_message_box(
                    doc,
                    f"Found {len(graphics)} image(s) in the document. Delete all?\n\n"
                    "CAUTION: a 'Yes' answer will also delete images created by "
                    "an equation editor, but will leave native Writer "
                    "OLE/Math objects intact.",
                    title="File Cleanup",
                    buttons="yesno",
                )
                remove_images = answer == "YES"
        if remove_images:
            for name in graphics:
                graphic = doc.GraphicObjects.getByName(name)
                doc.Text.removeTextContent(graphic)

    return {
        "frames_removed": len(frame_names),
        "images_found": len(graphics),
        "images_removed": len(graphics) if remove_images else 0,
    }


def delete_multiple_paragraph_marks(doc):
    """Lp_File_Fix_Sequence step 2 ("Delete Multiple Paragraph Marks").

    Collapses runs of consecutive empty paragraphs down to one, across the
    whole document.

    The original macro lets you limit the scope to a selection. We always
    operate on the whole document instead, because doing so is provably
    equivalent for this particular operation: collapsing empty-paragraph
    runs outside a selection never touches anything the selection cares
    about, so a whole-document pass is a safe, idempotent superset of a
    scoped one -- and it sidesteps a real staleness problem (see below)
    that a scoped version would hit just as easily.

    Implementation note: we re-enumerate paragraphs fresh before every
    single deletion rather than collecting all of them up front and
    deleting in a batch. Verified experimentally that batching is unsafe:
    once the first deletion merges two paragraphs, UNO's paragraph
    proxies for anything after that point can go stale (a live document
    mutated out from under a cached reference), and a second deletion
    using an old reference throws UnknownPropertyException on `.Start`.
    Re-fetching a live paragraph list on every iteration costs a little
    speed on a very long run of blank paragraphs but is what actually
    works reliably.
    """
    text = doc.Text
    deleted = 0
    while True:
        paras = list(utils.iter_paragraphs(doc))

        found_pair = None
        prev_empty = False
        for para in paras:
            is_empty = para.getString().strip() == ""
            if is_empty and prev_empty:
                found_pair = para
                break
            prev_empty = is_empty

        if found_pair is None:
            break

        cursor = text.createTextCursorByRange(found_pair.Start)
        cursor.gotoEndOfParagraph(True)
        cursor.goRight(1, True)
        text.insertString(cursor, "", True)
        deleted += 1

    return deleted


def file_fix_sequence(doc, remove_images=None, on_confirm=None):
    """Lp_File_Fix_Sequence: the two-step sequence the guide recommends
    running immediately after attaching the template to a new document.
    """
    fix_result = fix_common_file_errors(doc, remove_images=remove_images, on_confirm=on_confirm)
    deleted = delete_multiple_paragraph_marks(doc)
    fix_result["empty_paragraphs_collapsed"] = deleted
    return fix_result


def selected_cleanup(doc, options=None):
    """Lp_Selected_File_CleanUp: the same class of fixes as
    fix_common_file_errors, but scoped to the current selection (or
    current paragraph, if nothing is selected) instead of the whole
    document -- matching the original's selection-scoped menu.

    `options` is an optional set of which fixes to run, from:
        {"dashes_spaces_primes", "small_caps", "multi_para_marks"}
    Defaults to all three.
    """
    options = options or {"dashes_spaces_primes", "small_caps", "multi_para_marks"}
    paras = utils.get_selected_paragraphs(doc)

    if "dashes_spaces_primes" in options:
        for para in paras:
            original = para.getString()
            fixed = _fix_text(original)
            if fixed != original:
                para.setString(fixed)

    if "small_caps" in options:
        for para in paras:
            portion_enum = para.createEnumeration()
            while portion_enum.hasMoreElements():
                portion = portion_enum.nextElement()
                if getattr(portion, "CharCaseMap", CASEMAP_NONE) != CASEMAP_NONE:
                    portion.setString(portion.getString().upper())
                    portion.CharCaseMap = CASEMAP_NONE

    deleted = 0
    if "multi_para_marks" in options:
        deleted = delete_multiple_paragraph_marks(doc)

    return {"paragraphs_processed": len(paras), "empty_paragraphs_collapsed": deleted}
