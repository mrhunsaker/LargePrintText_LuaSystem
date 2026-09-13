"""
vistatype.dx.file_cleanup
============================
Port of the Braille/BANA-side "File Cleanup Group" (BANA Macros for 2.0
User Guide p.13-15) plus "Format a Spelling List" (p.21):

  Dx_File_Fix_Sequence     -> fix_common_file_errors, delete_multiple_paragraph_marks
                              (delete_multiple_paragraph_marks is identical
                              on both sides -- re-exported from lp.file_cleanup)
  Dx_Selected_File_CleanUp -> selected_cleanup
  Dx_Spelling_List         -> format_spelling_list

KEY DIFFERENCE FROM THE LARGE-PRINT SIDE
==========================================
The BANA guide's "Fix Common File Errors" is a near-identical list to the
VistaType LP one, with one explicit, important exception: "Removes spaces
before and after Em dashes and hyphens (**for EBAE only**)" (p.13). The
Large Print guide has no such qualifier -- that fix always applies there.
So this module's dash/space handling takes a `translation` argument and
only strips spaces around dashes when translation="EBAE"; for "UEB" the
dash-run collapsing still happens, but surrounding spaces are left alone.
"""
from __future__ import annotations

import re

from com.sun.star.style.CaseMap import NONE as CASEMAP_NONE

from .. import utils
from ..lp.file_cleanup import delete_multiple_paragraph_marks  # noqa: F401 re-export (identical on both sides)

_DASH_RUN_RE = re.compile(r"([\u2013\u2014-])\1+")
_MULTI_SPACE_RE = re.compile(r" {2,}")
_PRIME_AFTER_DIGIT_RE = re.compile(r"(\d)'")
_DOUBLE_PRIME_AFTER_DIGIT_RE = re.compile(r'(\d)"')
_SPACE_AROUND_DASH_RE = re.compile(r" ?([\u2013\u2014-]) ?")


def _fix_text(s: str, translation: str) -> str:
    def collapse_dash_run(m):
        return m.group(1)
    s = _DASH_RUN_RE.sub(collapse_dash_run, s)

    if translation.upper() == "EBAE":
        def strip_space_around_dash(m):
            return m.group(1)
        s = _SPACE_AROUND_DASH_RE.sub(strip_space_around_dash, s)

    s = _MULTI_SPACE_RE.sub(" ", s)
    s = _PRIME_AFTER_DIGIT_RE.sub("\\1\u2032", s)
    s = _DOUBLE_PRIME_AFTER_DIGIT_RE.sub("\\1\u2033", s)
    s = s.strip(" \t")
    return s


def fix_common_file_errors(doc, translation: str = "UEB", remove_images=None, on_confirm=None):
    """Dx_File_Fix_Sequence step 1. See module docstring for the
    UEB/EBAE dash-space difference; otherwise the same class of fixes as
    the LP side (small caps -> all caps, drop caps reset, text-frame
    un-boxing, confirmable image removal) -- see
    vistatype.lp.file_cleanup.fix_common_file_errors's docstring for the
    full rundown of what is/isn't covered at this layer.
    """
    for para in utils.iter_paragraphs(doc):
        original = para.getString()
        fixed = _fix_text(original, translation)
        if fixed != original:
            para.setString(fixed)

        portion_enum = para.createEnumeration()
        while portion_enum.hasMoreElements():
            portion = portion_enum.nextElement()
            if getattr(portion, "CharCaseMap", CASEMAP_NONE) != CASEMAP_NONE:
                portion.setString(portion.getString().upper())
                portion.CharCaseMap = CASEMAP_NONE

        if para.DropCapFormat.Count:
            fmt = para.DropCapFormat
            fmt.Count = 0
            fmt.Lines = 0
            para.DropCapFormat = fmt
            para.DropCapWholeWord = False

    frame_names = list(doc.TextFrames.ElementNames)
    if frame_names:
        import uno
        para_break = uno.getConstantByName("com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK")
        end_cursor = doc.Text.createTextCursorByRange(doc.Text.End)
        for name in frame_names:
            frame = doc.TextFrames.getByName(name)
            frame_text = frame.getString()
            if frame_text.strip():
                doc.Text.insertControlCharacter(end_cursor, para_break, False)
                doc.Text.insertString(end_cursor, frame_text, False)
            doc.Text.removeTextContent(frame)

    graphics = list(doc.GraphicObjects.ElementNames) if doc.GraphicObjects else []
    if graphics:
        if remove_images is None:
            if on_confirm is not None:
                remove_images = on_confirm(len(graphics))
            else:
                answer = utils.show_message_box(
                    doc,
                    f"Found {len(graphics)} image(s) in the document. Delete all?",
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


def selected_cleanup(doc, translation: str = "UEB", options=None):
    """Dx_Selected_File_CleanUp: same fixes as fix_common_file_errors,
    scoped to the current selection (or current paragraph).
    """
    options = options or {"dashes_spaces_primes", "small_caps"}
    paras = utils.get_selected_paragraphs(doc)

    if "dashes_spaces_primes" in options:
        for para in paras:
            original = para.getString()
            fixed = _fix_text(original, translation)
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

    return {"paragraphs_processed": len(paras)}


# ---------------------------------------------------------------------------
# Format a Spelling List (BANA guide p.21)
# ---------------------------------------------------------------------------

def format_spelling_list(doc, target_paragraphs=None, spaces_between: int = 2):
    """Dx_Spelling_List.

    Spelling lists are presented first in contracted braille form,
    followed by the same word in uncontracted form, so the student sees
    both -- DBT itself decides contraction (not us).

    IMPORTANT HONESTY NOTE: the guide (p.21) describes this behavior but
    never states the literal DBT bracket code that forces a span of text
    to translate uncontracted (unlike $pg, the continuation code, the
    letter-sign code, and the fraction codes, which the guide DOES spell
    out verbatim and which the other functions in this package use
    directly). Guessing at DBT markup syntax here would risk silently
    producing wrong braille for a blind reader, which is a real-world
    harm, not just a cosmetic bug -- so we deliberately do NOT invent one.

    What this function actually does is the mechanical, verifiably-
    correct part only: duplicate each line as "<line><spaces><line>" so
    DBT's normal (contracted) translation applies to the first copy. It
    does NOT mark the second copy as uncontracted. If you know DBT's
    actual uncontracted-braille bracket code (check the DBT/BANA template
    documentation, or a real .dotm with working source), tell me and I'll
    wire it into the second copy precisely -- or apply DBT's "translate
    uncontracted" formatting/style to the second copy yourself in
    LibreOffice after running this.
    """
    paras = target_paragraphs if target_paragraphs is not None else utils.get_selected_paragraphs(doc)
    sep = " " * max(1, min(2, spaces_between))

    count = 0
    for para in paras:
        text = para.getString()
        if not text.strip():
            continue
        para.setString(f"{text}{sep}{text}")
        count += 1
    return count
