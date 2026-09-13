"""
vistatype.shared
=================
Port of the "Sh_" (shared between Lp_ and Dx_) ribbon commands:

  Sh_Apply_Title_Case_Capitalization -> apply_title_case
  Sh_Doc_Info                        -> doc_info
  Sh_Keep_Lines_Of_Para_Together     -> toggle_keep_lines_together
  Sh_Move_Paragraph_To_Next_Page     -> toggle_move_to_next_page
  Sh_Show_Char_Val                   -> char_value_at_cursor

These operate on the real LibreOffice Writer object model (UNO), not a
Word compatibility shim -- e.g. "Doc Info" reports LibreOffice's own page
style properties rather than trying to fake Word's "attached template"
concept, which LO doesn't have in the same form.
"""
from __future__ import annotations

from com.sun.star.style.BreakType import PAGE_BEFORE, NONE as BREAK_NONE

from . import utils

# Words that stay lowercase in title case, unless they are the first or
# last word of the selection (standard title-case convention; matches the
# guide's "to, a, and, but, or, at, on, and by and many others").
_LOWERCASE_WORDS = {
    "a", "an", "and", "as", "at", "but", "by", "for", "if", "in", "nor",
    "of", "on", "or", "per", "so", "the", "to", "up", "vs", "via", "yet",
    "with", "from", "into", "onto", "than",
}


def apply_title_case(doc, target=None):
    """Sh_Apply_Title_Case_Capitalization.

    Capitalizes each word in the target paragraph, except the small words
    listed above -- which is as precisely as the guide states the rule
    ("Each word in the paragraph will be capitalized except for words
    like: to, a, and, but, or, at, on, and by and many others"; VistaType
    LP guide p.49). We add exactly one thing the guide doesn't spell out
    but that no reasonable implementation would skip: the very first word
    of the paragraph is always capitalized, even if it's on the small-
    words list, so a title never starts lowercase. We deliberately do NOT
    add further conventions some style guides use (capitalizing the last
    word, or the word after a colon) since those aren't evidenced by the
    guide and the real algorithm isn't recoverable from source (see
    docs/COMMAND_MAP.md) -- adding them would be guessing, not porting.
    """
    para = target if target is not None else utils.get_current_paragraph(doc)
    if para is None:
        return None
    text = para.getString()
    words = text.split(" ")
    new_words = []
    for i, w in enumerate(words):
        if w == "":
            new_words.append(w)
            continue
        # keep leading punctuation attached, title-case the alphabetic core
        lead = ""
        core = w
        while core and not core[0].isalpha():
            lead += core[0]
            core = core[1:]
        if not core:
            new_words.append(w)
            continue
        if i != 0 and core.lower() in _LOWERCASE_WORDS:
            new_core = core.lower()
        else:
            new_core = core[0].upper() + core[1:]
        new_words.append(lead + new_core)
    para.setString(" ".join(new_words))
    return para


def toggle_keep_lines_together(doc, target=None):
    """Sh_Keep_Lines_Of_Para_Together: toggles ParaSplit. When ParaSplit is
    False, LO will not split the paragraph's lines across a page boundary
    (matches Word's "keep lines together").
    """
    para = target if target is not None else utils.get_current_paragraph(doc)
    if para is None:
        return None
    para.ParaSplit = not para.ParaSplit
    # ParaSplit True == "splitting allowed" == NOT kept together.
    return not para.ParaSplit  # returns the new "kept together" state


def toggle_move_to_next_page(doc, target=None):
    """Sh_Move_Paragraph_To_Next_Page: toggles a page-break-before on the
    paragraph (BreakType.PAGE_BEFORE <-> NONE). The original macro always
    moves forward; unlike the true VBA macro (which is a one-shot "push to
    next page"), we treat it as a toggle so re-running it undoes the
    break -- more forgiving for a script-driven workflow. If you want the
    exact one-shot behavior, call with `toggle=False`.
    """
    para = target if target is not None else utils.get_current_paragraph(doc)
    if para is None:
        return None
    if para.BreakType == PAGE_BEFORE:
        para.BreakType = BREAK_NONE
        return False
    para.BreakType = PAGE_BEFORE
    return True


def char_value_at_cursor(doc):
    """Sh_Show_Char_Val: reports the Unicode code point of the character
    immediately to the right of the (view) cursor -- the LO equivalent of
    Word's "Show ANSI Character Value" find-and-replace helper. Word's
    ANSI/hex-code-page framing doesn't map cleanly onto LibreOffice (which
    is Unicode-native throughout, no legacy code-page layer to inspect),
    so this reports the values that actually matter for LO's own
    Find & Replace (which accepts \\uXXXX-style regex, not ANSI codes):
    the character itself, its decimal code point, and its hex code point.
    Returns None if the cursor is at the end of the text with no next
    character.
    """
    controller = doc.CurrentController
    if controller is None:
        return None
    vc = controller.ViewCursor
    text = vc.Text
    probe = text.createTextCursorByRange(vc.End)
    moved = probe.goRight(1, True)
    if not moved:
        return None
    ch = probe.getString()
    if not ch:
        return None
    cp = ord(ch[0])
    return {
        "char": ch[0],
        "decimal": cp,
        "hex": f"U+{cp:04X}",
    }


def doc_info(doc):
    """Sh_Doc_Info: reports page size/orientation/margins and the current
    paragraph's style -- the properties LibreOffice actually models. (LO
    documents don't carry a Word-style "AttachedTemplate" pointer once
    they're native .odt, so we report style-catalog membership instead --
    see `has_large_print_styles` below -- rather than fake a template path
    that wouldn't mean anything in LO.)
    """
    controller = doc.CurrentController
    page_style_name = "Standard"
    if controller is not None:
        para = utils.get_current_paragraph(doc)
        if para is not None and para.PageDescName:
            page_style_name = para.PageDescName

    page_styles = doc.StyleFamilies.getByName("PageStyles")
    page_style = (
        page_styles.getByName(page_style_name)
        if page_styles.hasByName(page_style_name)
        else page_styles.getByIndex(0)
    )

    width_1_100mm = page_style.Width
    height_1_100mm = page_style.Height

    def mm100_to_in(v):
        return round(v / 2540.0, 2)

    info = {
        "page_style": page_style_name,
        "width_in": mm100_to_in(width_1_100mm),
        "height_in": mm100_to_in(height_1_100mm),
        "orientation": "Landscape" if width_1_100mm > height_1_100mm else "Portrait",
        "top_margin_in": mm100_to_in(page_style.TopMargin),
        "bottom_margin_in": mm100_to_in(page_style.BottomMargin),
        "left_margin_in": mm100_to_in(page_style.LeftMargin),
        "right_margin_in": mm100_to_in(page_style.RightMargin),
    }
    if controller is not None:
        para = utils.get_current_paragraph(doc)
        info["current_paragraph_style"] = para.ParaStyleName if para else None
    # Report whether the accessibility styles this package relies on
    # (Print Pg Num, Para Black, etc.) are present, as a proxy for "has the
    # Large Print template's styles been applied to this document".
    para_styles = doc.StyleFamilies.getByName("ParagraphStyles")
    info["has_large_print_styles"] = para_styles.hasByName("Print Pg Num")
    return info
