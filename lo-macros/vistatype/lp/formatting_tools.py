"""
vistatype.lp.formatting_tools
===============================
Port of most of the "Other Formatting Tools Group"
(VistaType LP User Guide p.47-51):

  Lp_Horz_List_To_Vertical             -> horizontal_list_to_vertical
  Lp_Compress_Linear_Math              -> compress_linear_math
  Lp_Toggle_Space_After_Current_Para   -> toggle_space_after_para
  Lp_Keep_With_Next_Para               -> toggle_keep_with_next_para
  Lp_Type_Fill_In_Line                 -> type_fill_in_line
  Lp_Format_Exercise_Lv_1_and_Lv_2      -> format_exercise_levels

Not covered here (see docs/COMMAND_MAP.md for why + status):
  Lp_Picture_Color_Change_Menu, Lp_Picture_Tools_Menu_Starter, Lp_Table_Tools
"""
from __future__ import annotations

import re

import uno

from .. import utils

PARA_BREAK = uno.getConstantByName("com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK")


# ---------------------------------------------------------------------------
# Horizontal List to Vertical (p.50-51)
# ---------------------------------------------------------------------------

def _match_marker(token: str):
    m = re.match(r"^([A-Za-z]{1,2}[\.\)])(.*)$", token)
    if not m:
        m = re.match(r"^(\d{1,2}[\.\)])(.*)$", token)
    if not m:
        return None, None
    return m.group(1), m.group(2)


def _split_ordered(text: str):
    """Same >= 3-marker heuristic as pandoc/filters/lp-horizontal-list-to-
    vertical.lua, operating on a plain string (LO gives us real strings, no
    AST token games needed).
    """
    tokens = text.split(" ")
    segments = []
    current_marker = None
    current_words = []

    def flush():
        if current_marker is not None:
            segments.append((current_marker, " ".join(current_words).strip()))

    for tok in tokens:
        marker, rest = _match_marker(tok)
        if marker:
            flush()
            current_marker = marker
            current_words = [rest] if rest else []
        elif current_marker is not None:
            current_words.append(tok)
    flush()
    return segments if len(segments) >= 3 else None


def _marker_sort_key(marker: str):
    core = marker[:-1]
    if core.isdigit():
        return int(core)
    return ord(core.lower()[0]) - ord("a") + 1


def horizontal_list_to_vertical(doc, mode="auto", target=None, sort_ascending=True):
    """Lp_Horz_List_To_Vertical.

    `target` is the paragraph to convert (defaults to the current
    paragraph, matching "place the cursor in the paragraph; no need to
    manually select it" from the guide -- for multi-paragraph selections,
    pass a list of paragraphs from utils.get_selected_paragraphs(doc)
    instead and this runs on each one that matches).

    `mode`:
      "ordered" - letters/numbers ("a.", "b)", "12.") mark each item;
                  requires >= 3 markers found, same false-positive
                  safeguard as the Pandoc filter
      "spaced"  - single words separated by one-or-more spaces, no markers
      "tabbed"  - words/word-groups separated by real tab characters
      "auto"    - try "ordered" first; only the interactive modes
                  ("spaced"/"tabbed") require you to say so explicitly,
                  since -- unlike a marker run -- word-spacing or tabs
                  alone aren't a safe auto-detect signal (an ordinary
                  paragraph of prose IS words separated by spaces).

    Returns the list of newly-created paragraphs, or None if nothing
    matched (the original paragraph is left untouched in that case).
    """
    paras = target if isinstance(target, list) else (
        [target] if target is not None else [utils.get_current_paragraph(doc)]
    )
    all_new_paras = []
    for para in paras:
        if para is None:
            continue
        text = para.getString()
        lines = None

        if mode in ("auto", "ordered"):
            segments = _split_ordered(text)
            if segments:
                if sort_ascending:
                    segments = sorted(segments, key=lambda s: _marker_sort_key(s[0]))
                lines = [f"{marker} {rest}".strip() for marker, rest in segments]

        if lines is None and mode == "spaced":
            words = [w for w in text.split(" ") if w != ""]
            lines = words

        if lines is None and mode == "tabbed":
            lines = [seg for seg in text.split("\t") if seg != ""]

        if lines is None and mode not in ("ordered", "spaced", "tabbed", "auto"):
            raise ValueError(f"unknown mode: {mode!r}")

        if lines is None:
            continue  # nothing matched confidently; leave this paragraph alone

        # replace this paragraph's text with the first line, then insert
        # the rest as new paragraphs immediately after
        text_obj = para.Text
        cursor = text_obj.createTextCursorByRange(para.Start)
        cursor.gotoEndOfParagraph(True)
        text_obj.insertString(cursor, lines[0], True)
        insert_point = text_obj.createTextCursorByRange(cursor.End)
        for line in lines[1:]:
            text_obj.insertControlCharacter(insert_point, PARA_BREAK, False)
            text_obj.insertString(insert_point, line, False)
        all_new_paras.append(para)

    return all_new_paras or None


# ---------------------------------------------------------------------------
# Compress Linear Math
# ---------------------------------------------------------------------------

_COMPARISON_SYMBOLS = "=\u2248<>\u2260\u2264\u2265"
_THIN_SPACE = "\u2009"


def compress_linear_math(doc, target=None):
    """Lp_Compress_Linear_Math.

    Removes all normal spaces from the target paragraph's text, then
    inserts a thin space (U+2009) immediately before and after every
    comparison symbol (=, \u2248, <, >, \u2260, \u2264, \u2265) -- exactly
    as the guide describes ("removes all normal spaces then places very
    thin (hair spaces) between the math and symbols of comparison").

    NOTE: this differs from the Braille/Dx_ side's version of the same-
    named macro, which instead *keeps* normal spaces around comparison
    operators and removes them everywhere else -- the two guides describe
    genuinely different rules for what is nominally "the same" macro name.
    See vistatype.dx.formatting_tools.compress_linear_math for that
    variant, and docs/COMMAND_MAP.md for the side-by-side.
    """
    para = target if target is not None else utils.get_current_paragraph(doc)
    if para is None:
        return None
    text = para.getString()
    no_spaces = text.replace(" ", "")
    out = []
    for ch in no_spaces:
        if ch in _COMPARISON_SYMBOLS:
            out.append(_THIN_SPACE + ch + _THIN_SPACE)
        else:
            out.append(ch)
    result = "".join(out)
    # collapse any doubled-up thin spaces where two comparison symbols
    # (or a symbol at a word boundary) ended up adjacent
    result = re.sub(_THIN_SPACE + "+", _THIN_SPACE, result)
    para.setString(result)
    return para


# ---------------------------------------------------------------------------
# Toggle Space After Para / Keep With Next Para
# ---------------------------------------------------------------------------

def toggle_space_after_para(doc, target=None):
    """Lp_Toggle_Space_After_Current_Para: closes the gap between a titling
    paragraph and the one below it by zeroing ParaBottomMargin; running it
    again restores the value defined by the paragraph's own style.
    """
    para = target if target is not None else utils.get_current_paragraph(doc)
    if para is None:
        return None
    if para.ParaBottomMargin == 0:
        families = doc.StyleFamilies.getByName("ParagraphStyles")
        style_name = para.ParaStyleName
        restored = 0
        if families.hasByName(style_name):
            restored = families.getByName(style_name).ParaBottomMargin
        para.ParaBottomMargin = restored or 353  # ~0.25in (VistaType Normal default) as last resort
        return False  # space is back "on"
    para.ParaBottomMargin = 0
    return True  # space is now "off" (removed)


def toggle_keep_with_next_para(doc, target=None):
    """Lp_Keep_With_Next_Para: toggles ParaKeepTogether, which keeps this
    paragraph glued to the one immediately following it across a page
    break. (Distinct from the shared Sh_Keep_Lines_Of_Para_Together, which
    toggles ParaSplit -- keeping *this paragraph's own lines* from
    splitting across a page, not gluing it to the next paragraph.)
    """
    para = target if target is not None else utils.get_current_paragraph(doc)
    if para is None:
        return None
    para.ParaKeepTogether = not para.ParaKeepTogether
    return para.ParaKeepTogether


# ---------------------------------------------------------------------------
# Type Fill-In Line (p.48-49)
# ---------------------------------------------------------------------------

FILL_IN_CHAR = "_"


def _insert_underlined_run(doc, cursor, length, leading_space=True):
    from com.sun.star.awt.FontUnderline import SINGLE as UNDERLINE_SINGLE
    text_obj = cursor.Text
    if leading_space:
        text_obj.insertString(cursor, " ", False)
    run_start = text_obj.createTextCursorByRange(cursor.End)
    text_obj.insertString(cursor, FILL_IN_CHAR * length, False)
    run_end = text_obj.createTextCursorByRange(cursor.End)
    run_range = text_obj.createTextCursorByRange(run_start.Start)
    run_range.gotoRange(run_end.End, True)
    run_range.CharUnderline = UNDERLINE_SINGLE


def type_fill_in_line(doc, length=8, to_right_margin=False):
    """Lp_Type_Fill_In_Line: inserts a fill-in line (underlined
    underscores) at the cursor. `length` is 1-21 for a fixed-length line
    (each begins with a leading space, per the guide); pass
    `to_right_margin=True` for a line that runs to the paragraph's right
    margin instead of a fixed length.

    Fixed-length implementation matches the guide directly. The
    to-right-margin variant is implemented differently from Word's
    (unknown, non-recoverable) internal approach: we add a right-aligned
    paragraph tab stop at the page's right text margin, with '_' as its
    fill character, and insert a real tab -- LibreOffice's native
    mechanism for "a leader-filled run to a tab stop", which achieves the
    same visual and functional result (a fill-in line that reaches the
    margin regardless of how much text precedes it on the line).
    """
    controller = doc.CurrentController
    if controller is None:
        raise RuntimeError("type_fill_in_line needs a document with a view (ViewCursor)")
    vc = controller.ViewCursor

    if to_right_margin:
        para = utils.get_current_paragraph(doc)
        page_style_name = para.PageDescName or "Standard"
        page_styles = doc.StyleFamilies.getByName("PageStyles")
        page_style = page_styles.getByName(page_style_name)
        text_area_width = page_style.Width - page_style.LeftMargin - page_style.RightMargin

        from com.sun.star.style import TabStop
        from com.sun.star.style.TabAlign import RIGHT as TAB_RIGHT
        stop = TabStop()
        stop.Position = text_area_width
        stop.Alignment = TAB_RIGHT
        stop.FillChar = FILL_IN_CHAR
        stop.DecimalChar = "."
        para.ParaTabStops = (stop,)

        text_obj = vc.Text
        text_obj.insertString(vc, "\t", False)
        return None

    if not (1 <= length <= 21):
        raise ValueError("length must be between 1 and 21 (guide: 21 fixed-length buttons)")
    _insert_underlined_run(doc, vc, length, leading_space=True)
    return None


# ---------------------------------------------------------------------------
# Format Exercise Levels 1 & 2 (p.47-48)
# ---------------------------------------------------------------------------

QUESTION_STYLE = "List"
ANSWER_STYLE = "List 2"

_QUESTION_START_RE = re.compile(r"^\s*\d+[\.\s]")
_TAB_OR_UNDERSCORE_RUN_RE = re.compile(r"(?:\t|(?<![_\w])_(?![_\w]))+")


def _ensure_style(doc, name, based_on="Standard"):
    families = doc.StyleFamilies.getByName("ParagraphStyles")
    if families.hasByName(name):
        return families.getByName(name)
    style = doc.createInstance("com.sun.star.style.ParagraphStyle")
    families.insertByName(name, style)
    if families.hasByName(based_on):
        style.ParentStyle = based_on
    return style


def _substitute_fill_ins(text: str) -> str:
    """Each run of tabs (or a lone underscore) becomes one "large print"
    8-underscore fill-in; multiple tabs/underscores in an unbroken run
    become multiple such fill-ins joined by commas -- matches the guide's
    OCR-cleanup behavior exactly (p.48-49). This only inserts the plain
    text; the caller applies underline formatting afterward (easier to do
    correctly on the live paragraph after the text is in its final place).
    """
    def repl(m):
        run = m.group(0)
        count = run.count("\t") + run.count("_")
        count = max(count, 1)
        return ",".join(["_" * 8] * count)
    return _TAB_OR_UNDERSCORE_RUN_RE.sub(repl, text)


def _underline_underscore_runs(para):
    from com.sun.star.awt.FontUnderline import SINGLE as UNDERLINE_SINGLE
    text_obj = para.Text
    full_text = para.getString()
    for m in re.finditer(r"_+", full_text):
        run_cursor = text_obj.createTextCursorByRange(para.Start)
        run_cursor.goRight(m.start(), False)
        run_cursor.goRight(m.end() - m.start(), True)
        run_cursor.CharUnderline = UNDERLINE_SINGLE


def format_exercise_levels(doc, target_paragraphs=None):
    """Lp_Format_Exercise_Lv_1_and_Lv_2.

    Operates on `target_paragraphs` (defaults to the current selection via
    utils.get_selected_paragraphs) -- matching "select the exercise(s)
    before executing the macro" from the guide. A paragraph starting with
    a number + period/space is a question (styled "List"); anything else
    in the selection is treated as an answer choice (styled "List 2").
    Tabs and lone underscores within the selection are replaced with
    underlined 8-underscore fill-in lines (comma-joined for runs of more
    than one), per the guide's OCR-cleanup behavior.

    Creates the "List"/"List 2" paragraph styles on the fly (based on
    Standard) if the document doesn't already have them.
    """
    _ensure_style(doc, QUESTION_STYLE)
    _ensure_style(doc, ANSWER_STYLE)

    paras = target_paragraphs if target_paragraphs is not None else utils.get_selected_paragraphs(doc)

    formatted = {"questions": 0, "answers": 0}
    for para in paras:
        text = para.getString()
        new_text = _substitute_fill_ins(text)
        if new_text != text:
            para.setString(new_text)

        if _QUESTION_START_RE.match(new_text):
            para.ParaStyleName = QUESTION_STYLE
            formatted["questions"] += 1
        else:
            para.ParaStyleName = ANSWER_STYLE
            formatted["answers"] += 1

        _underline_underscore_runs(para)

    return formatted
