"""
vistatype.lp.reference_pages
==============================
Port of the "Reference Page Number Formatting Group"
(VistaType LP User Guide p.47-49):

  Lp_AutoTag_Page_Numbers      -> auto_tag_page_numbers
  Lp_Manual_Tag_with_Dollar_pg -> manual_tag_page_number
  Lp_Format_Page_Numbers       -> format_tagged_page_numbers
  Lp_Validate_Dollar_PG        -> validate_tagged_page_numbers

A "reference page number" is a page number from the *original print* book,
preserved in the large-print transcription so a reader/teacher can
cross-reference pages between the two -- not a page number of the large
print document itself.

The pattern-matching rules here are identical to
pandoc/filters/lp-tag-reference-page-numbers.lua (see that file for the
detailed rule commentary) -- both were written from the same guide section,
so a paragraph that gets tagged by the batch Pandoc pipeline will get
tagged the same way here.
"""
from __future__ import annotations

import re

from .. import utils

_ROMAN_RE = re.compile(r"^[MCDXLVImcdxlvi]+$")


def _is_roman_numeral(s: str) -> bool:
    if not (1 <= len(s) <= 8):
        return False
    return bool(_ROMAN_RE.match(s))


def _classify_page_number(text: str):
    text = text.strip()
    m = re.match(r"^([A-Za-z]?\d+[A-Za-z]?)-([A-Za-z]?\d+[A-Za-z]?)$", text)
    if m:
        return "range", text
    if re.match(r"^[A-Za-z]?\d+[A-Za-z]?$", text):
        return "arabic", text
    if _is_roman_numeral(text):
        return "roman", text
    return None, None


def auto_tag_page_numbers(doc):
    """Lp_AutoTag_Page_Numbers: scans every paragraph; any paragraph whose
    ENTIRE text is (and only is) a recognizable reference page number gets
    "$pg" prepended. A paragraph embedded in running prose, or with a page
    number only at its start/end, is left alone -- matching "the page
    numbers must appear in a paragraph of their own" from the guide.

    Returns the count of paragraphs tagged.
    """
    tagged = 0
    for para in utils.iter_paragraphs(doc):
        text = para.getString()
        if text.startswith("$pg"):
            continue  # already tagged
        kind, value = _classify_page_number(text)
        if kind:
            para.setString(f"$pg{value}")
            tagged += 1
    return tagged


def manual_tag_page_number(doc, target=None):
    """Lp_Manual_Tag_with_Dollar_pg: tags just the current paragraph (or
    `target`, if given) with "$pg", regardless of whether it matches the
    auto-detect pattern -- for reference numbers the automatic pass missed
    or that don't fit the recognized shapes. Place the cursor in the
    paragraph first, matching the original.
    """
    para = target if target is not None else utils.get_current_paragraph(doc)
    if para is None:
        return None
    text = para.getString()
    if not text.startswith("$pg"):
        para.setString(f"$pg{text}")
    return para


PRINT_PG_NUM_STYLE = "Print Pg Num"

# Fallback properties if the style doesn't already exist in the document's
# catalog (i.e. attach_lp_template() wasn't run first) -- copied from the
# real LargePrintTemplate.dotx's PrintPgNum style definition (top/bottom
# border, pink shading, right tab stop; see docs/COMMAND_MAP.md).
_PRINT_PG_NUM_FALLBACK_PROPS = {
    "ParaBackColor": 0xF4A3D4,
    "ParaTopMargin": 0,
}


def _ensure_print_pg_num_style(doc):
    families = doc.StyleFamilies
    para_styles = families.getByName("ParagraphStyles")
    if para_styles.hasByName(PRINT_PG_NUM_STYLE):
        return
    style = doc.createInstance("com.sun.star.style.ParagraphStyle")
    para_styles.insertByName(PRINT_PG_NUM_STYLE, style)
    style.ParentStyle = "Standard" if para_styles.hasByName("Standard") else "Default Paragraph Style"
    for k, v in _PRINT_PG_NUM_FALLBACK_PROPS.items():
        try:
            style.setPropertyValue(k, v)
        except Exception:
            pass


def format_tagged_page_numbers(doc):
    """Lp_Format_Page_Numbers: finds every paragraph tagged with "$pg",
    applies the "Print Pg Num" paragraph style, and removes the tag text
    itself, leaving just the bare reference page number -- e.g. "$pg12"
    becomes a paragraph containing "12" styled as Print Pg Num.

    Creates the Print Pg Num style on the fly (with the real template's
    property values) if the document doesn't already have it -- see
    _ensure_print_pg_num_style -- so this works even without having run
    attach_lp_template() first.
    """
    _ensure_print_pg_num_style(doc)
    formatted = 0
    for para in utils.iter_paragraphs(doc):
        text = para.getString()
        if text.startswith("$pg"):
            para.setString(text[len("$pg"):])
            para.ParaStyleName = PRINT_PG_NUM_STYLE
            formatted += 1
    return formatted


def validate_tagged_page_numbers(doc):
    """Lp_Validate_Dollar_PG: the original opens a temporary document
    listing every tagged reference page number in a column, so the
    transcriber can visually scan for missing numbers, out-of-order
    numbers, or false positives, then Alt-Tab back to the real document to
    fix anything. We reproduce that same "put them in a scratch document"
    workflow (rather than, say, a dialog box, since the whole point is to
    let the user see the surrounding text as they scan) -- but return the
    data as a plain Python list here, so the CALLER decides how to present
    it (a real new Writer window, a printed report, a test assertion,
    etc). See lo-macros/vistatype/entrypoints.py for the wrapper that
    actually opens a scratch window when invoked as a real LO macro.

    Returns a list of dicts: {"index": i, "text": paragraph_text,
    "in_order": bool, "sort_key": ...} where `in_order` flags any entry
    whose page value is not >= the previous tagged entry's value (a good
    signal for "this might be a false positive or an OCR misread").
    """
    entries = []
    prev_key = None
    idx = 0
    for para in utils.iter_paragraphs(doc):
        text = para.getString()
        if not text.startswith("$pg"):
            continue
        idx += 1
        value = text[len("$pg"):]
        kind, _ = _classify_page_number(value)
        sort_key = _sort_key_for(value, kind)
        in_order = (prev_key is None) or (sort_key is None) or (sort_key >= prev_key)
        entries.append({
            "index": idx,
            "text": text,
            "value": value,
            "kind": kind,
            "in_order": in_order,
        })
        if sort_key is not None:
            prev_key = sort_key
    return entries


def _sort_key_for(value, kind):
    if kind == "arabic":
        m = re.search(r"\d+", value)
        return int(m.group(0)) if m else None
    if kind == "range":
        m = re.search(r"\d+", value)
        return int(m.group(0)) if m else None
    if kind == "roman":
        return _roman_to_int(value.upper())
    return None


_ROMAN_VALUES = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}


def _roman_to_int(s: str) -> int:
    total = 0
    prev = 0
    for ch in reversed(s):
        val = _ROMAN_VALUES.get(ch, 0)
        if val < prev:
            total -= val
        else:
            total += val
            prev = val
    return total
