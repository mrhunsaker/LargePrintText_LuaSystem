"""
vistatype.dx.reference_pages
==============================
Port of the Braille/BANA "Reference Page Number Formatting Group"
(BANA Macros for 2.0 User Guide p.16-20):

  Dx_AutoTag_Page_Numbers        -> auto_tag_page_numbers
  Dx_Manual_Tag_with_Dollar_pg   -> manual_tag_page_number
  Dx_Format_Tagged_Page_Numbers  -> format_tagged_page_numbers
  Dx_Ref_Pg_Number_Sequence_Menu -> validate_tagged_page_numbers

The pattern-matching for WHICH paragraphs count as reference page numbers
is identical to the large-print side (vistatype.lp.reference_pages) --
both guides describe the same detection rule. What's genuinely different
for Braille, and is what this module adds on top of the shared detection
logic, is DBT (Duxbury Braille Translator) markup:

  - continuation page ranges get a DBT continuation-page code embedded
    after the range, followed by the page number that continues onto the
    next braille page (guide's own example: "$pg13-15" + the DBT
    continuation code + "15")
  - lowercase roman numerals need a DBT letter-sign code so DBT doesn't
    misread them as ordinary letters -- UEB only needs this for "c", "l",
    "v" specifically (single letters that are also valid lowercase roman
    numerals and could be ambiguous); EBAE needs it for ALL lowercase
    roman numerals

This module does NOT have a real BANA Braille Template to draw the
"BANA Braille Template Reference Page Number" paragraph style from (only
LargePrintTemplate.dotx was provided -- see docs/COMMAND_MAP.md). The
style is created on the fly with a reasonable placeholder appearance
(same shape as the LP side's fallback) rather than the real BANA-specific
formatting, which is not recoverable without that template file.
"""
from __future__ import annotations

import re

from .. import utils
from ..lp.reference_pages import _classify_page_number, _roman_to_int  # noqa: F401 (reused)

# DBT bracket codes, written out in full here (rather than split across a
# comment, which breaks Lua's long-comment parsing elsewhere in this repo
# -- Python doesn't have that problem, but keeping the same literal style
# for consistency).
DBT_CONTINUATION_CODE = "[[*lec*]]"
DBT_LETTER_SIGN_CODE = "[[*ii*]]"

_UEB_AMBIGUOUS_LOWER_ROMAN = {"c", "l", "v"}


def _tag_value(value: str, translation: str) -> str:
    """Apply the DBT codes described above to an already-classified page
    number value (no "$pg" prefix yet -- that's added by the caller).
    """
    kind, _ = _classify_page_number(value)

    if kind == "range":
        # "13-15" -> "13-15" + continuation code + "15" (the page that
        # continues onto the next braille page is the range's end value)
        start, end = value.split("-", 1)
        return f"{value}{DBT_CONTINUATION_CODE}{end}"

    if kind == "roman" and value.islower():
        if translation.upper() == "UEB":
            if value.lower() in _UEB_AMBIGUOUS_LOWER_ROMAN:
                return f"{value}{DBT_LETTER_SIGN_CODE}"
            return value
        # EBAE: every lowercase roman numeral gets the letter sign
        return f"{value}{DBT_LETTER_SIGN_CODE}"

    return value


def auto_tag_page_numbers(doc, translation: str = "UEB"):
    """Dx_AutoTag_Page_Numbers. `translation` is "UEB" or "EBAE" -- see
    module docstring for how that changes lowercase-roman-numeral
    handling. (The BANA template records this choice once per document
    when Dx_Attach_BANA_Template runs; since we don't have that template
    to attach, callers here just pass it explicitly.)
    """
    tagged = 0
    for para in utils.iter_paragraphs(doc):
        text = para.getString()
        if text.startswith("$pg"):
            continue
        kind, value = _classify_page_number(text)
        if kind:
            para.setString(f"$pg{_tag_value(value, translation)}")
            tagged += 1
    return tagged


def manual_tag_page_number(doc, target=None, translation: str = "UEB"):
    """Dx_Manual_Tag_with_Dollar_pg."""
    para = target if target is not None else utils.get_current_paragraph(doc)
    if para is None:
        return None
    text = para.getString()
    if not text.startswith("$pg"):
        kind, value = _classify_page_number(text)
        tagged_value = _tag_value(value, translation) if kind else text
        para.setString(f"$pg{tagged_value}")
    return para


BANA_REF_PG_STYLE = "BANA Braille Template Reference Page Number"

# Placeholder appearance only -- see module docstring. Distinct from the
# LP side's fallback fill colour so it's visually obvious in a mixed
# workflow that this is the Braille-side placeholder, not the real
# large-print "Print Pg Num" bar.
_BANA_REF_PG_FALLBACK_PROPS = {
    "ParaBackColor": 0xD9D9D9,
}


def _ensure_bana_ref_pg_style(doc):
    families = doc.StyleFamilies.getByName("ParagraphStyles")
    if families.hasByName(BANA_REF_PG_STYLE):
        return
    style = doc.createInstance("com.sun.star.style.ParagraphStyle")
    families.insertByName(BANA_REF_PG_STYLE, style)
    style.ParentStyle = "Standard" if families.hasByName("Standard") else "Default Paragraph Style"
    for k, v in _BANA_REF_PG_FALLBACK_PROPS.items():
        try:
            style.setPropertyValue(k, v)
        except Exception:
            pass


def format_tagged_page_numbers(doc):
    """Dx_Format_Tagged_Page_Numbers: same idea as the LP side's
    format_tagged_page_numbers -- strip "$pg" (and leave any DBT codes
    already embedded in the tag, e.g. the continuation code, intact as
    literal text, since DBT reads those directly from the Word/LO
    document text) -- styled with the BANA reference-page-number style
    (created on the fly; see module docstring for the caveat about not
    having the real BANA template to draw it from).
    """
    _ensure_bana_ref_pg_style(doc)
    formatted = 0
    for para in utils.iter_paragraphs(doc):
        text = para.getString()
        if text.startswith("$pg"):
            para.setString(text[len("$pg"):])
            para.ParaStyleName = BANA_REF_PG_STYLE
            formatted += 1
    return formatted


def validate_tagged_page_numbers(doc):
    """Dx_Ref_Pg_Number_Sequence_Menu (the "Validate $pg Tags" workflow).
    Same shape as the LP side's version -- returns a list of dicts rather
    than opening a UI, so the caller decides how to present it.
    """
    entries = []
    idx = 0
    for para in utils.iter_paragraphs(doc):
        text = para.getString()
        if not text.startswith("$pg"):
            continue
        idx += 1
        value = text[len("$pg"):]
        entries.append({"index": idx, "text": text, "value": value})
    return entries
