"""
vistatype.config
==================
Port of the "MS_" ribbon/QAT commands (VistaType LP User Guide p.6-7,
"Guide to the Quick Access Toolbar"):

  MS_Change_Word_Configuration_to_New_Install  -> set_config_new_install
  MS_Change_Word_Configuration_to_Large_Print  -> set_config_large_print
  MS_Change_Word_Configuration_to_Braille      -> set_config_braille

WHAT THESE ACTUALLY CHANGE, AND WHY THAT'S DIFFERENT IN LIBREOFFICE
=====================================================================
The original three QAT icons (House / Anchor / Magnifying Glass) switch a
bundle of *Word application-level* editing options: AutoCorrect behavior,
AutoFormat-as-you-type rules, and view settings, tuned differently for
plain editing vs. producing braille vs. producing large print.

LibreOffice doesn't expose an equivalent single-call "load this whole
editing-options profile" API, and many of the individual settings this
would touch (AutoCorrect replacement tables, "use smart quotes",
autocapitalize) are genuinely APPLICATION-wide preferences in LO
(configured once for the whole LO install via
Tools > AutoCorrect Options / Tools > Options), not per-document -- so
"switch to a different profile per document" isn't a natural operation the
way it is in Word.

What we CAN faithfully change per-document, and what the guide's own
description of the three configurations actually amounts to once you set
aside the AutoCorrect-table differences: formatting-mark visibility and
which paragraph style newly-typed text defaults to. That's what these
three functions do. If you also want application-wide AutoCorrect/
AutoFormat behavior changed, that's a one-time manual Tools > AutoCorrect
Options setup in LO, not something a per-document macro should be
silently doing on your behalf -- seev docs/COMMAND_MAP.md for the detailed
gap analysis.
"""
from __future__ import annotations

from . import utils


def set_formatting_marks_visible(doc, visible: bool = True):
    """Toggles View > Formatting Marks (spaces, tabs, paragraph marks,
    etc). The guide recommends this ON while editing for braille/large
    print work, matching Word's own "Show/Hide ¶" default recommendation.
    """
    controller = doc.CurrentController
    if controller is None:
        return None
    controller.ViewSettings.ShowNonprintingCharacters = visible
    return visible


def set_default_paragraph_style(doc, style_name: str):
    """Sets which paragraph style the document's own default/first
    paragraph uses, and -- if present -- the document's Default Paragraph
    Style entry, so *newly typed* paragraphs pick it up. Used to point new
    content at "Standard" (Normal) once accessibility styles are loaded.
    """
    para_styles = doc.StyleFamilies.getByName("ParagraphStyles")
    if not para_styles.hasByName(style_name):
        return False
    for para in utils.iter_paragraphs(doc):
        if para.getString() == "" and para.ParaStyleName in ("Standard", "Default Paragraph Style"):
            para.ParaStyleName = style_name
        break  # only worth nudging the very first paragraph if it's still blank
    return True


def set_config_new_install(doc):
    """MS_Change_Word_Configuration_to_New_Install: plain-editing profile
    -- formatting marks off, no special default style assumptions.
    """
    set_formatting_marks_visible(doc, False)
    return {"formatting_marks": False}


def set_config_large_print(doc):
    """MS_Change_Word_Configuration_to_Large_Print: recommended editing
    view for large-print work -- formatting marks on (so tabs/fill-in
    underscores/multiple spaces are visible while formatting exercises),
    default paragraph style "Standard" once attach_lp_template() has run.
    """
    set_formatting_marks_visible(doc, True)
    set_default_paragraph_style(doc, "Standard")
    return {"formatting_marks": True, "default_style": "Standard"}


def set_config_braille(doc):
    """MS_Change_Word_Configuration_to_Braille: recommended editing view
    for braille-prep work -- formatting marks on (BANA-template documents
    lean heavily on visible tags like $pg and DBT bracket codes, which are
    much easier to audit with marks visible).
    """
    set_formatting_marks_visible(doc, True)
    return {"formatting_marks": True}
