"""
vistatype.entrypoints
========================
Thin wrappers that make every function in this package runnable as a real
LibreOffice macro -- bindable to a toolbar button, menu item, or keyboard
shortcut via Tools > Customize, or run directly via
Tools > Macros > Run Macro.

WHY THIS FILE IS SEPARATE FROM THE REST OF THE PACKAGE
=========================================================
Every function elsewhere in this package (vistatype.shared,
vistatype.lp.*, vistatype.dx.*, ...) takes an explicit `doc` argument and
returns a plain Python value, on purpose: that's what makes them directly
unit-testable from an external UNO connection (see lo-macros/tests/).

LibreOffice's own scripting provider, on the other hand, calls a macro
with NO arguments and ignores whatever it returns -- it expects the macro
to reach for a magic global, `XSCRIPTCONTEXT`, to find out which document
and frame it's running against. That global only exists when LibreOffice
itself invokes the script; it does not exist in a plain Python process
connected over a UNO socket, which is why none of the "real" logic
functions reference it directly (see vistatype.utils.get_active_document
for the one place that boundary is enforced with a clear error rather than
a confusing NameError somewhere deep in a call stack).

INSTALLING THESE AS REAL MACROS
==================================
Copy (or symlink) the whole `vistatype/` package into LibreOffice's user
Python scripts directory:

    Linux:   ~/.config/libreoffice/4/user/Scripts/python/
    macOS:   ~/Library/Application Support/LibreOffice/4/user/Scripts/python/
    Windows: %APPDATA%\\LibreOffice\\4\\user\\Scripts\\python\\

so you end up with e.g.
    .../Scripts/python/vistatype/entrypoints.py
    .../Scripts/python/vistatype/shared.py
    .../Scripts/python/vistatype/lp/...
    .../Scripts/python/vistatype/dx/...

Restart LibreOffice, then Tools > Macros > Organize Macros > Python should
show "vistatype.entrypoints" with every function below listed and ready
to bind to a toolbar button (Tools > Customize > Toolbars > Add > Category
"My Macros" or similar, depending on LO version).
"""
from __future__ import annotations
import sys
import os

# --- FIX: Ensure the vistatype package is importable in LibreOffice's Python ---
# Get the absolute path to the directory containing this file (entrypoints.py)
this_dir = os.path.dirname(os.path.abspath(__file__))
# Get the parent directory (Scripts/python/vistatype/ -> Scripts/python/)
scripts_dir = os.path.dirname(this_dir)
# Add the Scripts/python/ directory to sys.path so Python can find 'vistatype'
if scripts_dir not in sys.path:
    sys.path.insert(0, scripts_dir)

# --- Absolute imports for the vistatype package ---
from vistatype import shared, config, dn_tools, utils
from vistatype.lp import (
    attach_template as lp_attach_template,
    file_cleanup as lp_file_cleanup,
    reference_pages as lp_reference_pages,
    formatting_tools as lp_formatting_tools,
)
from vistatype.dx import (
    reference_pages as dx_reference_pages,
    formatting_tools as dx_formatting_tools,
    file_cleanup as dx_file_cleanup,
)


def _ctx_doc_desktop():
    doc = XSCRIPTCONTEXT.getDocument()  # noqa: F821
    desktop = XSCRIPTCONTEXT.getDesktop()  # noqa: F821
    return doc, desktop


# --- Sh_ shared -------------------------------------------------------------

def Sh_Apply_Title_Case_Capitalization(*_):
    doc, _d = _ctx_doc_desktop()
    shared.apply_title_case(doc)


def Sh_Doc_Info(*_):
    doc, _d = _ctx_doc_desktop()
    info = shared.doc_info(doc)
    utils.show_message_box(doc, "\n".join(f"{k}: {v}" for k, v in info.items()), title="Document Info")


def Sh_Keep_Lines_Of_Para_Together(*_):
    doc, _d = _ctx_doc_desktop()
    shared.toggle_keep_lines_together(doc)


def Sh_Move_Paragraph_To_Next_Page(*_):
    doc, _d = _ctx_doc_desktop()
    shared.toggle_move_to_next_page(doc)


def Sh_Show_Char_Val(*_):
    doc, _d = _ctx_doc_desktop()
    val = shared.char_value_at_cursor(doc)
    msg = f"Character: {val['char']!r}  Decimal: {val['decimal']}  Hex: {val['hex']}" if val else "No character to the right of the cursor."
    utils.show_message_box(doc, msg, title="Character Value")


# --- MS_ configuration -------------------------------------------------------

def MS_Change_Word_Configuration_to_New_Install(*_):
    doc, _d = _ctx_doc_desktop()
    config.set_config_new_install(doc)


def MS_Change_Word_Configuration_to_Large_Print(*_):
    doc, _d = _ctx_doc_desktop()
    config.set_config_large_print(doc)


def MS_Change_Word_Configuration_to_Braille(*_):
    doc, _d = _ctx_doc_desktop()
    config.set_config_braille(doc)


# --- LP_/Lp_ Large Print -----------------------------------------------------

def LP_Attach_LP_Template(*_):
    doc, desktop = _ctx_doc_desktop()
    lp_attach_template.attach_lp_template(doc, desktop=desktop)


def Lp_File_Fix_Sequence(*_):
    doc, _d = _ctx_doc_desktop()
    lp_file_cleanup.file_fix_sequence(doc)


def Lp_Selected_File_CleanUp(*_):
    doc, _d = _ctx_doc_desktop()
    lp_file_cleanup.selected_cleanup(doc)


def Lp_AutoTag_Page_Numbers(*_):
    doc, _d = _ctx_doc_desktop()
    lp_reference_pages.auto_tag_page_numbers(doc)


def Lp_Manual_Tag_with_Dollar_pg(*_):
    doc, _d = _ctx_doc_desktop()
    lp_reference_pages.manual_tag_page_number(doc)


def Lp_Format_Page_Numbers(*_):
    doc, _d = _ctx_doc_desktop()
    lp_reference_pages.format_tagged_page_numbers(doc)


def Lp_Validate_Dollar_PG(*_):
    doc, desktop = _ctx_doc_desktop()
    entries = lp_reference_pages.validate_tagged_page_numbers(doc)
    scratch = desktop.loadComponentFromURL(
        "private:factory/swriter", "_blank", 0,
        (utils.make_prop("Hidden", False),),
    )
    lines = [f"{e['index']:>4}  {e['text']:<20} {'' if e['in_order'] else '<-- CHECK ORDER'}" for e in entries]
    scratch.Text.setString("\n".join(lines) if lines else "No $pg-tagged paragraphs found.")


def Lp_Horz_List_To_Vertical(*_):
    doc, _d = _ctx_doc_desktop()
    lp_formatting_tools.horizontal_list_to_vertical(doc, mode="auto")


def Lp_Compress_Linear_Math(*_):
    doc, _d = _ctx_doc_desktop()
    lp_formatting_tools.compress_linear_math(doc)


def Lp_Toggle_Space_After_Current_Para(*_):
    doc, _d = _ctx_doc_desktop()
    lp_formatting_tools.toggle_space_after_para(doc)


def Lp_Keep_With_Next_Para(*_):
    doc, _d = _ctx_doc_desktop()
    lp_formatting_tools.toggle_keep_with_next_para(doc)


def Lp_Type_Fill_In_Line(*_):
    doc, _d = _ctx_doc_desktop()
    lp_formatting_tools.type_fill_in_line(doc, length=8)


def Lp_Format_Exercise_Lv_1_and_Lv_2(*_):
    doc, _d = _ctx_doc_desktop()
    lp_formatting_tools.format_exercise_levels(doc)


# --- Dx_ Braille/BANA --------------------------------------------------------

def Dx_AutoTag_Page_Numbers(*_):
    doc, _d = _ctx_doc_desktop()
    dx_reference_pages.auto_tag_page_numbers(doc)


def Dx_Manual_Tag_with_Dollar_pg(*_):
    doc, _d = _ctx_doc_desktop()
    dx_reference_pages.manual_tag_page_number(doc)


def Dx_Format_Tagged_Page_Numbers(*_):
    doc, _d = _ctx_doc_desktop()
    dx_reference_pages.format_tagged_page_numbers(doc)


def Dx_Compress_Linear_Math(*_):
    doc, _d = _ctx_doc_desktop()
    dx_formatting_tools.compress_linear_math(doc)


def Dx_Horz_List_To_Vertical(*_):
    doc, _d = _ctx_doc_desktop()
    dx_formatting_tools.horizontal_list_to_vertical(doc, mode="auto")


def Dx_File_Fix_Sequence(*_):
    doc, _d = _ctx_doc_desktop()
    dx_file_cleanup.fix_common_file_errors(doc)
    dx_file_cleanup.delete_multiple_paragraph_marks(doc)


def Dx_Selected_File_CleanUp(*_):
    doc, _d = _ctx_doc_desktop()
    dx_file_cleanup.selected_cleanup(doc)


def Dx_Spelling_List(*_):
    doc, _d = _ctx_doc_desktop()
    dx_file_cleanup.format_spelling_list(doc)


# --- DN_ DAISY/NIMAS/Text -----------------------------------------------------

def DN_Remove_Para_Formatting_From_Text_Files(*_):
    """Unlike the others, this one operates on the document's own plain
    text in place (select-all, fix, replace), since DN_ tools are
    fundamentally text-level operations -- see vistatype.dn_tools.
    """
    doc, _d = _ctx_doc_desktop()
    doc.Text.setString(dn_tools.fix_text_file_paragraphs(doc.Text.getString()))


# Every *public* function above must be listed here for LibreOffice's
# Python provider to expose it as a runnable macro.
g_exportedScripts = (
    Sh_Apply_Title_Case_Capitalization, Sh_Doc_Info, Sh_Keep_Lines_Of_Para_Together,
    Sh_Move_Paragraph_To_Next_Page, Sh_Show_Char_Val,
    MS_Change_Word_Configuration_to_New_Install, MS_Change_Word_Configuration_to_Large_Print,
    MS_Change_Word_Configuration_to_Braille,
    LP_Attach_LP_Template, Lp_File_Fix_Sequence, Lp_Selected_File_CleanUp,
    Lp_AutoTag_Page_Numbers, Lp_Manual_Tag_with_Dollar_pg, Lp_Format_Page_Numbers,
    Lp_Validate_Dollar_PG, Lp_Horz_List_To_Vertical, Lp_Compress_Linear_Math,
    Lp_Toggle_Space_After_Current_Para, Lp_Keep_With_Next_Para, Lp_Type_Fill_In_Line,
    Lp_Format_Exercise_Lv_1_and_Lv_2,
    Dx_AutoTag_Page_Numbers, Dx_Manual_Tag_with_Dollar_pg, Dx_Format_Tagged_Page_Numbers,
    Dx_Compress_Linear_Math, Dx_Horz_List_To_Vertical, Dx_File_Fix_Sequence,
    Dx_Selected_File_CleanUp, Dx_Spelling_List,
    DN_Remove_Para_Formatting_From_Text_Files,
)
