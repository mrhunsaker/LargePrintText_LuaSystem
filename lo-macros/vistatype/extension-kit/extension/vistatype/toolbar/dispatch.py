"""Explicit-document dispatcher used by the UNO toolbar controller."""
from ..lp import (
    attach_template as lp_attach_template,
    file_cleanup as lp_file_cleanup,
    reference_pages as lp_reference_pages,
    formatting_tools as lp_formatting_tools,
)
from .. import utils


def run_lp_command(command_name, doc, desktop):
    if command_name == "LP_Attach_LP_Template":
        return lp_attach_template.attach_lp_template(doc, desktop=desktop)
    if command_name == "Lp_File_Fix_Sequence":
        return lp_file_cleanup.file_fix_sequence(doc)
    if command_name == "Lp_Selected_File_CleanUp":
        return lp_file_cleanup.selected_cleanup(doc)
    if command_name == "Lp_AutoTag_Page_Numbers":
        return lp_reference_pages.auto_tag_page_numbers(doc)
    if command_name == "Lp_Manual_Tag_with_Dollar_pg":
        return lp_reference_pages.manual_tag_page_number(doc)
    if command_name == "Lp_Format_Page_Numbers":
        return lp_reference_pages.format_tagged_page_numbers(doc)
    if command_name == "Lp_Validate_Dollar_PG":
        entries = lp_reference_pages.validate_tagged_page_numbers(doc)
        scratch = desktop.loadComponentFromURL(
            "private:factory/swriter", "_blank", 0,
            (utils.make_prop("Hidden", False),),
        )
        lines = [
            f"{e['index']:>4}  {e['text']:<20} "
            f"{'' if e['in_order'] else '<-- CHECK ORDER'}"
            for e in entries
        ]
        scratch.Text.setString(
            "\n".join(lines) if lines else "No $pg-tagged paragraphs found."
        )
        return
    if command_name == "Lp_Horz_List_To_Vertical":
        return lp_formatting_tools.horizontal_list_to_vertical(doc, mode="auto")
    if command_name == "Lp_Compress_Linear_Math":
        return lp_formatting_tools.compress_linear_math(doc)
    if command_name == "Lp_Toggle_Space_After_Current_Para":
        return lp_formatting_tools.toggle_space_after_para(doc)
    if command_name == "Lp_Keep_With_Next_Para":
        return lp_formatting_tools.toggle_keep_with_next_para(doc)
    if command_name == "Lp_Type_Fill_In_Line":
        return lp_formatting_tools.type_fill_in_line(doc, length=8)
    if command_name == "Lp_Format_Exercise_Lv_1_and_Lv_2":
        return lp_formatting_tools.format_exercise_levels(doc)
    raise ValueError(f"Unknown LP VISTATYPE command: {command_name}")
