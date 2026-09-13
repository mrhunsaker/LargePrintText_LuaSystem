"""
vistatype.dn_tools
====================
Port of the "DN_" (DAISY/NIMAS) ribbon commands
(VistaType LP User Guide p.29-32, "Guide to Special Procedures" +
BANA Macros guide p.26, same two tools):

  DN_Add_PgNo_Tags_To_DAISY_or_NIMAS         -> add_pg_tags_to_xml
  DN_Remove_Para_Formatting_From_Text_Files  -> fix_text_file_paragraphs

Both of these are plain-text/XML transformations on files on disk, not
operations on an open Writer document -- the original macros work by
round-tripping the content through a blank Word document only because
VBA's own string/regex tools are weaker than editing text directly. We
skip that round trip and operate on the file content directly, which is
more robust and testable, and works identically whether or not
LibreOffice is even running.
"""
from __future__ import annotations

import re

# Reuse the exact same reference-page-number pattern rules as
# vistatype.lp.reference_pages, so a NIMAS/DAISY XML file tagged here and
# a Word-sourced document tagged there end up using the same $pg
# convention.
from .lp.reference_pages import _classify_page_number


def add_pg_tags_to_xml(xml_text: str) -> str:
    """DN_Add_PgNo_Tags_To_DAISY_or_NIMAS.

    NIMAS/DAISY XML marks original-print page boundaries with
    <pagenum .../> (self-closing) or <pagenum ...>N</pagenum> elements
    (the DTBook/NIMAS "pagenum" element; see the guide's description of
    the NIMAS/DAISY file set, p.29-30). This finds the page-number text
    inside each such element and prepends "$pg" to it, exactly like
    auto_tag_page_numbers does for paragraph text in a live document --
    so downstream processing (pasting into Word/LO and running
    Format_Page_Numbers / format_tagged_page_numbers) treats them
    identically.

    Handles both:
        <pagenum ...>12</pagenum>
        <pagenum ... value="12"/>
    """
    def repl_element_text(m):
        open_tag, number, close_tag = m.group(1), m.group(2), m.group(3)
        if number.startswith("$pg"):
            return m.group(0)
        kind, value = _classify_page_number(number)
        tagged = f"$pg{value}" if kind else number
        return f"{open_tag}{tagged}{close_tag}"

    xml_text = re.sub(
        r'(<pagenum\b[^>]*>)([^<]+)(</pagenum>)',
        repl_element_text,
        xml_text,
        flags=re.IGNORECASE,
    )

    def repl_attr(m):
        prefix, number = m.group(1), m.group(2)
        if number.startswith("$pg"):
            return m.group(0)
        kind, value = _classify_page_number(number)
        tagged = f"$pg{value}" if kind else number
        return f'{prefix}{tagged}"'

    xml_text = re.sub(
        r'(<pagenum\b[^>]*\bvalue=")([^"]+)"',
        repl_attr,
        xml_text,
        flags=re.IGNORECASE,
    )
    return xml_text


def fix_text_file_paragraphs(text: str) -> str:
    """DN_Remove_Para_Formatting_From_Text_Files ("Txt File Para Fix").

    Some downloadable text files hard-wrap every line (a paragraph mark at
    the end of each line) and use one blank line between real paragraphs.
    This removes the within-paragraph line breaks while preserving the
    blank-line-delimited paragraph boundaries, so the text can re-flow
    normally once brought into Writer -- matching the guide's example
    (Tom Sawyer dialogue sample, p.26-27) exactly: consecutive
    non-blank lines get joined with a single space; a blank line marks a
    real paragraph break and is preserved as exactly one blank line.
    """
    # Normalize line endings first.
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = normalized.split("\n")

    paragraphs = []
    current = []
    for line in lines:
        if line.strip() == "":
            if current:
                paragraphs.append(" ".join(current))
                current = []
            paragraphs.append("")  # preserve the blank line itself
        else:
            current.append(line.strip())
    if current:
        paragraphs.append(" ".join(current))

    # Collapse any run of multiple blank-line markers we may have
    # accumulated (e.g. from multiple consecutive blank lines in the
    # source) down to a single blank line between paragraphs.
    collapsed = []
    prev_blank = False
    for p in paragraphs:
        is_blank = p == ""
        if is_blank and prev_blank:
            continue
        collapsed.append(p)
        prev_blank = is_blank

    return "\n".join(collapsed)
