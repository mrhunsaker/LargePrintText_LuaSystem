# Command map: Word.officeUI → this project

This is the full, verified inventory of every ribbon/QAT command referenced
by `Word.officeUI` (45 total, extracted directly from the file's
`onAction="..."` attributes — this list is ground truth, not a guess), and
where each one lives now.

**Status legend**
- ✅ **Implemented & tested** — has a real implementation, exercised against
  a live headless LibreOffice instance in `lo-macros/tests/`.
- 🟡 **Partially implemented** — core behavior works; a documented,
  specific piece of the original is out of scope (see notes).
- ⬜ **Not implemented (roadmap)** — not built yet; notes explain why and
  what it would take.
- ⛔ **Deliberately not ported** — trivial/out-of-scope by nature (About
  dialogs, links to external videos), not because it's hard.

Every function below was implemented **from the two PDF user guides**, not
from the original VBA source. That source turned out to be unusable for
this purpose — see `docs/FINDINGS.md` for the full technical writeup, but
in short: `LPandBRL.dotm`'s visible VBA is "stomped" (the compiled p-code
that actually runs doesn't match the visible source), and none of the
~150 real macro implementations the ribbon calls by name are present in
what can be statically extracted. What *is* present is ~190 UserForm
button-click handlers, which were useful for confirming dialog option
shapes (e.g. the exact wording of the Horizontal-List-to-Vertical dialog's
three list-type options) but contain no algorithm bodies.

---

## Sh_ (shared between Lp_ and Dx_) — 5/5 ✅

| Command | Status | Implementation |
|---|---|---|
| `Sh_Apply_Title_Case_Capitalization` | ✅ | `vistatype.shared.apply_title_case` |
| `Sh_Doc_Info` | ✅ | `vistatype.shared.doc_info` |
| `Sh_Keep_Lines_Of_Para_Together` | ✅ | `vistatype.shared.toggle_keep_lines_together` (toggles `ParaSplit`) |
| `Sh_Move_Paragraph_To_Next_Page` | ✅ | `vistatype.shared.toggle_move_to_next_page` (toggles `BreakType.PAGE_BEFORE`) |
| `Sh_Show_Char_Val` | 🟡 | `vistatype.shared.char_value_at_cursor` — reports Unicode code point (decimal + hex), not a legacy ANSI/code-page value, since LibreOffice has no such layer to inspect (see function docstring) |

## MS_ (Word/LO configuration) — 3/3 🟡

| Command | Status | Implementation |
|---|---|---|
| `MS_Change_Word_Configuration_to_New_Install` | 🟡 | `vistatype.config.set_config_new_install` |
| `MS_Change_Word_Configuration_to_Large_Print` | 🟡 | `vistatype.config.set_config_large_print` |
| `MS_Change_Word_Configuration_to_Braille` | 🟡 | `vistatype.config.set_config_braille` |

All three only change what's genuinely a **per-document** setting in LO
(formatting-mark visibility, default paragraph style). The original also
swaps a whole AutoCorrect/AutoFormat-as-you-type profile, which is an
**application-wide** LibreOffice preference (Tools ▸ AutoCorrect Options),
not something a per-document macro should silently override — see the
module docstring for the full reasoning.

## LP_/Lp_ (Large Print) — 13/17 ✅, 3/17 ⬜, 1/17 ⛔

| Command | Status | Implementation |
|---|---|---|
| `LP_Attach_LP_Template` | 🟡 | `vistatype.lp.attach_template.attach_lp_template` — imports all real styles + page geometry from the bundled reference doc (verified byte-for-byte against the source `.dotx`). **Not covered:** the paper/screen-size-and-binding wizard and DBT-media-selection dialog described on p.35-43 of the guide — that's a large, separate feature (see "Not yet built" below). |
| `Lp_AutoTag_Page_Numbers` | ✅ | `vistatype.lp.reference_pages.auto_tag_page_numbers` |
| `Lp_Manual_Tag_with_Dollar_pg` | ✅ | `vistatype.lp.reference_pages.manual_tag_page_number` |
| `Lp_Format_Page_Numbers` | ✅ | `vistatype.lp.reference_pages.format_tagged_page_numbers` |
| `Lp_Validate_Dollar_PG` | ✅ | `vistatype.lp.reference_pages.validate_tagged_page_numbers` (returns structured data; `entrypoints.py` wraps it in a scratch-document view, matching the original's "temporary document" workflow) |
| `Lp_File_Fix_Sequence` | 🟡 | `vistatype.lp.file_cleanup.file_fix_sequence` — see the module docstring for the itemized list of which of the 15+ original sub-fixes are/aren't covered (Abbyy-style remap and pixel-perfect text-box re-anchoring are the two gaps) |
| `Lp_Selected_File_CleanUp` | ✅ | `vistatype.lp.file_cleanup.selected_cleanup` |
| `Lp_Horz_List_To_Vertical` | 🟡 | `vistatype.lp.formatting_tools.horizontal_list_to_vertical` — "ordered" (lettered/numbered) mode is auto-detect-safe and fully implemented; "spaced"/"tabbed" modes work but require you to say so explicitly (see function docstring for why auto-detecting those two would be unsafe on ordinary prose) |
| `Lp_Compress_Linear_Math` | ✅ | `vistatype.lp.formatting_tools.compress_linear_math` |
| `Lp_Toggle_Space_After_Current_Para` | ✅ | `vistatype.lp.formatting_tools.toggle_space_after_para` |
| `Lp_Keep_With_Next_Para` | ✅ | `vistatype.lp.formatting_tools.toggle_keep_with_next_para` (toggles `ParaKeepTogether`) |
| `Lp_Type_Fill_In_Line` | 🟡 | `vistatype.lp.formatting_tools.type_fill_in_line` — fixed-length (1-21) matches the guide directly; "fill to right margin" is implemented via a LibreOffice-native right tab stop with `_` as its fill character (a different mechanism from Word's unknown/unrecoverable one, same visual/functional result — see function docstring) |
| `Lp_Format_Exercise_Lv_1_and_Lv_2` | ✅ | `vistatype.lp.formatting_tools.format_exercise_levels` |
| `Lp_Picture_Color_Change_Menu` | ⬜ | Not built. Guide (p.51-52ish) describes a screen-background-color toggle (paired with the `*Inverted` character styles) and picture color/contrast tools. The `Words Black Inverted` / `Para Black Inverted` etc. styles are already faithfully imported by `attach_lp_template`, but the toggle UI and picture-recoloring logic itself isn't written yet. |
| `Lp_Picture_Tools_Menu_Starter` | ⬜ | Not built (picture resize/positioning helpers — guide section not yet mined in detail). |
| `Lp_Table_Tools` | ⬜ | Not built. Guide's Attach Group section mentions table rotation/splitting/conversion-to-list features tied to this; substantial enough (and different enough from anything else here) to deserve its own dedicated pass rather than a rushed partial port. |
| `Lp_About` | ⛔ | Static text (license, version). Not a workflow function. |
| `Lp_Video_Links` | ⛔ | Opens the author's external video-hosting links. Nothing to port. |

## Dx_ (Braille/BANA) — 8/17 ✅, 6/17 ⬜, 3/17 ⛔

Built as a representative, tested slice rather than exhaustively —
prioritized functions with a clearly documented, DBT-specific behavior
that's genuinely different from the Large Print side (continuation-page
codes, letter-sign codes, the UEB/EBAE dash-space distinction), since
those are the parts most likely to be silently wrong if ported carelessly.

| Command | Status | Implementation |
|---|---|---|
| `Dx_AutoTag_Page_Numbers` | ✅ | `vistatype.dx.reference_pages.auto_tag_page_numbers` — includes the DBT continuation-page code and the UEB-vs-EBAE lowercase-roman-numeral letter-sign rule |
| `Dx_Manual_Tag_with_Dollar_pg` | ✅ | `vistatype.dx.reference_pages.manual_tag_page_number` |
| `Dx_Format_Tagged_Page_Numbers` | 🟡 | `vistatype.dx.reference_pages.format_tagged_page_numbers` — styles with a placeholder paragraph style (`BANA Braille Template Reference Page Number`), since we don't have a real BANA template file to draw its actual formatting from (only `LargePrintTemplate.dotx` was provided — see `FINDINGS.md`) |
| `Dx_Ref_Pg_Number_Sequence_Menu` | ✅ | `vistatype.dx.reference_pages.validate_tagged_page_numbers` |
| `Dx_File_Fix_Sequence` | ✅ | `vistatype.dx.file_cleanup.fix_common_file_errors` + re-exported `delete_multiple_paragraph_marks` — includes the EBAE-only dash-space rule the guide explicitly calls out as different from the LP side |
| `Dx_Selected_File_CleanUp` | ✅ | `vistatype.dx.file_cleanup.selected_cleanup` |
| `Dx_Compress_Linear_Math` | ✅ | `vistatype.dx.formatting_tools.compress_linear_math` — deliberately different rule from the LP version of the same name (keeps spaces around comparison operators instead of hair-spacing them; see docstring) |
| `Dx_Horz_List_To_Vertical` | ✅ | re-exported from `vistatype.lp.formatting_tools` — guide describes identical behavior on both sides |
| `Dx_Type_Dashes` | ✅ | `vistatype.dx.formatting_tools.insert_dash` / `insert_prime` / `insert_fraction` / `encode_fractions_in_selection` |
| `Dx_Spelling_List` | 🟡 | `vistatype.dx.file_cleanup.format_spelling_list` — **only does the mechanical text-duplication part.** The guide never states DBT's literal bracket code for "force uncontracted braille," unlike every other DBT code used elsewhere in this project (`$pg`, the continuation code, the letter-sign code, the fraction codes — all of which DO appear verbatim in the guide and are used as-is). Guessing at one here risked producing genuinely wrong braille for a blind reader, so the function docstring says exactly that instead of inventing a plausible-looking code. If you know DBT's real code for this, tell me and it's a one-line fix. |
| `Dx_Format_Exercise_Lv_1_and_Lv_2` | ⬜ | Not built. Same shape as the LP version but with BANA-specific `Exercise1`/`Exercise2` styles and EBAE/UEB-specific fill-in characters (four hyphens vs one underscore, per the guide) instead of the LP side's `List`/`List 2` + underlined underscores — different enough to need its own pass, not a copy-paste of the LP version. |
| `Dx_Add_Color_To_Foreign_Language_Words` | ⬜ | Not built. Conceptually straightforward in LO (character `CharLocale` already tags language per run; would map to the BANA template's foreign-language character styles), but those target styles aren't in the one template file we have (`LargePrintTemplate.dotx` has no foreign-language styles — that's a Braille-template-specific need). |
| `Dx_Embed_Ref_Pg_No` / `Dx_UnEmbed_Ref_Pg_No` | ⬜ | Not built. Converts a reference page number between "between paragraphs" and "mid-paragraph" placement; needs the real BANA "Embedded (Mid-Paragraph) Reference Page Number" style, which we don't have a source file for. |
| `Dx_About` | ⛔ | Static text. |
| `Dx_Video_Links` | ⛔ | External links. |

## DN_ (DAISY/NIMAS/Text tools) — 2/2 ✅

| Command | Status | Implementation |
|---|---|---|
| `DN_Add_PgNo_Tags_To_DAISY_or_NIMAS` | ✅ | `vistatype.dn_tools.add_pg_tags_to_xml` — operates directly on XML text (not via a Word round-trip like the original), same `$pg` convention as everything else here |
| `DN_Remove_Para_Formatting_From_Text_Files` | ✅ | `vistatype.dn_tools.fix_text_file_paragraphs` |

---

## The Pandoc/ODT batch pipeline (separate from the LibreOffice macros above)

The macros above are all **interactive**, operating on a document you
already have open — they mirror the original's "place your cursor, then
run the macro" UX. The `pandoc/filters/` Lua filters are a **different,
batch** pipeline for the specific ask of "convert a `.docx`/`.md` straight
to `.odt` with LargePrintTemplate characteristics," and only cover the
subset of macros that are safe to run unattended over a whole document
with no human in the loop:

| Filter | Covers | Notes |
|---|---|---|
| `reference-doc/build_reference_odt.sh` | `LP_Attach_LP_Template` (styles/page-setup part) | Verified: Normal/18pt/letter-spacing, Heading 1, all Para\*/Words\*/Text\*/Box\* contrast styles, 8.5×11 @ 0.5in margins all survive and are correctly inherited by Pandoc's own generated styles |
| `lp-delete-multiple-paragraph-marks.lua` | `Lp_File_Fix_Sequence` (step 2) | Turned out to be close to a no-op for DOCX input specifically — Pandoc's own DOCX reader already drops fully-empty paragraphs before the AST exists (see file header comment); kept for ODT input and Markdown's nbsp-paragraph idiom |
| `lp-fix-common-file-errors.lua` | `Lp_File_Fix_Sequence` (step 1), partial | Dash-run collapsing, space-stripping around dashes, small-caps flattening, prime conversion, edge-whitespace trimming. Drop caps, text-box un-boxing, and confirmable image removal are NOT in this filter — those need a live document object model, not an AST; use the LibreOffice macro (`vistatype.lp.file_cleanup.fix_common_file_errors`) for those |
| `lp-horizontal-list-to-vertical.lua` | `Lp_Horz_List_To_Vertical`, "ordered" mode only | Same 3-marker-minimum safety threshold as the LO macro, tuned to avoid false-positiving on ordinary prose like "I went to a) the store and b) the bank" |
| `lp-tag-reference-page-numbers.lua` + `lp-format-tagged-page-numbers.lua` | `Lp_AutoTag_Page_Numbers` + `Lp_Format_Page_Numbers` | Verified end-to-end against a real LibreOffice instance: all four page-number shapes (plain, lettered, range, roman numeral) get tagged and correctly styled with the genuine "Print Pg Num" formatting (background `#F4A3D4` confirmed) |

Nothing selection-scoped (Manual Tag, Selected Cleanup, Type Fill-In Line,
Format Exercise Levels, Compress Linear Math on a specific paragraph, ...)
is in the batch pipeline, on purpose — there's no "selection" in a batch
Pandoc conversion, and guessing which paragraph the user means would be
worse than just pointing at the interactive LibreOffice macro instead.

---

## Not yet built, and roughly how big each piece is

- **Attach LP Template's paper/screen-size wizard** (p.35-43): lets the
  transcriber pick target paper size, binding/mirroring, and screen-vs-
  print output profiles at attach time. This is a substantial, mostly-new
  UI/dialog-driven feature in its own right, not a style-copy operation —
  worth its own dedicated design pass.
- **`Lp_Table_Tools`** and **`Lp_Picture_Tools_Menu_Starter`/
  `Lp_Picture_Color_Change_Menu`**: guide sections for these weren't fully
  mined yet (162 pages total; the Attach/File-Cleanup/Reference-Page/
  Other-Formatting/DAISY/Help groups were prioritized first since they
  cover the core transcription workflow end to end). Table tools in
  particular (rotation, splitting titles, list conversion) look like a
  meaningfully-sized feature.
- **`Dx_Format_Exercise_Lv_1_and_Lv_2`**, **`Dx_Add_Color_To_Foreign_
  Language_Words`**, **`Dx_Embed_Ref_Pg_No`/`Dx_UnEmbed_Ref_Pg_No`**: all
  need real BANA-template style names/definitions we don't have a source
  file for (only `LargePrintTemplate.dotx` was uploaded — no BANA `.dotx`).
  Buildable once you can point me at one, or tell me the exact style
  names/properties to hardcode as a fallback the way `Print Pg Num` and
  `List`/`List 2` are handled now.
