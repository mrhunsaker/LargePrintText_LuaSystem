# Findings from inspecting the source files

Before writing any code, the three source files were actually unpacked and
inspected on disk (not just read as PDFs) — here's what that turned up,
since it materially shaped what was and wasn't possible to build.

## `Word.officeUI`

A real Office ribbon/QAT customization XML file. It references exactly 45
distinct commands via `onAction="..."` attributes:

- `Dx_*` (17) — Braille/BANA Template macros
- `Lp_*`/`LP_*` (17) — Large Print (VistaType LP) macros
- `Sh_*` (5) — shared between both
- `MS_*` (3) — Word/application configuration
- `DN_*` (2) — DAISY/NIMAS/text-file tools

This list is ground truth for the full command inventory — see
`docs/COMMAND_MAP.md` for the complete extraction and per-command status.

## `LargePrintTemplate.dotx`

A clean, macro-free Word template (confirmed: no `vbaProject.bin`, just
styles/settings/page setup). Converting it through headless LibreOffice
(`soffice --headless --convert-to odt`) and inspecting the result byte-for-
byte confirmed full fidelity:

- `Normal` → Tahoma 18pt, +1pt letter-spacing, 0.25in space-after
- `Heading 1` → 28pt bold, +2pt letter-spacing (based on Normal)
- `Print Pg Num` → top/bottom border, `#F4A3D4` pink shading, right tab
  stop at 10598 twips, `keepNext`
- 22 accessibility contrast styles: `Para*`/`Box*` are **paragraph**
  styles (whole-paragraph background/border color); `Words*`/`Text*` are
  **character** styles (selected-word/phrase color or highlight) — this
  split wasn't obvious from the guide text alone, only from the raw XML's
  `w:type="paragraph"` vs `w:type="character"` attribute
- Page geometry: Letter (8.5×11in), 0.5in margins all around

This is used directly as the Pandoc `--reference-doc`, and its styles are
imported live into LibreOffice documents by
`vistatype.lp.attach_template.attach_lp_template` (see that module's
docstring for why a direct style-loader API wasn't used instead).

## `LPandBRL.dotm` — VBA stomping (the important one)

Extracting the VBA source with `oletools.olevba` produced a flagged
warning: **"Suspicious VBA Stomping"** — the visible source code doesn't
match the compiled p-code that Word actually executes when the macros run.
This was verified concretely, not just taken on the tool's say-so:

1. Cleanly extracted all 45 real "VBA MACRO" code streams (filtering out
   ~190 `"VBA FORM STRING"` entries, which turned out to be garbled binary
   picture-resource data embedded in UserForms, not code).
2. Searched the ~45 extracted modules for the actual ribbon-callback
   entry points — `Dx_Attach_BANA_Template`, `Lp_Horz_List_To_Vertical`,
   `Sh_Doc_Info`, all 45 names from `Word.officeUI`. **None of them exist
   anywhere in the extracted source.**
3. Listed every `Sub`/`Function` name that IS present: one hundred percent
   of them are UserForm button-click handlers (`CmdOkay_Click`,
   `CmdCancel_Click`, and similar), each just calling out to a *differently
   named* worker macro via `Application.Run MacroName:="..."` — a dynamic
   string-based dispatch, so even the visible click-handlers don't
   reference the real logic by a name that shows up in a static search.

Net effect: **the actual implementations of all ~150 real macros are not
recoverable from this file.** What's readable is dialog/form scaffolding —
genuinely useful for confirming exact option wording and structure (e.g.
the Horizontal-List-to-Vertical dialog's three list-type radio buttons, or
the Dashes/Primes/Fractions button-grid layout), but it contains no
algorithm bodies.

This isn't necessarily evidence of malicious intent — VBA source-
protection tools that deliberately strip visible source while leaving
compiled p-code intact produce exactly this same signature, and plenty of
legitimate freeware authors use them to protect their work from casual
copying. But it does mean this project is a **reconstruction from the two
PDF user guides** (43 and 162 pages, both read in full), not a translation
of existing code — which is a meaningfully different, more judgment-heavy
undertaking, and the reason this was built incrementally with real testing
at each step rather than all at once.

### One concrete consequence: an invented detail was caught and removed

While writing `vistatype.dx.file_cleanup.format_spelling_list`, an early
draft invented a plausible-looking DBT bracket code
(`[[*unc*]]...[[*unc-e*]]`) for "render this span uncontracted," by
analogy with the *real* DBT codes the guide does spell out verbatim
elsewhere (`$pg`, the continuation-page code, the letter-sign code, the
fraction codes). On review, the guide never actually states literal
syntax for this one — so that invented code was removed before it could
ship. Getting a DBT bracket code wrong isn't a cosmetic bug: it risks
producing genuinely incorrect braille for a blind reader. The shipped
version does only the mechanical, verifiably-correct part (duplicating
the text) and says plainly, in the function's docstring, what it doesn't
know — see that function for the full explanation.

This is flagged here explicitly as a reminder for anyone extending this
project: where a DBT/BANA bracket code isn't directly quoted in a guide,
don't guess at one — ask, or leave it out and say so.
