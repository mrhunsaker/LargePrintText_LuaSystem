# VistaType LP → LibreOffice + Pandoc

A ground-up, **LibreOffice-native** reimplementation of the VistaType LP /
BANA Braille Template macro suite — a Word VBA add-in originally written by
Jerry Whittaker for producing large-print and braille-ready transcriptions.
This project contains **no VBA and no dependency on Microsoft Word**. It's
two independent components that share a common template/style source of
truth:

1. **A Pandoc conversion pipeline** (`pandoc/`) — Lua filters plus a
   verified reference `.odt`, so a single `pandoc` command converts a
   `.docx` or `.md` file straight to `.odt` carrying `LargePrintTemplate`'s
   real fonts, spacing, page geometry, and accessibility styles.
2. **A LibreOffice macro package** (`lo-macros/`) — pure Python using the
   UNO API, installable as real toolbar-button/menu macros inside
   LibreOffice Writer, covering the interactive (cursor- and
   selection-based) tools from the original ribbon.

If you only read one more paragraph of this file, read this one: **the
original `.dotm`'s VBA source was not usable as a translation source** (see
[Why this is a reconstruction, not a translation](#why-this-is-a-reconstruction-not-a-translation)
below for exactly why and how that was confirmed). Everything in this
repository was built from the two PDF user guides that document the
add-in's *behavior*, then verified function-by-function against a real,
running LibreOffice instance. Where that verification surfaced a real gap
or a place this project had to guess, it's called out explicitly rather
than silently papered over — see
[Known limitations & unfinished work](#known-limitations--unfinished-work).

---

## Table of contents

- [Why this is a reconstruction, not a translation](#why-this-is-a-reconstruction-not-a-translation)
- [Repository structure](#repository-structure)
- [Requirements](#requirements)
- [Installation](#installation)
  - [Pandoc pipeline](#installation-pandoc-pipeline)
  - [LibreOffice macros](#installation-libreoffice-macros)
- [Usage](#usage)
  - [Converting a document with Pandoc](#converting-a-document-with-pandoc)
  - [Running macros inside LibreOffice](#running-macros-inside-libreoffice)
  - [Calling functions from your own scripts](#calling-functions-from-your-own-scripts)
- [Pandoc filter reference](#pandoc-filter-reference)
- [LibreOffice macro API reference](#libreoffice-macro-api-reference)
  - [`vistatype.utils`](#vistatypeutils)
  - [`vistatype.shared`](#vistatypeshared--sh_-commands)
  - [`vistatype.config`](#vistatypeconfig--ms_-commands)
  - [`vistatype.dn_tools`](#vistatypedn_tools--dn_-commands)
  - [`vistatype.lp.attach_template`](#vistatypelpattach_template)
  - [`vistatype.lp.file_cleanup`](#vistatypelpfile_cleanup)
  - [`vistatype.lp.reference_pages`](#vistatypelpreference_pages)
  - [`vistatype.lp.formatting_tools`](#vistatypelpformatting_tools)
  - [`vistatype.dx.reference_pages`](#vistatypedxreference_pages)
  - [`vistatype.dx.formatting_tools`](#vistatypedxformatting_tools)
  - [`vistatype.dx.file_cleanup`](#vistatypedxfile_cleanup)
  - [`vistatype.entrypoints`](#vistatypeentrypoints)
- [Complete command map](#complete-command-map)
- [Testing](#testing)
- [Technical notes for contributors](#technical-notes-for-contributors)
- [Known limitations & unfinished work](#known-limitations--unfinished-work)
- [License](#license)
- [Acknowledgments](#acknowledgments)
- [Contributing](#contributing)

---

## Why this is a reconstruction, not a translation

Three source files were inspected directly (unpacked and read on disk, not
just skimmed as PDFs) before any code was written:

**`Word.officeUI`** — a real ribbon/QAT customization file. It references
exactly **45 distinct commands**, extracted directly from its
`onAction="..."` attributes:
`Dx_*` (17, Braille/BANA), `Lp_*`/`LP_*` (17, Large Print), `Sh_*` (5,
shared), `MS_*` (3, Word configuration), `DN_*` (2, DAISY/NIMAS/text). This
list is the ground-truth command inventory for the whole project — see
[Complete command map](#complete-command-map).

**`LargePrintTemplate.dotx`** — a clean, macro-free Word template. No VBA
project at all. Converting it through headless LibreOffice
(`soffice --headless --convert-to odt`) and inspecting the result
byte-for-byte confirmed full fidelity of every style that matters:
`Normal` (Tahoma 18pt, +1pt letter-spacing, 0.25in space-after), `Heading 1`
(28pt bold), `Print Pg Num` (top/bottom border, `#F4A3D4` shading, right tab
stop), and 22 accessibility contrast styles (`Para*`/`Box*` are
**paragraph** styles; `Words*`/`Text*` are **character** styles — this
split is only visible in the raw XML's `w:type` attribute, not stated
anywhere in the guide text). Page geometry: Letter, 0.5in margins all
around. This file is used directly, unmodified in spirit, as both the
Pandoc `--reference-doc` source and the LibreOffice style-import source.

**`LPandBRL.dotm`** — the macro file itself, and the one with the problem.
Extracting its VBA with `oletools.olevba` produced a flagged warning:
**"Suspicious VBA Stomping"** — meaning the visible source code doesn't
match the compiled p-code Word actually executes. This was verified
concretely, not just taken on the tool's word for it:

1. Cleanly extracted all 45 real `VBA MACRO` code streams (filtering out
   ~190 `VBA FORM STRING` entries, which turned out to be garbled binary
   picture-resource data embedded in UserForms — not code).
2. Searched those 45 modules for the actual ribbon-callback entry points —
   `Dx_Attach_BANA_Template`, `Lp_Horz_List_To_Vertical`, `Sh_Doc_Info`, all
   45 names from `Word.officeUI`. **None of them exist anywhere in the
   extracted source.**
3. Listed every `Sub`/`Function` name that IS present: **100% of them are
   UserForm button-click handlers**, each just calling out to a
   *differently named* worker macro via `Application.Run
   MacroName:="..."` — dynamic, string-based dispatch, so even the visible
   click-handlers don't reference the real logic by a name a static search
   can follow.

Net effect: the actual implementations of the ~150 real macros this add-in
runs are **not recoverable from this file**. What's readable is dialog/form
scaffolding — genuinely useful for confirming exact option wording and
structure (e.g. the Horizontal-List-to-Vertical dialog's three list-type
radio buttons, or the Dashes/Primes/Fractions button-grid layout), but no
algorithm bodies.

This isn't necessarily evidence of malicious intent — VBA source-protection
tools that deliberately strip visible source while leaving compiled p-code
intact produce exactly this same signature, and plenty of legitimate
freeware authors use them to protect their work from casual copying. But it
does mean this project had to be built as a **reconstruction from the two
PDF user guides** (a 43-page BANA Macros guide and a 162-page VistaType LP
guide, both read in full), not a line-by-line port — a meaningfully
different and more judgment-heavy undertaking. Full technical detail is in
[`docs/FINDINGS.md`](docs/FINDINGS.md).

One concrete consequence, documented because it matters: an early draft of
the Braille spelling-list formatter invented a plausible-looking DBT
(Duxbury Braille Translator) bracket code for "render this uncontracted,"
by analogy with the *real* codes the guide does quote verbatim elsewhere
(`$pg`, the continuation-page code, the letter-sign code, the fraction
codes). On review, the guide never actually states literal syntax for that
one, so the invented code was removed before shipping — see
[`vistatype.dx.file_cleanup.format_spelling_list`](#vistatypedxfile_cleanup)
and `docs/FINDINGS.md`'s last section. Getting a DBT bracket code wrong
isn't a cosmetic bug: it risks producing genuinely incorrect braille for a
blind reader. Where this project doesn't know something, it says so in the
docstring rather than guessing.

---

## Repository structure

```
reference-doc/
  build_reference_odt.sh          Rebuilds the reference .odt from a .dotx
  LargePrintTemplate-reference.odt Pre-built, verified reference doc (checked in)

pandoc/
  filters/                        Lua filters for the batch conversion pipeline
    lp-fix-common-file-errors.lua
    lp-delete-multiple-paragraph-marks.lua
    lp-horizontal-list-to-vertical.lua
    lp-tag-reference-page-numbers.lua
    lp-format-tagged-page-numbers.lua
  test/
    smoke_test.sh                 Runnable example against small fixtures
    *.md                          Fixture files

lo-macros/
  vistatype/                      <-- install THIS directory into LibreOffice
    utils.py                      Shared UNO helpers (selection/paragraph handling)
    shared.py                     Sh_ commands (shared between Lp_/Dx_)
    config.py                     MS_ commands (view/editing configuration)
    dn_tools.py                   DN_ commands (DAISY/NIMAS/text-file tools)
    entrypoints.py                XSCRIPTCONTEXT wrappers -- the actual macro entry points
    lp/                           Lp_/LP_ commands (Large Print)
      attach_template.py
      file_cleanup.py
      reference_pages.py
      formatting_tools.py
    dx/                           Dx_ commands (Braille/BANA)
      reference_pages.py
      formatting_tools.py
      file_cleanup.py
  tests/
    uno_harness.py                Minimal headless-soffice connection helper
    test_*.py                     One (or more) test file per module
    run_all_tests.sh              Runs the whole suite

docs/
  FINDINGS.md                     Full technical writeup of the VBA-stomping discovery
  COMMAND_MAP.md                  All 45 commands x status x implementation location, in detail
```

---

## Requirements

- **LibreOffice** ≥ 7.x (developed and tested against 24.2). Needed for:
  the reference-doc build script, running the macros interactively, and
  the LibreOffice-based test suite.
- **Python** 3.9+ with the `python-uno` bridge (`import uno`). This ships
  with LibreOffice on Linux (`/usr/lib/python3/dist-packages/uno.py` or
  similar) and with the LibreOffice installation on macOS/Windows — you do
  not need a separate install, just make sure the Python that runs the
  test suite is the same one (or is on the same `sys.path`) as LibreOffice's
  own bundled Python, or that `uno` is otherwise importable.
- **Pandoc** ≥ 2.19 (developed and tested against 3.1.3; needs Lua filter
  support, which has been standard for a long time — `pandoc --version`
  should show `+lua`).
- **`python-docx`** (`pip install python-docx`) — only needed to run the
  test suite's synthetic `.docx` fixtures; not a runtime dependency of the
  package itself.

Nothing here requires Microsoft Word, a Windows machine, or any commercial
software.

---

## Installation

### Installation: Pandoc pipeline

The pipeline needs a reference `.odt` built from a real copy of
`LargePrintTemplate.dotx`. A pre-built, verified one is already checked in
at `reference-doc/LargePrintTemplate-reference.odt` — you only need to
rebuild it if you have a different or newer template file:

```bash
cd reference-doc
./build_reference_odt.sh /path/to/LargePrintTemplate.dotx LargePrintTemplate-reference.odt
```

The script converts the `.dotx` through headless LibreOffice and then
verifies (prints OK/MISSING for each) that the key styles and page geometry
survived the conversion — `Normal`, `Print Pg Num`, `Para Black`,
`Words Black`, `Box Black`, `Box Red`, `List Paragraph`, and the page
size/margins. No further installation step is needed for the Pandoc side;
the Lua filters are invoked directly by path (see
[Usage](#converting-a-document-with-pandoc)).

### Installation: LibreOffice macros

Copy (or symlink) the `vistatype` package into LibreOffice's user Python
scripts directory:

| Platform | Path |
|---|---|
| Linux | `~/.config/libreoffice/4/user/Scripts/python/` |
| macOS | `~/Library/Application Support/LibreOffice/4/user/Scripts/python/` |
| Windows | `%APPDATA%\LibreOffice\4\user\Scripts\python\` |

```bash
cp -r lo-macros/vistatype ~/.config/libreoffice/4/user/Scripts/python/
```

You should end up with, e.g.,
`.../Scripts/python/vistatype/entrypoints.py`,
`.../Scripts/python/vistatype/lp/attach_template.py`, etc.

Restart LibreOffice. **Tools ▸ Macros ▸ Organize Macros ▸ Python** should
now show `vistatype.entrypoints` with every command listed in
[`vistatype.entrypoints`](#vistatypeentrypoints) below, ready to run
directly or bind to a toolbar button / menu item / keyboard shortcut via
**Tools ▸ Customize**.

---

## Usage

### Converting a document with Pandoc

```bash
pandoc input.docx -o output.odt \
    --reference-doc=reference-doc/LargePrintTemplate-reference.odt \
    --lua-filter=pandoc/filters/lp-fix-common-file-errors.lua \
    --lua-filter=pandoc/filters/lp-delete-multiple-paragraph-marks.lua \
    --lua-filter=pandoc/filters/lp-tag-reference-page-numbers.lua \
    --lua-filter=pandoc/filters/lp-format-tagged-page-numbers.lua \
    --lua-filter=pandoc/filters/lp-horizontal-list-to-vertical.lua
```

Works identically for a Markdown source (`input.md` instead of
`input.docx`). Filter order matters somewhat:

1. `lp-fix-common-file-errors.lua` first, so stray whitespace doesn't
   interfere with later pattern matching.
2. `lp-delete-multiple-paragraph-marks.lua` next.
3. `lp-tag-reference-page-numbers.lua` **before**
   `lp-format-tagged-page-numbers.lua` — the latter only acts on
   paragraphs the former has already tagged with `$pg`.
4. `lp-horizontal-list-to-vertical.lua` can run any time after step 1
   (it doesn't depend on or interfere with the others).

Only include the filters you actually want; each is independent and safe
to omit. See [Pandoc filter reference](#pandoc-filter-reference) for full
detail on each one, including exactly what each does and doesn't cover.
Run `pandoc/test/smoke_test.sh` for a working example against small
fixture files.

### Running macros inside LibreOffice

Once installed (see above), open **Tools ▸ Macros ▸ Organize Macros ▸
Python**, navigate to `vistatype.entrypoints`, and either run a function
directly or assign it to a toolbar button / menu entry / keyboard shortcut
via **Tools ▸ Customize**. Every entry point in
[`vistatype.entrypoints`](#vistatypeentrypoints) corresponds 1:1 to an
original ribbon command name (e.g. `Lp_AutoTag_Page_Numbers`,
`Dx_Compress_Linear_Math`), so you can reproduce the original's ribbon
layout in LibreOffice's own toolbar/menu customization UI if you want a
visually similar experience.

Most of these operate the same way the originals did: for a
cursor-position–based macro (title case, compress linear math, keep with
next paragraph, ...), place your cursor in the target paragraph first; for
a selection-based one (Format Exercise Levels, Selected File Cleanup),
select the relevant paragraph(s) first, or leave nothing selected to just
affect the current paragraph.

### Calling functions from your own scripts

Every real function (as opposed to the thin `entrypoints.py` wrappers)
takes an explicit `doc` argument and returns a plain Python value, so you
can call it directly from any Python process connected to LibreOffice over
UNO — a batch script, your own automation, or the test suite. This is
also, not coincidentally, exactly how every function in this project was
verified during development: see `lo-macros/tests/uno_harness.py` for the
~20-line connection helper, and any `lo-macros/tests/test_*.py` file for a
complete, runnable example.

```python
import sys
sys.path.insert(0, "lo-macros/tests")
sys.path.insert(0, "lo-macros")
from uno_harness import LOSession
from vistatype.lp import reference_pages

with LOSession() as lo:
    doc = lo.open("/path/to/document.odt")
    tagged_count = reference_pages.auto_tag_page_numbers(doc)
    print(f"Tagged {tagged_count} reference page numbers")
    doc.store()
    doc.close(False)
```

See the full [LibreOffice macro API reference](#libreoffice-macro-api-reference)
below for every function's signature, parameters, return value, and
behavior.

---

## Pandoc filter reference

Every filter here is intentionally scoped to **batch-safe** operations —
things that are safe to run unattended over an entire document with no
human choosing which paragraph is meant. Anything from the original that's
inherently interactive (place your cursor here, select this text, choose
an option from a dialog) is a LibreOffice macro instead — see
[LibreOffice macro API reference](#libreoffice-macro-api-reference).

### `lp-fix-common-file-errors.lua`

Ports part of **Fix Common File Errors** / `Lp_File_Fix_Sequence` (step 1).

Applies, across every paragraph in the document:
- Collapses runs of repeated hyphen / en dash / em dash characters (each
  dash type collapses only with itself) down to a single instance —
  including inside a mixed alphanumeric token like `"triple---dash"` →
  `"triple-dash"`.
- Strips a space immediately before or after a token that starts or ends
  with one of those dash characters.
- Collapses runs of literal space characters inside a single text run.
- Converts `SmallCaps`-styled text to plain, real upper-case text.
- Strips leading/trailing space/tab from each paragraph.
- Converts a straight `'` or `"` immediately following a digit into a
  proper prime (`′`, U+2032) / double prime (`″`, U+2033) mark.

**Not covered here** (needs a different layer — see
[`vistatype.lp.file_cleanup.fix_common_file_errors`](#vistatypelpfile_cleanup)
instead, which has real access to LibreOffice's document object model):
Abbyy FineReader style remapping, drop-cap reset (OOXML drop caps aren't
modeled in Pandoc's AST at all), un-boxing text frames while preserving
text, confirmable inline-image removal.

### `lp-delete-multiple-paragraph-marks.lua`

Ports **Delete Multiple Paragraph Marks** / `Lp_File_Fix_Sequence` (step
2). Collapses runs of consecutive empty paragraphs down to a single one.

**Note verified during development:** for `.docx` input specifically, this
is close to a no-op — Pandoc's own DOCX reader silently drops fully-empty
paragraphs *before the AST even exists*, so there's usually nothing left
for this filter to collapse. It matters for ODT input and for Markdown's
own way of representing an intentionally-blank paragraph (a run of only
spaces or a non-breaking space). Kept in the default pipeline as a safe,
cheap, harmless step regardless of input format.

### `lp-horizontal-list-to-vertical.lua`

Ports **Horizontal List to Vertical** / `Lp_Horz_List_To_Vertical`, in its
"Ordered List" mode only.

Detects a paragraph containing **3 or more** sequential same-family
markers (`a.`/`b.`/`c.`, or `1)`/`2)`/`3)`, etc.) and splits it into one
paragraph per item, sorted ascending by marker value. Example:

```
Input:  a. the Sparrow b. the Fly c. the Fish d. the Beetle
Output: a. the Sparrow
        b. the Fly
        c. the Fish
        d. the Beetle
```

The 3-marker minimum (rather than 2) is a deliberate false-positive
safeguard, confirmed necessary during development: 2 sequential markers
("I went to a) the store and b) the bank today") occur often enough in
ordinary prose that auto-converting on a 2-marker trigger produced wrong
output on a real test case; 3+ same-family sequential markers on one line
essentially never occurs outside a genuine horizontal exercise list.

Also handles the case where Pandoc's own Markdown reader has already
consumed the paragraph's leading marker as an `OrderedList`'s own start
number (e.g., a line beginning with `"a. "` in Markdown parses as a
1-item ordered list, not a plain paragraph) — this was a real, confirmed
data-loss bug during development (the text between the consumed marker and
the next recognized one was silently dropped) and is now handled by
reconstructing the implied marker from the list's own `start`/`style`/
`delimiter` attributes before running the same splitting logic.

**"Spaced List"** (bare words separated by spaces) and **"Tabbed List"**
(segments separated by real tab characters) modes from the original are
**not** in this filter — auto-detecting those over a whole document would
misfire constantly on ordinary prose (which *is* words separated by
spaces). Use
[`vistatype.lp.formatting_tools.horizontal_list_to_vertical`](#vistatypelpformatting_tools)
with an explicit `mode="spaced"` or `mode="tabbed"` instead, which — like
the original — operates on a paragraph you've selected.

### `lp-tag-reference-page-numbers.lua`

Ports **Auto Tag Page Numbers** / `Lp_AutoTag_Page_Numbers`.

A paragraph whose **entire** content (and nothing else) is a recognizable
reference page number gets `$pg` prepended. Recognized shapes:

| Shape | Example | Tagged as |
|---|---|---|
| Plain Arabic | `12` | `$pg12` |
| Letter + Arabic | `A12`, `12C`, `A12B` | `$pgA12`, `$pg12C`, `$pgA12B` |
| Continuation range | `13-15` | `$pg13-15` |
| Roman numeral (I–CCC / 1–300) | `XIV` | `$pgXIV` |

A page number embedded within, or at the start/end of, a paragraph of
running text is left alone — matching the guide's own stated rule.

### `lp-format-tagged-page-numbers.lua`

Ports **Format Page Numbers** / `Lp_Format_Page_Numbers`. Must run after
`lp-tag-reference-page-numbers.lua` (or after tags have been added by
hand).

Finds every paragraph starting with `$pg`, strips the tag, and wraps the
paragraph in a `Div` with `custom-style="Print Pg Num"` — a mechanism
Pandoc's ODT writer maps directly onto the real named style from the
reference document. **Verified against a real LibreOffice instance**, not
just assumed: the resulting paragraph resolves to the actual imported
"Print Pg Num" style, background colour `#F4A3D4` and all, not a
same-named-but-empty placeholder (see
`lo-macros/tests/test_reference_pages.py` for the check, and the filter's
own header comment for a link to how this was confirmed).

---

## LibreOffice macro API reference

Every module below documents **every public function** — signature,
parameters, return value, and behavior. Private helpers (leading
underscore) are omitted; read the source if you need those. Every function
listed as tested has a corresponding assertion in
`lo-macros/tests/test_*.py`, run against a real headless LibreOffice
instance — none of this is "written and assumed to work."

Design convention used throughout: every function takes an explicit `doc`
(a `com.sun.star.text.TextDocument`) rather than reaching for a global, and
most take an optional `target` (a specific paragraph object) that defaults
to the current paragraph (`vistatype.utils.get_current_paragraph(doc)`) if
omitted — matching the original macros' "place the cursor in the
paragraph, no need to select it" convention.

### `vistatype.utils`

Shared UNO helpers used by every other module.

#### `get_active_document()`
Returns the document LibreOffice itself is currently running a script
against. **Only works when invoked by LibreOffice's own scripting
provider** (i.e., from inside `entrypoints.py`-style code, where the
`XSCRIPTCONTEXT` global exists). Raises a clear `RuntimeError` (not a
confusing `NameError` from deep in a call stack) if called from an
external UNO connection — pass `doc` explicitly there instead.

#### `make_prop(name, value)`
Builds a `com.sun.star.beans.PropertyValue` — the small struct UNO uses
pervasively for named-argument-style API calls (`Name`, `Value`).

#### `iter_paragraphs(doc)`
Generator yielding every top-level paragraph (`com.sun.star.text.Paragraph`)
in the document body, in order. Does not descend into tables or text
frames — walk `doc.TextTables` / `doc.TextFrames` separately for those.

#### `get_current_paragraph(doc)`
Returns the paragraph object containing the view cursor (the user's
caret), or `None` if the document has no controller (e.g., opened
`Hidden`). This is how every "place your cursor in the paragraph" macro
finds its target when no `target=` is passed explicitly.

#### `has_selection(doc)`
Returns `True` if the view cursor currently spans a non-empty selection.

#### `get_selected_paragraphs(doc)`
Returns a list of every paragraph touched by the current selection (using
a correctly-oriented interval-overlap test against collapsed point
cursors — see [Technical notes](#technical-notes-for-contributors) for a
real sign-convention bug this caught during development), or just
`[current_paragraph]` if nothing is selected, or `[]` if there's no
controller and no selection.

#### `set_break_before_page(paragraph, on: bool)`
Sets or clears a page-break-before on `paragraph` directly (as opposed to
the toggle wrapper in `vistatype.shared`).

#### `ensure_paragraph_style(doc, name: str, based_on: str = "Standard", props: dict | None = None)`
Get-or-create a paragraph style by name: returns the existing style if
`name` is already in the document's `ParagraphStyles` family, otherwise
creates a new one parented to `based_on` and applies any properties in
`props`. Used so styles like `Print Pg Num` or `List`/`List 2` work even
on a document that never had the real template attached.

#### `show_message_box(doc, message: str, title: str = "VistaType LP", buttons="ok")`
Pops a real UNO message box (`"ok"`, `"yesno"`, or `"yesnocancel"`),
returning `"OK"`/`"CANCEL"`/`"YES"`/`"NO"`. **Requires a real window
toolkit** — raises in a plain `--headless --invisible` soffice instance
with no window server, which is why every function that would otherwise
pop this dialog (e.g. confirmable image removal) accepts an explicit
`confirm=`/`on_confirm=` argument to bypass it entirely for unattended or
automated use.

### `vistatype.shared` — `Sh_` commands

Commands shared between the Large Print and Braille sides.

#### `apply_title_case(doc, target=None)`
*Ports `Sh_Apply_Title_Case_Capitalization`.* Capitalizes each word in the
target paragraph except a fixed list of small words (`a`, `an`, `and`,
`as`, `at`, `but`, `by`, `for`, `if`, `in`, `nor`, `of`, `on`, `or`, `per`,
`so`, `the`, `to`, `up`, `vs`, `via`, `yet`, `with`, `from`, `into`, `onto`,
`than`) — matching the guide's own wording as precisely as it's stated.
The one addition beyond the guide's literal text: the very first word of
the paragraph is always capitalized, even if it's on the small-words list,
so a title never starts lowercase. Deliberately does **not** implement
further conventions some style guides use (capitalizing the last word, or
the word after a colon) since those aren't evidenced by the guide and
would be guessing, not porting. Returns the modified paragraph object, or
`None` if there's no target paragraph.

#### `toggle_keep_lines_together(doc, target=None)`
*Ports `Sh_Keep_Lines_Of_Para_Together`.* Toggles the paragraph's
`ParaSplit` property. When `ParaSplit` is `False`, LibreOffice will not
split the paragraph's lines across a page boundary — this is Word's "keep
lines together." Returns the new "kept together" state (`True`/`False`),
or `None` if there's no target.

#### `toggle_move_to_next_page(doc, target=None)`
*Ports `Sh_Move_Paragraph_To_Next_Page`.* Toggles a page-break-before on
the paragraph (`BreakType.PAGE_BEFORE` ↔ `NONE`). Unlike the original
(a one-shot "push to next page"), this is a toggle, so re-running it
undoes the break — more forgiving for a script-driven workflow. Returns
`True` if a break was just added, `False` if just removed.

#### `char_value_at_cursor(doc)`
*Ports `Sh_Show_Char_Val`.* Reports the character immediately to the right
of the view cursor as a dict: `{"char": <str>, "decimal": <int>,
"hex": "U+XXXX"}`. Returns `None` if there's no controller or no next
character. Note: Word's ANSI/legacy-code-page framing doesn't map onto
LibreOffice, which is Unicode-native throughout with no such layer to
inspect — this reports what actually matters for LibreOffice's own
Find & Replace (which accepts `\uXXXX`-style regex, not ANSI codes).

#### `doc_info(doc)`
*Ports `Sh_Doc_Info`.* Returns a dict describing the document's current
page style and (if there's a controller) current paragraph style:
```python
{
    "page_style": "Standard",
    "width_in": 8.5, "height_in": 11.0, "orientation": "Portrait",
    "top_margin_in": 0.5, "bottom_margin_in": 0.5,
    "left_margin_in": 0.5, "right_margin_in": 0.5,
    "current_paragraph_style": "Standard",         # only if there's a controller
    "has_large_print_styles": True,                # proxy for "was attach_lp_template() run?"
}
```
LibreOffice documents don't carry a Word-style "AttachedTemplate" pointer
once they're native `.odt`, so instead of faking one, this reports
`has_large_print_styles` (whether the `Print Pg Num` style is present) as
a meaningful proxy for the same underlying question the original macro
answered.

### `vistatype.config` — `MS_` commands

#### `set_formatting_marks_visible(doc, visible: bool = True)`
Toggles View ▸ Formatting Marks (spaces, tabs, paragraph marks, etc.) —
per-document, via `ViewSettings.ShowNonprintingCharacters`. Returns
`visible`, or `None` if there's no controller.

#### `set_default_paragraph_style(doc, style_name: str)`
If the document's very first paragraph is blank and currently uses
`Standard`/`Default Paragraph Style`, switches it to `style_name` (as long
as that style exists in the document). Returns `True`/`False` for whether
`style_name` exists at all.

#### `set_config_new_install(doc)`
*Ports `MS_Change_Word_Configuration_to_New_Install`.* Plain-editing
profile: formatting marks off. Returns `{"formatting_marks": False}`.

#### `set_config_large_print(doc)`
*Ports `MS_Change_Word_Configuration_to_Large_Print`.* Formatting marks
on, default paragraph style set to `Standard`. Returns
`{"formatting_marks": True, "default_style": "Standard"}`.

#### `set_config_braille(doc)`
*Ports `MS_Change_Word_Configuration_to_Braille`.* Formatting marks on
(BANA documents lean on visible `$pg`/DBT bracket-code tags, much easier
to audit with marks shown). Returns `{"formatting_marks": True}`.

> **What these deliberately don't do:** the original three QAT icons also
> swap a whole Word AutoCorrect/AutoFormat-as-you-type profile. That's an
> **application-wide** LibreOffice preference (Tools ▸ AutoCorrect
> Options), not a per-document one, so a per-document macro silently
> changing it on your behalf would be surprising and hard to reason about.
> See the module's own docstring for the full reasoning.

### `vistatype.dn_tools` — `DN_` commands

Pure Python text/XML transforms — no LibreOffice document required, no UNO
dependency. Operate on plain strings, so they work identically whether or
not LibreOffice is even running.

#### `add_pg_tags_to_xml(xml_text: str) -> str`
*Ports `DN_Add_PgNo_Tags_To_DAISY_or_NIMAS`.* Finds NIMAS/DAISY `<pagenum>`
elements (both `<pagenum>12</pagenum>` and `<pagenum value="12"/>` forms)
and prepends `$pg` to the page-number text inside, using the exact same
pattern-classification rules as
[`vistatype.lp.reference_pages`](#vistatypelpreference_pages), so a
tagged NIMAS file and a tagged Word/LO document use one consistent
convention downstream. Already-tagged elements (starting with `$pg`) are
left untouched.

#### `fix_text_file_paragraphs(text: str) -> str`
*Ports `DN_Remove_Para_Formatting_From_Text_Files` ("Txt File Para
Fix").* Removes within-paragraph line breaks from a hard-wrapped text file
(one that has a line break at the end of every line and a blank line
between real paragraphs), re-joining each paragraph's lines with a single
space while preserving blank-line-delimited paragraph boundaries.
Normalizes `\r\n`/`\r` to `\n` first. Matches the guide's own Tom
Sawyer-excerpt example exactly.

### `vistatype.lp.attach_template`

#### `import_styles(doc, template_path=None, family_names=None, desktop=None)`
Copies every settable property of every style in `family_names` (default:
`["ParagraphStyles", "CharacterStyles", "PageStyles"]`) from the template
file at `template_path` (default: the bundled
`reference-doc/LargePrintTemplate-reference.odt`) into `doc`, creating
each style in `doc` first if it doesn't already exist. Returns a nested
dict: `{family_name: {style_name: (properties_copied, properties_failed)}}`
(a handful of "failed" per style is normal and expected — those are
read-only/computed properties like `CharDiffHeight`, not a sign of a real
problem).

**`desktop` is required** and must come from the same UNO connection as
`doc` — see [Technical notes](#technical-notes-for-contributors) for why a
fresh `uno.getComponentContext()` call here caused a segfault during
development, and why this function refuses to create one for you instead
of silently doing the wrong thing.

#### `attach_lp_template(doc, template_path=None, desktop=None)`
*Ports `LP_Attach_LP_Template`* (the style/page-setup portion only — see
below). Calls `import_styles(...)` for all three families, then explicitly
sets the first paragraph's `PageDescName` to `"Standard"` so the imported
page geometry actually takes effect immediately. After this call, `doc`
has `Standard` (18pt Tahoma + letter-spacing), `Print Pg Num`, and all
22 `Para*`/`Words*`/`Text*`/`Box*` accessibility styles, at Letter size
with 0.5in margins — verified byte-for-byte against the source template
in `lo-macros/tests/test_attach_template.py`. Returns the same dict
`import_styles` returns.

**Not covered:** the original also includes a paper/screen-size-and-
binding-and-media-selection wizard (target paper size, mirroring, print
vs. screen output profiles) — a substantial, separate feature in its own
right, not built yet. See
[Known limitations](#known-limitations--unfinished-work).

### `vistatype.lp.file_cleanup`

#### `fix_common_file_errors(doc, remove_images=None, on_confirm=None)`
*Ports `Lp_File_Fix_Sequence` step 1 ("Fix Common File Errors").* Across
the whole document:
- dash-run collapsing + space-stripping around dashes, multi-space
  collapsing, digit+quote → prime/double-prime (same rules as the Pandoc
  filter, applied to real Python strings — no AST token games needed)
- small caps → real upper-case text
- drop caps reset to normal text (`DropCapFormat.Count`/`.Lines` zeroed)
- text frames "un-boxed": each frame's text is appended as a new paragraph
  at the end of the document (position is approximate — UNO doesn't
  expose an unambiguous "insert exactly where the frame was anchored" —
  review placement afterward), then the frame is removed
- inline images removed, **with confirmation**: if `remove_images` is
  `None` (the default), pops a real message box ("Found N images. Delete
  all?") via `on_confirm(image_count)` if given, or
  `utils.show_message_box` otherwise. Pass an explicit `True`/`False` to
  skip the prompt for unattended/automated use (required in a pure
  headless session, which has no window server for the message box).

Returns `{"frames_removed": int, "images_found": int, "images_removed": int}`.

**Not covered** (needs a different layer entirely, or isn't recoverable —
see the module's own docstring): Abbyy FineReader style remapping.

#### `delete_multiple_paragraph_marks(doc)`
*Ports `Lp_File_Fix_Sequence` step 2.* Collapses runs of consecutive empty
paragraphs down to one, across the whole document (always whole-document,
not selection-scoped — provably safe and equivalent to a scoped version
for this specific operation; see the function's own docstring). Returns
the number of paragraphs actually deleted.

**Implementation note, load-bearing for anyone modifying this function:**
re-enumerates paragraphs fresh before every single deletion rather than
collecting them all up front and batch-deleting. Confirmed necessary
during development — batching caused a real `UnknownPropertyException`
("cannot get value Start") on the second deletion, because the first
deletion's merge invalidated UNO's paragraph proxy for anything positioned
after it in the document.

#### `file_fix_sequence(doc, remove_images=None, on_confirm=None)`
*Ports `Lp_File_Fix_Sequence`* (the full two-step sequence). Runs
`fix_common_file_errors` then `delete_multiple_paragraph_marks`, returning
the first call's result dict with an added `"empty_paragraphs_collapsed"`
key.

#### `selected_cleanup(doc, options=None)`
*Ports `Lp_Selected_File_CleanUp`.* Same class of fixes as
`fix_common_file_errors`, scoped to the current selection (or current
paragraph). `options` is a set drawn from `{"dashes_spaces_primes",
"small_caps", "multi_para_marks"}` (default: all three). Returns
`{"paragraphs_processed": int, "empty_paragraphs_collapsed": int}`.

### `vistatype.lp.reference_pages`

A "reference page number" is a page number from the *original print* book,
preserved in the transcription for cross-referencing — not a page number
of the transcription itself.

#### `auto_tag_page_numbers(doc)`
*Ports `Lp_AutoTag_Page_Numbers`.* Scans every paragraph; any paragraph
whose entire text is (and only is) a recognizable reference page number
(plain Arabic, letter+Arabic, a continuation range, or a Roman numeral
I–CCC) gets `$pg` prepended. Already-tagged paragraphs are skipped.
Returns the count tagged.

#### `manual_tag_page_number(doc, target=None)`
*Ports `Lp_Manual_Tag_with_Dollar_pg`.* Tags just the target paragraph
(current paragraph by default) with `$pg`, regardless of whether it
matches the auto-detect pattern — for numbers the automatic pass missed or
that don't fit a recognized shape. Returns the paragraph object, or `None`.

#### `format_tagged_page_numbers(doc)`
*Ports `Lp_Format_Page_Numbers`.* Finds every paragraph starting with
`$pg`, strips the tag, and applies the `Print Pg Num` paragraph style —
creating that style on the fly (with the real template's exact property
values: `#F4A3D4` background) if the document doesn't already have it, so
this works even without having run `attach_lp_template()` first. Returns
the count formatted.

#### `validate_tagged_page_numbers(doc)`
*Ports `Lp_Validate_Dollar_PG`.* Returns a list of dicts, one per tagged
paragraph:
```python
{"index": 1, "text": "$pg12", "value": "12", "kind": "arabic", "in_order": True}
```
`in_order` flags any entry whose numeric value isn't ≥ the previous
tagged entry's — a signal for "this might be a false positive or an OCR
misread." Unlike the original (which opens a scratch document with the
tags listed in a column for visual scanning), this returns structured
data so the *caller* decides how to present it — `entrypoints.py`'s
wrapper opens a real scratch Writer window with the same list, matching
the original's workflow, when invoked as an actual LibreOffice macro.

### `vistatype.lp.formatting_tools`

#### `horizontal_list_to_vertical(doc, mode="auto", target=None, sort_ascending=True)`
*Ports `Lp_Horz_List_To_Vertical`.* `target` may be a single paragraph, a
list of paragraphs (e.g. from `utils.get_selected_paragraphs(doc)`), or
omitted (defaults to the current paragraph). `mode`:
- `"ordered"` — letters/numbers (`a.`, `b)`, `12.`) mark each item;
  requires ≥3 markers, same false-positive safeguard as the Pandoc filter.
- `"spaced"` — single words separated by one-or-more spaces, no markers.
- `"tabbed"` — words/word-groups separated by real tab characters.
- `"auto"` (default) — tries `"ordered"` only; `"spaced"`/`"tabbed"`
  require you to say so explicitly, since — unlike a marker run — spacing
  or tabs alone aren't a safe auto-detect signal.

Returns the list of paragraphs that were converted, or `None` if nothing
matched (originals left untouched).

#### `compress_linear_math(doc, target=None)`
*Ports `Lp_Compress_Linear_Math`.* Removes all normal spaces from the
target paragraph, then inserts a thin space (U+2009) immediately before
and after every comparison symbol (`=`, `≈`, `<`, `>`, `≠`, `≤`, `≥`) —
exactly the guide's wording ("removes all normal spaces then places very
thin (hair spaces) between the math and symbols of comparison"). Returns
the paragraph object, or `None`.

> ⚠️ **This is a genuinely different rule from the Braille side's function
> of the same name** — see
> [`vistatype.dx.formatting_tools.compress_linear_math`](#vistatypedxformatting_tools),
> which *keeps* normal spaces around comparison operators and removes
> everywhere else. Both are implemented faithfully to their own guide's
> wording; the two guides simply specify different behavior for a macro
> with the same name across the two products.

#### `toggle_space_after_para(doc, target=None)`
*Ports `Lp_Toggle_Space_After_Current_Para`.* Closes the gap between a
titling paragraph and the one below it by zeroing `ParaBottomMargin`;
running it again restores the value defined by the paragraph's own style
(with a ~0.25in fallback if that's also zero). Returns `True` if space was
just removed, `False` if just restored. Note: this is a **stateless**
toggle — it restores to the *style's* value, not necessarily this specific
paragraph's own prior direct-formatting value, if those ever differed.

#### `toggle_keep_with_next_para(doc, target=None)`
*Ports `Lp_Keep_With_Next_Para`.* Toggles `ParaKeepTogether`, which keeps
this paragraph glued to the one immediately following it across a page
break. **Distinct from** `vistatype.shared.toggle_keep_lines_together`
(which toggles `ParaSplit` — keeping *this paragraph's own lines* from
splitting, not gluing it to the next paragraph). Returns the new state.

#### `type_fill_in_line(doc, length=8, to_right_margin=False)`
*Ports `Lp_Type_Fill_In_Line`.* Inserts a fill-in line (underlined
underscores) at the cursor.
- Fixed-length (`length` 1–21, matching the guide's 21 buttons): inserts
  a leading space then `length` underlined underscore characters.
- `to_right_margin=True`: adds a right-aligned paragraph tab stop at the
  page's right text margin with `_` as its **fill character**, then
  inserts a real tab — LibreOffice's native "leader-filled run to a tab
  stop" mechanism. This is a different underlying mechanism from Word's
  (unrecoverable) internal approach, chosen because it achieves the same
  visual/functional result: a fill-in line that reaches the margin
  regardless of how much text precedes it, verified against a live
  document.

Returns `None` either way (the insertion is the side effect).

#### `format_exercise_levels(doc, target_paragraphs=None)`
*Ports `Lp_Format_Exercise_Lv_1_and_Lv_2`.* Operates on
`target_paragraphs` (default: the current selection). A paragraph starting
with a number + period/space is a question, styled `"List"`; anything else
in the selection is an answer choice, styled `"List 2"` (both created on
the fly, based on `Standard`, if the document doesn't already have them).
Tabs and lone underscores within the selection are replaced with
underlined 8-underscore fill-in runs (comma-joined for a run of more than
one tab/underscore in immediate sequence), matching the guide's described
OCR-cleanup behavior. Returns `{"questions": int, "answers": int}`.

### `vistatype.dx.reference_pages`

Same detection rules as the Large Print side
([`vistatype.lp.reference_pages`](#vistatypelpreference_pages)), plus DBT
(Duxbury Braille Translator) markup the Braille side needs on top.

#### `auto_tag_page_numbers(doc, translation: str = "UEB")`
*Ports `Dx_AutoTag_Page_Numbers`.* `translation` is `"UEB"` or `"EBAE"` —
changes lowercase-roman-numeral handling:
- **Continuation ranges** (`"13-15"`) get the DBT continuation-page code
  appended after the range's end value, exactly per the guide's own
  example: `$pg13-15[[*lec*]]15`.
- **Lowercase roman numerals**: under UEB, only the ambiguous single
  letters `c`, `l`, `v` get the DBT letter-sign code appended
  (`$pgc[[*ii*]]`); under EBAE, **every** lowercase roman numeral does
  (`$pgxiv[[*ii*]]`).

Both DBT codes (`[[*lec*]]`, `[[*ii*]]`) are quoted **verbatim** from the
BANA guide, not invented — see
[Why this is a reconstruction](#why-this-is-a-reconstruction-not-a-translation).
Returns the count tagged.

#### `manual_tag_page_number(doc, target=None, translation: str = "UEB")`
*Ports `Dx_Manual_Tag_with_Dollar_pg`.* Same DBT-code logic as
`auto_tag_page_numbers`, applied to a single target paragraph regardless of
whether it matches the auto-detect shape.

#### `format_tagged_page_numbers(doc)`
*Ports `Dx_Format_Tagged_Page_Numbers`.* Same idea as the LP side: strips
`$pg`, leaves any embedded DBT codes as literal text (DBT reads those
directly from the document), applies a paragraph style — here,
`"BANA Braille Template Reference Page Number"`, created on the fly with a
**placeholder** appearance (`#D9D9D9` grey, distinct from the LP side's
pink, so it's visually obvious in a mixed workflow that this is a
placeholder). ⚠️ We don't have a real BANA template file to draw the
actual formatting from — only `LargePrintTemplate.dotx` was provided. See
[Known limitations](#known-limitations--unfinished-work).

#### `validate_tagged_page_numbers(doc)`
*Ports `Dx_Ref_Pg_Number_Sequence_Menu`.* Same shape as the LP side's
version; returns a list of `{"index", "text", "value"}` dicts.

### `vistatype.dx.formatting_tools`

#### `compress_linear_math(doc, target=None)`
*Ports `Dx_Compress_Linear_Math`.* Removes **all** spaces from the target
paragraph's linear math **except** those immediately preceding/following a
comparison symbol (`=`, `≈`, `<>`, `<`, `>`, `≠`, `≥`, `≤`), which are left
as ordinary spaces — the Braille-side rule verbatim from the guide
("removes all spaces from linear math except for those preceding and
following signs of comparison"). See the LP-side function's docstring for
the explicit side-by-side with that different rule.

#### `insert_dash(doc, kind: str)`
*Part of `Dx_Type_Dashes`.* Inserts one of: `"long_dash"` (─, horizontal
bar), `"em_dash"`, `"en_dash"`, `"minus_sign"` (−), `"hyphen"` (Unicode
hyphen, not ASCII `-`), `"nonbreaking_hyphen"`, at the cursor.

#### `insert_prime(doc, kind: str)`
*Part of `Dx_Type_Dashes`.* Inserts `"single"` (′, U+2032) or `"double"`
(″, U+2033) at the cursor. The guide stresses these **must** be used
(not straight quotes) for correct braille translation of
feet/arcminutes/minutes (single) and inches/arcseconds/seconds (double).

#### `insert_fraction(doc, numerator, denominator)`
*Part of `Dx_Type_Dashes`.* If `(numerator, denominator)` is one of the 18
compact fractions Word/DBT render as a single Unicode glyph (½, ⅓, ⅔, ¼,
¾, ⅕–⅘, ⅙, ⅚, ⅐, ⅛–⅞, ⅑, ⅒), inserts that glyph directly. Otherwise
(the guide's "Any Fraction" case, decimals allowed) inserts the DBT-coded
form `[[*fs*]]NUMERATOR[[*fl*]]DENOMINATOR[[*fe*]]` — matching the
guide's own example (`132.65/567.9`) exactly.

#### `encode_fractions_in_selection(doc, target=None)`
*Part of `Dx_Type_Dashes` ("Convert Fractions to DBT Coded Fractions").*
Finds `NUMBER/NUMBER` patterns (decimals allowed) in the target paragraph
and replaces each with the DBT-coded form. Per the guide, this is meant
for prose fractions, not dates — like the original, it's the caller's
responsibility not to run this over date-like text; the function does not
try to distinguish `12/25` (a fraction) from `12/25` (a date) itself.

#### `horizontal_list_to_vertical`
*Ports `Dx_Horz_List_To_Vertical`.* Re-exported directly from
`vistatype.lp.formatting_tools` — the guide describes byte-for-byte
identical behavior on both sides.

### `vistatype.dx.file_cleanup`

#### `fix_common_file_errors(doc, translation: str = "UEB", remove_images=None, on_confirm=None)`
*Ports `Dx_File_Fix_Sequence` step 1.* Same class of fixes as the LP
side's function (small caps → all caps, drop-cap reset, text-frame
un-boxing, confirmable image removal), with **one explicit, guide-stated
difference**: dash-run collapsing always happens, but the
space-stripping-around-dashes rule only applies when
`translation="EBAE"` — the BANA guide states this qualifier explicitly
("Removes spaces before and after Em dashes and hyphens (for EBAE
only)"), unlike the LP guide, which has no such qualifier and always
applies it.

```python
>>> fix_common_file_errors(doc, translation="UEB")   # "word --- surrounded" -> "word - surrounded" (spaces kept)
>>> fix_common_file_errors(doc, translation="EBAE")  # "word --- surrounded" -> "word-surrounded"  (spaces removed)
```

Returns the same shape as the LP side's function.

#### `delete_multiple_paragraph_marks`
Re-exported directly from `vistatype.lp.file_cleanup` — identical on both
sides (the guides don't describe any Braille-specific variation for this
one).

#### `selected_cleanup(doc, translation: str = "UEB", options=None)`
*Ports `Dx_Selected_File_CleanUp`.* Same fixes as `fix_common_file_errors`,
scoped to the current selection. `options` drawn from
`{"dashes_spaces_primes", "small_caps"}` (default: both).

#### `format_spelling_list(doc, target_paragraphs=None, spaces_between: int = 2)`
*Ports `Dx_Spelling_List`.* Spelling lists present a word first in
contracted braille form, then the same word uncontracted, so the student
sees both — DBT itself decides contraction, not this function.

> ⚠️ **Deliberately incomplete, on purpose.** This duplicates each line as
> `"<line><spaces><line>"` (the mechanical, verifiably-correct part) but
> does **not** mark the second copy as uncontracted, because the guide
> never states DBT's literal bracket code for "force uncontracted braille"
> — unlike every other DBT code used elsewhere in this project, which
> **are** quoted verbatim in a guide and used as-is. An earlier draft
> guessed at one (`[[*unc*]]`); it was removed on review, because shipping
> a wrong DBT code risks producing genuinely incorrect braille for a blind
> reader. If you know DBT's actual code for this, it's a one-line addition
> — open an issue or PR.

### `vistatype.entrypoints`

Thin wrappers — one per original ribbon command — that make every function
above callable as a real LibreOffice macro. Each one takes no meaningful
arguments (LibreOffice's Python scripting provider calls macros with no
arguments and ignores return values) and instead reaches for the
`XSCRIPTCONTEXT` global that only exists when LibreOffice itself invokes
the script (see `vistatype.utils.get_active_document` for where that
boundary is enforced elsewhere in the codebase). This file is what you
actually see listed in **Tools ▸ Macros ▸ Organize Macros ▸ Python** once
installed. Every original ribbon command name maps to exactly one function
here, listed in full in [Complete command map](#complete-command-map).

---

## Complete command map

Every one of the 45 commands `Word.officeUI` references, extracted
directly from its `onAction="..."` attributes — this is ground truth, not
a guess. Full detail (including exactly what's missing and why, for every
partial/unimplemented row) is in [`docs/COMMAND_MAP.md`](docs/COMMAND_MAP.md);
this is the condensed version.

**Legend:** ✅ implemented & tested · 🟡 partially implemented · ⬜ not
implemented (roadmap) · ⛔ deliberately not ported (About dialogs, external
video links — trivial, not workflow functions)

| Original command | Status | New home |
|---|:---:|---|
| `Sh_Apply_Title_Case_Capitalization` | ✅ | `vistatype.shared.apply_title_case` |
| `Sh_Doc_Info` | ✅ | `vistatype.shared.doc_info` |
| `Sh_Keep_Lines_Of_Para_Together` | ✅ | `vistatype.shared.toggle_keep_lines_together` |
| `Sh_Move_Paragraph_To_Next_Page` | ✅ | `vistatype.shared.toggle_move_to_next_page` |
| `Sh_Show_Char_Val` | 🟡 | `vistatype.shared.char_value_at_cursor` (Unicode code point, not legacy ANSI) |
| `MS_Change_Word_Configuration_to_New_Install` | 🟡 | `vistatype.config.set_config_new_install` |
| `MS_Change_Word_Configuration_to_Large_Print` | 🟡 | `vistatype.config.set_config_large_print` |
| `MS_Change_Word_Configuration_to_Braille` | 🟡 | `vistatype.config.set_config_braille` |
| `LP_Attach_LP_Template` | 🟡 | `vistatype.lp.attach_template.attach_lp_template` (styles/page-setup only, no size/binding wizard) |
| `Lp_AutoTag_Page_Numbers` | ✅ | `vistatype.lp.reference_pages.auto_tag_page_numbers` |
| `Lp_Manual_Tag_with_Dollar_pg` | ✅ | `vistatype.lp.reference_pages.manual_tag_page_number` |
| `Lp_Format_Page_Numbers` | ✅ | `vistatype.lp.reference_pages.format_tagged_page_numbers` |
| `Lp_Validate_Dollar_PG` | ✅ | `vistatype.lp.reference_pages.validate_tagged_page_numbers` |
| `Lp_File_Fix_Sequence` | 🟡 | `vistatype.lp.file_cleanup.file_fix_sequence` |
| `Lp_Selected_File_CleanUp` | ✅ | `vistatype.lp.file_cleanup.selected_cleanup` |
| `Lp_Horz_List_To_Vertical` | 🟡 | `vistatype.lp.formatting_tools.horizontal_list_to_vertical` ("ordered" mode auto-detect-safe; "spaced"/"tabbed" need an explicit mode) |
| `Lp_Compress_Linear_Math` | ✅ | `vistatype.lp.formatting_tools.compress_linear_math` |
| `Lp_Toggle_Space_After_Current_Para` | ✅ | `vistatype.lp.formatting_tools.toggle_space_after_para` |
| `Lp_Keep_With_Next_Para` | ✅ | `vistatype.lp.formatting_tools.toggle_keep_with_next_para` |
| `Lp_Type_Fill_In_Line` | 🟡 | `vistatype.lp.formatting_tools.type_fill_in_line` (right-margin variant uses a different, LO-native mechanism) |
| `Lp_Format_Exercise_Lv_1_and_Lv_2` | ✅ | `vistatype.lp.formatting_tools.format_exercise_levels` |
| `Lp_Picture_Color_Change_Menu` | ⬜ | Not built — see `docs/COMMAND_MAP.md` |
| `Lp_Picture_Tools_Menu_Starter` | ⬜ | Not built |
| `Lp_Table_Tools` | ⬜ | Not built |
| `Lp_About` | ⛔ | — |
| `Lp_Video_Links` | ⛔ | — |
| `Dx_AutoTag_Page_Numbers` | ✅ | `vistatype.dx.reference_pages.auto_tag_page_numbers` |
| `Dx_Manual_Tag_with_Dollar_pg` | ✅ | `vistatype.dx.reference_pages.manual_tag_page_number` |
| `Dx_Format_Tagged_Page_Numbers` | 🟡 | `vistatype.dx.reference_pages.format_tagged_page_numbers` (placeholder style — no real BANA template available) |
| `Dx_Ref_Pg_Number_Sequence_Menu` | ✅ | `vistatype.dx.reference_pages.validate_tagged_page_numbers` |
| `Dx_File_Fix_Sequence` | ✅ | `vistatype.dx.file_cleanup.fix_common_file_errors` + `delete_multiple_paragraph_marks` |
| `Dx_Selected_File_CleanUp` | ✅ | `vistatype.dx.file_cleanup.selected_cleanup` |
| `Dx_Compress_Linear_Math` | ✅ | `vistatype.dx.formatting_tools.compress_linear_math` |
| `Dx_Horz_List_To_Vertical` | ✅ | re-exported from `vistatype.lp.formatting_tools` |
| `Dx_Type_Dashes` | ✅ | `vistatype.dx.formatting_tools.insert_dash` / `insert_prime` / `insert_fraction` / `encode_fractions_in_selection` |
| `Dx_Spelling_List` | 🟡 | `vistatype.dx.file_cleanup.format_spelling_list` (mechanical duplication only — no invented DBT code) |
| `Dx_Format_Exercise_Lv_1_and_Lv_2` | ⬜ | Not built — needs BANA-specific styles/fill characters, different from the LP version |
| `Dx_Add_Color_To_Foreign_Language_Words` | ⬜ | Not built — needs BANA foreign-language character styles we don't have a source for |
| `Dx_Embed_Ref_Pg_No` | ⬜ | Not built — needs the real BANA "Embedded" style |
| `Dx_UnEmbed_Ref_Pg_No` | ⬜ | Not built — same reason |
| `Dx_About` | ⛔ | — |
| `Dx_Video_Links` | ⛔ | — |
| `DN_Add_PgNo_Tags_To_DAISY_or_NIMAS` | ✅ | `vistatype.dn_tools.add_pg_tags_to_xml` |
| `DN_Remove_Para_Formatting_From_Text_Files` | ✅ | `vistatype.dn_tools.fix_text_file_paragraphs` |

**Totals: 30 ✅ fully implemented & tested, 9 🟡 partially implemented, 6 ⬜ not yet built, 5 ⛔ deliberately skipped (trivial).**

---

## Testing

Every module has a corresponding test file that connects to a real,
disposable, headless LibreOffice instance (`lo-macros/tests/uno_harness.py`
starts one, connects over a UNO socket, and tears it down again per test
run) and asserts on actual document state after calling the function under
test — not mocks, not assumptions.

```bash
cd lo-macros/tests
./run_all_tests.sh
```

Requires `soffice` on `PATH` and `python-docx` installed
(`pip install python-docx`) for the tests that build synthetic `.docx`
fixtures. Each `test_*.py` is also independently runnable
(`python3 test_reference_pages.py`) if you're iterating on one module.

For the Pandoc side:

```bash
cd pandoc/test
./smoke_test.sh
```

This is a smoke test (prints output for a quick eyeball check across all
five filters), not an automated pass/fail suite — some checks (e.g.
confirming the real "Print Pg Num" style properties actually took effect,
not just a same-named placeholder) need LibreOffice's object model to
verify properly, which is exactly what
`lo-macros/tests/test_reference_pages.py` does instead.

---

## Technical notes for contributors

A few non-obvious things, confirmed empirically during development, that
will save you time if you extend this project:

- **`compareRegionStarts(r1, r2)`'s sign convention is inverted from
  typical `compareTo()` semantics**: it returns **+1 if `r1` is
  positioned *before* `r2`**, and −1 if `r1` is after — confirmed with an
  isolated three-point test (`start`/`mid`/`end` cursors on a known
  string) after a wrong assumption here produced backwards selection
  results in `utils.get_selected_paragraphs`.
- **Never call `uno.getComponentContext()` fresh from a script running in
  an external process** connected to soffice over a socket bridge — it
  bootstraps an unrelated local/second context instead of reusing the live
  connection, and reliably segfaults the moment you use it to create a
  cross-process service instance (confirmed: this is exactly what
  happened while building `attach_template.import_styles`). Always thread
  the `desktop` object through from wherever the original UNO connection
  was established (`uno_harness.LOSession.desktop` in tests;
  `XSCRIPTCONTEXT.getDesktop()` in a real LibreOffice-invoked macro).
- **`XStyleLoader.loadStylesFromURL` is not implemented for Writer text
  documents** in at least LibreOffice 24.2 — `queryInterface` returns
  `None`. Don't rely on it; copy style properties directly instead (see
  `lp.attach_template._copy_style_properties`), which is slower but
  actually works and doesn't depend on an interface that may or may not
  be present depending on the user's LO version.
- **Don't batch-delete paragraphs by collecting references up front.**
  Re-enumerate fresh before each deletion (see
  `lp.file_cleanup.delete_multiple_paragraph_marks`'s docstring) — a
  cached paragraph reference from before an earlier structural edit can
  throw `UnknownPropertyException` on `.Start`/`.End` access.
- **Pandoc has no `Tab` Inline type.** A literal tab character collapses
  to a plain `Space` during Markdown parsing (`printf 'a\tb'` parses to
  `Str "a", Space, Str "b"`) — don't write a filter that checks for
  `inline.t == "Tab"`, it will silently never match anything.
- **Pandoc's own DOCX reader silently drops fully-empty paragraphs**
  before they ever reach the AST — confirmed by round-tripping a
  `python-docx`-authored document with real empty paragraphs through
  `pandoc -t native` and finding zero `Para` nodes for them. This is
  *stronger* than "collapse multiple to one" (it's "collapse to zero"),
  so `lp-delete-multiple-paragraph-marks.lua` is close to a no-op for
  DOCX input specifically — see that filter's own header comment.
- **Pandoc's markdown reader has `+smart` on by default**
  (`pandoc --list-extensions=markdown` confirms it), meaning `"---"` is
  already converted to an em dash and `"--"` to an en dash *before* any
  filter sees the text. Test dash-collapsing filters with `-f
  markdown-smart` (smart *disabled*) if you want to see raw repeated-hyphen
  input, or you'll be testing against text that's already been
  transformed.
- **A single Pandoc `Str` token can span a marker and following word
  fused together** — e.g. Markdown's own list-continuation-ambiguity
  avoidance inserts a non-breaking space, producing a token like
  `"c.\u00A0the"` rather than separate `"c."`, `Space`, `"the"` tokens.
  Any filter matching markers/patterns at the `Str`-token level needs to
  handle a fused remainder, not assume clean token boundaries.

---

## Known limitations & unfinished work

See [`docs/COMMAND_MAP.md`](docs/COMMAND_MAP.md) for the row-by-row
detail behind every 🟡/⬜ in the table above. In rough order of size:

1. **`LP_Attach_LP_Template`'s paper/screen-size wizard** (target paper
   size, mirroring/binding, print-vs-screen output profiles) — a
   substantial, mostly-new dialog-driven feature, not a style-copy
   operation. Not started.
2. **`Lp_Table_Tools`** and **`Lp_Picture_Tools_Menu_Starter`/
   `Lp_Picture_Color_Change_Menu`** — guide sections for these weren't
   fully mined yet (the 162-page VistaType LP guide's Attach/File-Cleanup/
   Reference-Page/Other-Formatting/DAISY/Help sections were prioritized
   first, since they cover the core transcription workflow end to end).
   Table tools in particular (rotation, title-splitting, list conversion)
   look like a meaningfully-sized feature on their own.
3. **Several `Dx_` (Braille) commands need a real BANA Braille Template
   file we don't have**: `Dx_Format_Exercise_Lv_1_and_Lv_2` (BANA-specific
   `Exercise1`/`Exercise2` styles and EBAE/UEB-specific fill characters),
   `Dx_Add_Color_To_Foreign_Language_Words` (BANA foreign-language
   character styles), `Dx_Embed_Ref_Pg_No`/`Dx_UnEmbed_Ref_Pg_No` (the
   real "Embedded (Mid-Paragraph) Reference Page Number" style). Only
   `LargePrintTemplate.dotx` was available as a source template — no BANA
   `.dotx`. Buildable once such a file (or the exact style
   names/properties to hard-code as a fallback, the way `Print Pg Num`
   and `List`/`List 2` are handled now) is available.
4. **`Dx_Spelling_List`'s "uncontracted" DBT bracket code** — deliberately
   left undone rather than guessed at; see that function's docs above and
   `docs/FINDINGS.md`.

If your workflow depends on exact byte-for-byte parity with the original
add-in's output for any of the ✅/🟡 items, test against your own real
documents before relying on this for production transcription work — this
is a good-faith reconstruction from documentation, verified function-by-
function, not a certified drop-in replacement.

---

## License

The original VistaType LP / BANA Macros add-in's license (quoted from its
user guide) grants broad rights:

> Permission is hereby granted, free of charge, to any individual or
> organization obtaining a copy of this software and its associated
> documentation files, to use the Software without restriction for
> personal, educational, or commercial use. ... The Software is provided
> "as is," without any warranties or conditions of any kind.

This repository's code is a new, independent implementation (see
[Why this is a reconstruction](#why-this-is-a-reconstruction-not-a-translation)),
built from the original's documented *behavior*, not its source. No
license file is checked in yet — add one (e.g. MIT, matching the spirit of
the original's permissive grant) before treating this as a formally
licensed open-source project.

## Acknowledgments

The original VistaType LP and BANA Braille Template macro suite was
written by **Jerry Whittaker** (jerry@thewhittakers.org). This project
would not exist without his many years of work building and documenting
tools for large-print and braille transcription accessibility — the two
user guides referenced throughout this README and `docs/` are his.

## Contributing

Issues and PRs welcome, especially:
- A working copy of a real BANA Braille Template file, or the exact style
  names/properties from one, to unblock the `Dx_` items in
  [Known limitations](#known-limitations--unfinished-work).
- DBT's actual bracket code for "render this span uncontracted" (see
  `Dx_Spelling_List` above).
- Anything from [Known limitations](#known-limitations--unfinished-work).

When contributing a new function, please follow the existing convention:
explicit `doc` argument, optional `target=` defaulting to the current
paragraph/selection, a docstring that states which original command it
ports and cites the guide page number where possible, and a corresponding
`test_*.py` that actually exercises it against a live headless LibreOffice
instance — see any existing module for the pattern.
