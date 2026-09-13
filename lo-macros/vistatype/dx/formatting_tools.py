"""
vistatype.dx.formatting_tools
================================
Port of the Braille/BANA-side "Other Formatting Tools Group"
(BANA Macros for 2.0 User Guide p.20-25):

  Dx_Compress_Linear_Math  -> compress_linear_math
  Dx_Type_Dashes           -> insert_dash, insert_prime, insert_fraction,
                              encode_fractions_in_selection
  Dx_Horz_List_To_Vertical -> re-exported from vistatype.lp.formatting_tools
                              (the guide describes byte-for-byte the same
                              behavior on both sides -- see p.22-23 there
                              vs p.50-51 on the LP side)
"""
from __future__ import annotations

import re

from .. import utils
from ..lp.formatting_tools import horizontal_list_to_vertical  # noqa: F401 re-export

_COMPARISON_SYMBOLS = "=\u2248<>\u2260\u2264\u2265"


def compress_linear_math(doc, target=None):
    """Dx_Compress_Linear_Math.

    Removes ALL spaces from the target paragraph's linear math EXCEPT
    those immediately preceding and following a comparison symbol (=,
    \u2248, <>, <, >, \u2260, \u2265, \u2264), which are left as ordinary
    spaces -- this is the Braille-side rule verbatim from the guide
    ("removes all spaces from linear math except for those preceding and
    following signs of comparison", p.25).

    This is a DIFFERENT rule from the Large Print side's macro of the
    same name (which removes ALL spaces, including around comparison
    symbols, then adds thin/hair spaces back around them) -- see
    vistatype.lp.formatting_tools.compress_linear_math's docstring for
    the side-by-side. Both are implemented faithfully to their own
    guide's wording; the two guides simply specify different behavior for
    a macro with the same name.
    """
    para = target if target is not None else utils.get_current_paragraph(doc)
    if para is None:
        return None
    text = para.getString()

    # Walk the string; keep a space only if it's adjacent to a comparison
    # symbol, drop every other space.
    out = []
    n = len(text)
    for i, ch in enumerate(text):
        if ch == " ":
            prev_ch = text[i - 1] if i > 0 else ""
            next_ch = text[i + 1] if i + 1 < n else ""
            if prev_ch in _COMPARISON_SYMBOLS or next_ch in _COMPARISON_SYMBOLS:
                out.append(ch)
            # else: drop it
        else:
            out.append(ch)
    para.setString("".join(out))
    return para


# ---------------------------------------------------------------------------
# Dashes / Primes / Fractions (BANA guide p.24-25)
# ---------------------------------------------------------------------------

DASHES = {
    "long_dash": "\u2015",   # horizontal bar
    "em_dash": "\u2014",
    "en_dash": "\u2013",
    "minus_sign": "\u2212",
    "hyphen": "\u2010",
    "nonbreaking_hyphen": "\u2011",
}

PRIMES = {
    "single": "\u2032",
    "double": "\u2033",
}

# The eighteen compact fractions the guide says Word can produce natively
# (and that DBT translates correctly without extra markup) -- Unicode
# vulgar fraction characters, matching the menu's fraction buttons.
COMPACT_FRACTIONS = {
    (1, 2): "\u00BD", (1, 3): "\u2153", (2, 3): "\u2154",
    (1, 4): "\u00BC", (3, 4): "\u00BE",
    (1, 5): "\u2155", (2, 5): "\u2156", (3, 5): "\u2157", (4, 5): "\u2158",
    (1, 6): "\u2159", (5, 6): "\u215A",
    (1, 7): "\u2150",
    (1, 8): "\u215B", (3, 8): "\u215C", (5, 8): "\u215D", (7, 8): "\u215E",
    (1, 9): "\u2151",
    (1, 10): "\u2152",
}

FRACTION_START = "[[*fs*]]"
FRACTION_LINE = "[[*fl*]]"
FRACTION_END = "[[*fe*]]"


def insert_dash(doc, kind: str):
    """Dx_Type_Dashes menu: "Dashes" section. `kind` is one of
    DASHES.keys() ("long_dash", "em_dash", "en_dash", "minus_sign",
    "hyphen", "nonbreaking_hyphen").
    """
    if kind not in DASHES:
        raise ValueError(f"unknown dash kind: {kind!r}; choices: {sorted(DASHES)}")
    _insert_at_cursor(doc, DASHES[kind])


def insert_prime(doc, kind: str):
    """Dx_Type_Dashes menu: "Primes" section. `kind` is "single" or
    "double". The guide stresses these MUST be used (not straight quotes)
    for correct braille translation of feet/arcminutes/minutes (single)
    and inches/arcseconds/seconds (double).
    """
    if kind not in PRIMES:
        raise ValueError(f"unknown prime kind: {kind!r}; choices: {sorted(PRIMES)}")
    _insert_at_cursor(doc, PRIMES[kind])


def insert_fraction(doc, numerator, denominator):
    """Dx_Type_Dashes menu: "Fractions" / "Any Fraction" sections.

    If (numerator, denominator) is one of the eighteen compact fractions
    Word can render natively (see COMPACT_FRACTIONS), inserts that single
    Unicode glyph. Otherwise (the "Any Fraction" case, which also allows
    decimals per the guide), inserts the DBT-coded form:
    [[*fs*]]NUMERATOR[[*fl*]]DENOMINATOR[[*fe*]] -- matching the guide's
    own example (132.65/567.9).
    """
    key = (numerator, denominator)
    if key in COMPACT_FRACTIONS:
        _insert_at_cursor(doc, COMPACT_FRACTIONS[key])
        return
    _insert_at_cursor(doc, f"{FRACTION_START}{numerator}{FRACTION_LINE}{denominator}{FRACTION_END}")


_TEXT_FRACTION_RE = re.compile(r"(?<!\d)(\d+(?:\.\d+)?)/(\d+(?:\.\d+)?)(?!\d)")


def encode_fractions_in_selection(doc, target=None):
    """Dx_Type_Dashes menu: "Convert Fractions to DBT Coded Fractions".

    Finds NUMBER/NUMBER patterns (decimals allowed) in the target
    paragraph and replaces them with the DBT-coded form. Per the guide,
    this is meant for prose fractions, not dates -- callers are
    responsible for not running this over date-like text (same caveat the
    original macro gives: "Do not include text that contains dates in
    this format").
    """
    para = target if target is not None else utils.get_current_paragraph(doc)
    if para is None:
        return None
    text = para.getString()

    def repl(m):
        return f"{FRACTION_START}{m.group(1)}{FRACTION_LINE}{m.group(2)}{FRACTION_END}"

    new_text = _TEXT_FRACTION_RE.sub(repl, text)
    if new_text != text:
        para.setString(new_text)
    return para


def _insert_at_cursor(doc, s: str):
    controller = doc.CurrentController
    if controller is None:
        raise RuntimeError("insert_* functions need a document with a view (ViewCursor)")
    vc = controller.ViewCursor
    vc.Text.insertString(vc, s, False)
