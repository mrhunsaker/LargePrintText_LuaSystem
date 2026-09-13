"""
vistatype.lp.attach_template
=============================
Port of "Attach LP Template" (LP_Attach_LP_Template ribbon command;
VistaType LP User Guide p.35-43, "Attach Group").

The original macro attaches LargePrintTemplate.dotx to the active Word
document (and, per the guide, also handles a whole paper/screen-size and
media-selection wizard we are NOT reproducing here -- see
docs/COMMAND_MAP.md for that gap).

WHY THIS DOESN'T JUST "ATTACH" A TEMPLATE FILE
================================================
LibreOffice documents don't carry a live pointer to an external "attached
template" the way Word's .dotx attachment works. The closest built-in
mechanism, XStyleLoader.loadStylesFromURL, is NOT implemented for Writer
text documents in this LibreOffice build (verified empirically -- queryInterface
returns None). Rather than depend on an interface that may or may not exist
depending on the user's LO version, this module copies styles the robust,
version-independent way: open the reference document, walk its
ParagraphStyles / CharacterStyles / PageStyles, and copy every settable
property onto a same-named style in the target document (creating it if
needed). This was verified against a real headless LibreOffice instance to
faithfully reproduce font size, letter-spacing, shading, borders, and page
geometry -- see lo-macros/tests/test_attach_template.py.

By default this pulls from the bundled `reference-doc/LargePrintTemplate-
reference.odt` (the same file the Pandoc pipeline uses as --reference-doc,
so both paths draw from one verified source of truth). Pass a different
`template_path` to use your own copy of LargePrintTemplate.dotx, or an
.ott you've built from it, directly.
"""
from __future__ import annotations

import os

from com.sun.star.beans import PropertyValue

_HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE_PATH = os.path.normpath(
    os.path.join(_HERE, "..", "..", "..", "reference-doc", "LargePrintTemplate-reference.odt")
)

_STYLE_FAMILY_SERVICE = {
    "ParagraphStyles": "com.sun.star.style.ParagraphStyle",
    "CharacterStyles": "com.sun.star.style.CharacterStyle",
    "PageStyles": "com.sun.star.style.PageStyle",
}


def _make_prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def _copy_style_properties(src_style, dst_style):
    """Copy every gettable+settable property from src_style to dst_style.
    Read-only / computed properties (there are always a couple dozen --
    things like CharDiffHeight, BottomBorderComplexColor) fail silently;
    that's expected and harmless, verified against a real style catalog.
    """
    copied, failed = 0, 0
    for prop in src_style.getPropertySetInfo().getProperties():
        try:
            value = src_style.getPropertyValue(prop.Name)
            dst_style.setPropertyValue(prop.Name, value)
            copied += 1
        except Exception:
            failed += 1
    return copied, failed


def import_styles(doc, template_path=None, family_names=None, desktop=None):
    """Copy all styles (or just `family_names`, e.g. ["ParagraphStyles"])
    from the template file at `template_path` into `doc`. `doc` and the
    template are opened/compared live via UNO; the template is opened
    hidden and closed again before returning.

    `desktop` must be a real com.sun.star.frame.Desktop from the SAME UNO
    connection `doc` came from. We deliberately don't create one for you
    via a fresh `uno.getComponentContext()` -- verified experimentally that
    doing so crashes (segfaults) when this runs in an external client
    process talking to soffice over a socket bridge, because it bootstraps
    an unrelated local/second context instead of reusing the live
    connection. Callers get a desktop two ways:
      - from LibreOffice's own scripting provider: XSCRIPTCONTEXT.getDesktop()
      - from an external UNO connection: whatever you used to create `doc`
        in the first place (see lo-macros/tests/uno_harness.py: LOSession
        keeps `self.desktop` around for exactly this reason).

    Returns a dict: {family_name: {style_name: (copied_count, failed_count)}}
    for anyone who wants to sanity-check what happened.
    """
    template_path = template_path or DEFAULT_TEMPLATE_PATH
    if not os.path.exists(template_path):
        raise FileNotFoundError(
            f"Template not found at {template_path!r}. Pass template_path= "
            "explicitly, or run reference-doc/build_reference_odt.sh first."
        )
    if desktop is None:
        raise ValueError(
            "import_styles() requires a `desktop` from the same UNO "
            "connection as `doc` -- see the docstring above."
        )

    url = "file://" + os.path.abspath(template_path)
    template_doc = desktop.loadComponentFromURL(
        url, "_blank", 0, (_make_prop("Hidden", True), _make_prop("ReadOnly", True))
    )

    results = {}
    try:
        families = family_names or list(_STYLE_FAMILY_SERVICE.keys())
        for family_name in families:
            service = _STYLE_FAMILY_SERVICE[family_name]
            src_family = template_doc.StyleFamilies.getByName(family_name)
            dst_family = doc.StyleFamilies.getByName(family_name)
            family_results = {}
            for name in src_family.ElementNames:
                src_style = src_family.getByName(name)
                if not dst_family.hasByName(name):
                    try:
                        new_style = doc.createInstance(service)
                        dst_family.insertByName(name, new_style)
                    except Exception:
                        continue
                dst_style = dst_family.getByName(name)
                family_results[name] = _copy_style_properties(src_style, dst_style)
            results[family_name] = family_results
    finally:
        template_doc.close(False)

    return results


def attach_lp_template(doc, template_path=None, desktop=None):
    """LP_Attach_LP_Template (style/page-setup portion only -- see module
    docstring and docs/COMMAND_MAP.md for the paper/screen-size wizard and
    media-selection parts of the original that this does not cover).

    After this call: the document's paragraph/character styles include
    Standard (Normal), Print Pg Num, and all the Para*/Words*/Text*/Box*
    accessibility contrast styles, with the same properties as the real
    LargePrintTemplate.dotx; and the default page style has the template's
    page size and margins. Existing body text keeps whatever direct
    formatting it already had -- like Word's own template attachment, this
    changes the *style catalog*, not text that already overrides it.

    See import_styles() above for why `desktop` must be passed explicitly.
    """
    results = import_styles(doc, template_path, desktop=desktop)

    # Make sure new content actually uses the imported page geometry: set
    # the whole document's page style explicitly (a fresh blank document's
    # paragraphs may not yet reference "Standard" for PageDescName).
    for para in doc.Text.createEnumeration():
        if para.supportsService("com.sun.star.text.Paragraph"):
            para.PageDescName = "Standard"
            break  # only the first paragraph needs PageDescName set

    return results
