import sys
sys.path.insert(0, ".")
sys.path.insert(0, "..")
import uno
from uno_harness import LOSession
from vistatype import shared
from vistatype.lp import attach_template

with LOSession() as lo:
    doc = lo.open_blank_writer()
    results = attach_template.attach_lp_template(doc, desktop=lo.desktop)
    for family, styles in results.items():
        print(family, "->", len(styles), "styles processed")

    para_styles = doc.StyleFamilies.getByName("ParagraphStyles")
    assert para_styles.hasByName("Print Pg Num")
    assert para_styles.hasByName("Para Black")
    char_styles = doc.StyleFamilies.getByName("CharacterStyles")
    assert char_styles.hasByName("Words Black Inverted")

    standard = para_styles.getByName("Standard")
    print("Standard font size:", standard.CharHeight, "letter spacing:", standard.CharKerning)
    assert standard.CharHeight == 18.0

    info = shared.doc_info(doc)
    print("doc_info after attach:", info)
    assert info["width_in"] == 8.5
    assert info["height_in"] == 11.0
    assert info["has_large_print_styles"] is True

    doc.close(False)
    print("ATTACH TEMPLATE TEST PASSED")
