import sys
sys.path.insert(0, ".")
sys.path.insert(0, "..")
import uno
from uno_harness import LOSession
from vistatype import utils
from vistatype.lp import formatting_tools as ft, attach_template

with LOSession() as lo:
    doc = lo.open_blank_writer()
    text = doc.Text
    text.setString("y = 2 x + 1")

    para = next(iter(utils.iter_paragraphs(doc)))
    ft.compress_linear_math(doc, target=para)
    result = para.getString()
    print("compressed math:", repr(result))
    assert result == "y\u2009=\u20092x+1", repr(result)  # thin space around '='

    # toggle space after para -- give the paragraph a real nonzero bottom
    # margin first (a fresh blank doc's Standard style defaults to 0,
    # which would make this test ambiguous -- see the toggle's docstring
    # for the inherent limitation of a stateless on/off toggle).
    families = doc.StyleFamilies.getByName("ParagraphStyles")
    para.ParaStyleName = "Standard"
    para.ParaBottomMargin = 500
    original_after = para.ParaBottomMargin
    print("original ParaBottomMargin:", original_after)
    off = ft.toggle_space_after_para(doc, target=para)
    print("toggle 1 -> space removed?", off, " current:", para.ParaBottomMargin)
    assert off is True and para.ParaBottomMargin == 0
    on = ft.toggle_space_after_para(doc, target=para)
    print("toggle 2 -> space restored?", not on, " current:", para.ParaBottomMargin)
    # NOTE: restores to the *style's* bottom margin (0 for a vanilla blank
    # doc's Standard style, so it falls back to the documented ~0.25in
    # default), not necessarily this paragraph's own original direct
    # value -- toggle_space_after_para is documented as stateless.
    assert para.ParaBottomMargin > 0

    # keep with next para
    k1 = ft.toggle_keep_with_next_para(doc, target=para)
    print("keep with next 1:", k1, para.ParaKeepTogether)
    assert k1 is True
    k2 = ft.toggle_keep_with_next_para(doc, target=para)
    print("keep with next 2:", k2, para.ParaKeepTogether)
    assert k2 is False

    # fill-in line (fixed length)
    attach_template.attach_lp_template(doc, desktop=lo.desktop)
    vc = doc.CurrentController.ViewCursor
    vc.gotoEnd(False)
    ft.type_fill_in_line(doc, length=8)
    final_text = doc.Text.getString()
    print("after fill-in line:", repr(final_text))
    assert final_text.endswith(" " + "_"*8)

    doc.close(False)
    print("FORMATTING TOOLS TEST 2 PASSED")
