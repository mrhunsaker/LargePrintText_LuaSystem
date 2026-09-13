import sys
sys.path.insert(0, ".")
sys.path.insert(0, "..")
import uno
from uno_harness import LOSession
from vistatype import utils
from vistatype.lp import formatting_tools as ft, attach_template

PARA_BREAK = uno.getConstantByName("com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK")
def add_para(text_obj, cursor, s):
    text_obj.insertString(cursor, s, False)
    text_obj.insertControlCharacter(cursor, PARA_BREAK, False)

with LOSession() as lo:
    doc = lo.open_blank_writer()
    attach_template.attach_lp_template(doc, desktop=lo.desktop)

    text = doc.Text
    cur = text.createTextCursor()
    text.insertString(cur, "Name:", False)
    vc = doc.CurrentController.ViewCursor
    vc.gotoEnd(False)
    ft.type_fill_in_line(doc, to_right_margin=True)
    print("paragraph after fill-to-margin:", repr(doc.Text.getString()))
    para = next(iter(utils.iter_paragraphs(doc)))
    print("tab stops:", [(s.Position, s.FillChar, s.Alignment) for s in para.ParaTabStops])
    assert doc.Text.getString() == "Name:\t"
    assert para.ParaTabStops[-1].FillChar == "_"

    # Exercise levels test
    doc2 = lo.open_blank_writer()
    attach_template.attach_lp_template(doc2, desktop=lo.desktop)
    text2 = doc2.Text
    cur2 = text2.createTextCursor()
    add_para(text2, cur2, "1. Who killed Cock Robin?")
    add_para(text2, cur2, "a. the Sparrow\t")
    add_para(text2, cur2, "b. the Fly _")
    text2.insertString(cur2, "c. the Fish", False)

    paras2 = list(utils.iter_paragraphs(doc2))
    vc2 = doc2.CurrentController.ViewCursor
    vc2.gotoRange(paras2[0].Start, False)
    vc2.gotoRange(paras2[-1].End, True)

    result = ft.format_exercise_levels(doc2)
    print("exercise format result:", result)
    for p in utils.iter_paragraphs(doc2):
        print("  ", repr(p.getString()), p.ParaStyleName)

    assert result == {"questions": 1, "answers": 3}
    styles = [p.ParaStyleName for p in utils.iter_paragraphs(doc2)]
    assert styles == ["List", "List 2", "List 2", "List 2"]
    texts = [p.getString() for p in utils.iter_paragraphs(doc2)]
    # NOTE: the automatic tab/underscore substitution inside exercise
    # formatting does not add a leading space (that's only documented for
    # the standalone Lp_Type_Fill_In_Line dialog tool) -- it just replaces
    # the tab/underscore character(s) themselves in place.
    assert texts[1] == "a. the Sparrow________", texts[1]
    assert texts[2] == "b. the Fly ________", texts[2]

    doc.close(False)
    doc2.close(False)
    print("FORMATTING TOOLS TEST 3 PASSED")
