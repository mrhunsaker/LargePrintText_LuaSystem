import sys
sys.path.insert(0, ".")
sys.path.insert(0, "..")
import uno
from uno_harness import LOSession
from vistatype import utils
from vistatype.lp import formatting_tools as ft

PARA_BREAK = uno.getConstantByName("com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK")
def add_para(text_obj, cursor, s):
    text_obj.insertString(cursor, s, False)
    text_obj.insertControlCharacter(cursor, PARA_BREAK, False)

with LOSession() as lo:
    doc = lo.open_blank_writer()
    text = doc.Text
    cur = text.createTextCursor()
    add_para(text, cur, "1. Who killed Cock Robin?")
    add_para(text, cur, "a. the Sparrow b. the Fly c. the Fish d. the Beetle")
    text.insertString(cur, "Final paragraph.", False)

    vc = doc.CurrentController.ViewCursor
    paras = list(utils.iter_paragraphs(doc))
    vc.gotoRange(paras[1].Start, False)

    result = ft.horizontal_list_to_vertical(doc, mode="auto")
    print("horiz->vert result:", [p.getString() if hasattr(p,'getString') else p for p in (result or [])])
    all_paras = [p.getString() for p in utils.iter_paragraphs(doc)]
    print("all paragraphs now:", all_paras)
    assert all_paras == [
        "1. Who killed Cock Robin?",
        "a. the Sparrow",
        "b. the Fly",
        "c. the Fish",
        "d. the Beetle",
        "Final paragraph.",
    ], all_paras

    doc.close(False)
    print("HORIZ TO VERT TEST PASSED")
