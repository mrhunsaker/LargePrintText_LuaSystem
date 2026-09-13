import sys, os
sys.path.insert(0, ".")
sys.path.insert(0, "..")
import uno
from uno_harness import LOSession
from vistatype import utils

with LOSession() as lo:
    doc = lo.open_blank_writer()
    text = doc.Text
    cur = text.createTextCursor()
    text.insertString(cur, "Para A", False)
    text.insertControlCharacter(cur, uno.getConstantByName("com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK"), False)
    text.insertString(cur, "Para B", False)
    text.insertControlCharacter(cur, uno.getConstantByName("com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK"), False)
    text.insertString(cur, "Para C", False)

    paras = list(utils.iter_paragraphs(doc))
    print("paragraph count:", len(paras), [p.getString() for p in paras])

    # place view cursor in Para B, test get_current_paragraph
    vc = doc.CurrentController.ViewCursor
    vc.gotoRange(paras[1].Start, False)
    cur_para = utils.get_current_paragraph(doc)
    print("current paragraph (expect Para B):", cur_para.getString() if cur_para else None)

    # select from Para A to Para B, test get_selected_paragraphs
    vc.gotoRange(paras[0].Start, False)
    vc.gotoRange(paras[1].End, True)
    print("selection text:", repr(vc.getString()))
    sel_paras = utils.get_selected_paragraphs(doc)
    print("selected paragraphs (expect A, B):", [p.getString() for p in sel_paras])

    doc.close(False)
