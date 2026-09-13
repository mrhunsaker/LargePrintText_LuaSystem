import sys
sys.path.insert(0, ".")
sys.path.insert(0, "..")
import uno
from uno_harness import LOSession
from vistatype import utils
from vistatype.lp import file_cleanup

PARA_BREAK = None
def add_para(text_obj, cursor, s):
    global PARA_BREAK
    if PARA_BREAK is None:
        PARA_BREAK = uno.getConstantByName("com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK")
    text_obj.insertString(cursor, s, False)
    text_obj.insertControlCharacter(cursor, PARA_BREAK, False)

with LOSession() as lo:
    doc = lo.open_blank_writer()
    text = doc.Text
    cur = text.createTextCursor()
    add_para(text, cur, "First")
    add_para(text, cur, "")
    add_para(text, cur, "")
    add_para(text, cur, "")
    text.insertString(cur, "Last", False)

    print("before:", [p.getString() for p in utils.iter_paragraphs(doc)])
    n = file_cleanup.delete_multiple_paragraph_marks(doc)
    print("deleted:", n)
    print("after:", [p.getString() for p in utils.iter_paragraphs(doc)])
    doc.close(False)
