import sys
sys.path.insert(0, ".")
sys.path.insert(0, "..")
import uno
from uno_harness import LOSession
from vistatype import utils
from vistatype.dx import reference_pages as dxrp

PARA_BREAK = uno.getConstantByName("com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK")
def add_para(text_obj, cursor, s):
    text_obj.insertString(cursor, s, False)
    text_obj.insertControlCharacter(cursor, PARA_BREAK, False)

with LOSession() as lo:
    doc = lo.open_blank_writer()
    text = doc.Text
    cur = text.createTextCursor()
    add_para(text, cur, "13-15")
    add_para(text, cur, "c")
    add_para(text, cur, "xiv")
    text.insertString(cur, "12", False)

    n = dxrp.auto_tag_page_numbers(doc, translation="UEB")
    tagged = [p.getString() for p in utils.iter_paragraphs(doc)]
    print("UEB tagged:", tagged)
    assert tagged[0] == "$pg13-15[[*lec*]]15"
    assert tagged[1] == "$pgc[[*ii*]]"   # 'c' is UEB-ambiguous -> gets letter sign
    assert tagged[2] == "$pgxiv"          # 'xiv' not in ambiguous set -> no letter sign
    assert tagged[3] == "$pg12"

    m = dxrp.format_tagged_page_numbers(doc)
    styled = [(p.getString(), p.ParaStyleName) for p in utils.iter_paragraphs(doc)]
    print("styled:", styled)
    assert styled[0] == ("13-15[[*lec*]]15", "BANA Braille Template Reference Page Number")
    assert styled[1] == ("c[[*ii*]]", "BANA Braille Template Reference Page Number")

    doc.close(False)

    # Separately test EBAE mode: all lowercase romans get the letter sign
    doc2 = lo.open_blank_writer()
    text2 = doc2.Text
    cur2 = text2.createTextCursor()
    add_para(text2, cur2, "xiv")
    text2.insertString(cur2, "c", False)
    dxrp.auto_tag_page_numbers(doc2, translation="EBAE")
    tagged2 = [p.getString() for p in utils.iter_paragraphs(doc2)]
    print("EBAE tagged:", tagged2)
    assert tagged2[0] == "$pgxiv[[*ii*]]"
    assert tagged2[1] == "$pgc[[*ii*]]"
    doc2.close(False)

    print("DX REFERENCE PAGES TEST PASSED")
