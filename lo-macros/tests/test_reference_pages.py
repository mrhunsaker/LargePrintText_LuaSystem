import sys
sys.path.insert(0, ".")
sys.path.insert(0, "..")
import uno
from uno_harness import LOSession
from vistatype import utils
from vistatype.lp import reference_pages as rp

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
    add_para(text, cur, "Chapter One")
    add_para(text, cur, "Some prose here describing things.")
    add_para(text, cur, "12")
    add_para(text, cur, "More prose after.")
    add_para(text, cur, "A13")
    add_para(text, cur, "Even more text.")
    add_para(text, cur, "16-18")
    add_para(text, cur, "XIV")
    text.insertString(cur, "Final paragraph of the test.", False)

    n = rp.auto_tag_page_numbers(doc)
    print("tagged:", n)
    tagged_texts = [p.getString() for p in utils.iter_paragraphs(doc) if p.getString().startswith("$pg")]
    print("tagged paragraphs:", tagged_texts)
    assert tagged_texts == ["$pg12", "$pgA13", "$pg16-18", "$pgXIV"]

    entries = rp.validate_tagged_page_numbers(doc)
    for e in entries:
        print(e)

    m = rp.format_tagged_page_numbers(doc)
    print("formatted:", m)

    para_styles = doc.StyleFamilies.getByName("ParagraphStyles")
    styled = [(p.getString(), p.ParaStyleName) for p in utils.iter_paragraphs(doc) if p.ParaStyleName == "Print Pg Num"]
    print("styled paragraphs:", styled)
    assert styled == [("12", "Print Pg Num"), ("A13", "Print Pg Num"), ("16-18", "Print Pg Num"), ("XIV", "Print Pg Num")]
    back = para_styles.getByName("Print Pg Num").ParaBackColor
    print("Print Pg Num back color:", hex(back))
    assert back == 0xF4A3D4

    # manual tag test: untag one, then manually tag it back
    para0 = next(iter(utils.iter_paragraphs(doc)))
    rp.manual_tag_page_number(doc, target=para0)
    print("manually tagged first paragraph:", para0.getString())
    assert para0.getString().startswith("$pg")

    doc.close(False)
    print("REFERENCE PAGES TEST PASSED")
