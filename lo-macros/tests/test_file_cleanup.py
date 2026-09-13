import sys
sys.path.insert(0, ".")
sys.path.insert(0, "..")
import uno
from uno_harness import LOSession
from vistatype import utils
from vistatype.lp import file_cleanup
from com.sun.star.style.CaseMap import SMALLCAPS

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
    add_para(text, cur, "This has a triple---dash and a double--dash issue.")
    add_para(text, cur, "")
    add_para(text, cur, "")
    add_para(text, cur, "")
    add_para(text, cur, "Score was 98.6' and 12\" today.")
    text.insertString(cur, "Final paragraph.", False)

    paras_before = [p.getString() for p in utils.iter_paragraphs(doc)]
    print("BEFORE:", paras_before)

    # add a small-caps portion to paragraph 1 for testing
    paras = list(utils.iter_paragraphs(doc))
    p0_cursor = text.createTextCursorByRange(paras[0].Start)
    p0_cursor.gotoEndOfParagraph(True)
    p0_cursor.CharCaseMap = SMALLCAPS

    result = file_cleanup.file_fix_sequence(doc, remove_images=False)
    print("result:", result)

    paras_after = [p.getString() for p in utils.iter_paragraphs(doc)]
    print("AFTER:", paras_after)

    assert paras_after[0] == "THIS HAS A TRIPLE-DASH AND A DOUBLE-DASH ISSUE."
    assert paras_after[1] == "", "expected exactly one empty paragraph to remain (collapsed from 3), not zero"
    assert "98.6\u2032" in paras_after[2] and "12\u2033" in paras_after[2]
    assert paras_after[3] == "Final paragraph."
    assert len(paras_after) == 4, f"expected 4 paragraphs after collapsing 3 empties down to 1, got {len(paras_after)}"

    doc.close(False)
    print("FILE CLEANUP TEST PASSED")
