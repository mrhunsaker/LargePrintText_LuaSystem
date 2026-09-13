import sys
sys.path.insert(0, ".")
sys.path.insert(0, "..")
import uno
from uno_harness import LOSession
from vistatype import utils, shared

with LOSession() as lo:
    doc = lo.open_blank_writer()
    text = doc.Text
    cur = text.createTextCursor()
    PARA_BREAK = uno.getConstantByName("com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK")
    text.insertString(cur, "the quick brown fox and the lazy dog: a tale of two cities", False)
    text.insertControlCharacter(cur, PARA_BREAK, False)
    text.insertString(cur, "second paragraph, unrelated", False)

    vc = doc.CurrentController.ViewCursor
    vc.gotoStart(False)

    para = shared.apply_title_case(doc)
    print("title case:", repr(para.getString()))
    assert para.getString() == "The Quick Brown Fox and the Lazy Dog: a Tale of Two Cities", para.getString()

    kept = shared.toggle_keep_lines_together(doc)
    print("keep lines together ->", kept, " ParaSplit=", para.ParaSplit)
    assert kept is True and para.ParaSplit is False

    moved = shared.toggle_move_to_next_page(doc)
    print("move to next page ->", moved, " BreakType=", para.BreakType)
    assert moved is True

    moved_back = shared.toggle_move_to_next_page(doc)
    print("move to next page (toggle back) ->", moved_back, " BreakType=", para.BreakType)
    assert moved_back is False

    vc.gotoStart(False)
    charval = shared.char_value_at_cursor(doc)
    print("char at cursor:", charval)
    assert charval["char"] == "T"  # after title-casing, first char is 'T'

    info = shared.doc_info(doc)
    print("doc_info:", info)
    # NOTE: a blank LO doc defaults to this locale's page size (A4 here),
    # not Letter -- that's expected since attach_lp_template() hasn't been
    # run in this test. Just check the shape of the result.
    assert info["width_in"] > 0 and info["height_in"] > 0
    assert info["has_large_print_styles"] is False

    doc.close(False)
    print("ALL SHARED TESTS PASSED")
