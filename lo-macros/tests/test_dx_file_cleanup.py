import sys
sys.path.insert(0, ".")
sys.path.insert(0, "..")
import uno
from uno_harness import LOSession
from vistatype import utils
from vistatype.dx import file_cleanup as dxfc

with LOSession() as lo:
    doc = lo.open_blank_writer()
    text = doc.Text
    text.setString("word --- surrounded by dashes")
    para = next(iter(utils.iter_paragraphs(doc)))

    dxfc.fix_common_file_errors(doc, translation="UEB", remove_images=False)
    ueb_result = para.getString()
    print("UEB result (spaces around dash kept):", repr(ueb_result))
    assert ueb_result == "word - surrounded by dashes"

    doc2 = lo.open_blank_writer()
    doc2.Text.setString("word --- surrounded by dashes")
    para2 = next(iter(utils.iter_paragraphs(doc2)))
    dxfc.fix_common_file_errors(doc2, translation="EBAE", remove_images=False)
    ebae_result = para2.getString()
    print("EBAE result (spaces around dash removed):", repr(ebae_result))
    assert ebae_result == "word-surrounded by dashes"

    doc3 = lo.open_blank_writer()
    text3 = doc3.Text
    cur3 = text3.createTextCursor()
    PARA_BREAK = uno.getConstantByName("com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK")
    text3.insertString(cur3, "cat", False)
    text3.insertControlCharacter(cur3, PARA_BREAK, False)
    text3.insertString(cur3, "dog", False)
    n = dxfc.format_spelling_list(doc3, target_paragraphs=list(utils.iter_paragraphs(doc3)))
    print("spelling list result:", [p.getString() for p in utils.iter_paragraphs(doc3)])
    assert [p.getString() for p in utils.iter_paragraphs(doc3)] == ["cat  cat", "dog  dog"]

    doc.close(False); doc2.close(False); doc3.close(False)
    print("DX FILE CLEANUP TEST PASSED")
