import sys
sys.path.insert(0, ".")
sys.path.insert(0, "..")
import uno
from uno_harness import LOSession
from vistatype import utils
from vistatype.lp import file_cleanup

with LOSession() as lo:
    doc = lo.open_blank_writer()
    text = doc.Text
    text.setString("Some body text.")

    # add a text frame with content
    frame = doc.createInstance("com.sun.star.text.TextFrame")
    from com.sun.star.awt import Size
    frame.setSize(Size(5000, 3000))
    text.insertTextContent(text.End, frame, False)
    frame.Text.setString("Text inside a frame that should be preserved.")

    # set a drop cap on the first paragraph
    para = next(iter(utils.iter_paragraphs(doc)))
    fmt = para.DropCapFormat
    fmt.Lines = 3
    fmt.Count = 1
    para.DropCapFormat = fmt

    print("frames before:", list(doc.TextFrames.ElementNames))
    print("drop cap count before:", para.DropCapFormat.Count)

    result = file_cleanup.fix_common_file_errors(doc, remove_images=False)
    print("result:", result)

    print("frames after:", list(doc.TextFrames.ElementNames))
    paras_after = [p.getString() for p in utils.iter_paragraphs(doc)]
    print("paragraphs after:", paras_after)
    # re-fetch para since setString/DropCapFormat changes might have created new proxy
    para2 = next(iter(utils.iter_paragraphs(doc)))
    print("drop cap count after:", para2.DropCapFormat.Count)

    assert list(doc.TextFrames.ElementNames) == []
    assert any("preserved" in p for p in paras_after)
    assert para2.DropCapFormat.Count == 0

    doc.close(False)
    print("FRAME/DROPCAP TEST PASSED")
