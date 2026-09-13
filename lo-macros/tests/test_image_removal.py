import sys, os
sys.path.insert(0, ".")
sys.path.insert(0, "..")
import uno
from uno_harness import LOSession
from vistatype.lp import file_cleanup

with LOSession() as lo:
    doc = lo.open_blank_writer()
    text = doc.Text
    text.setString("Text with an image below.")

    # create a simple graphic object placeholder (no real image bytes needed
    # to test the removal path -- an empty GraphicObject still registers in
    # doc.GraphicObjects)
    graphic = doc.createInstance("com.sun.star.text.TextGraphicObject")
    from com.sun.star.awt import Size
    graphic.setSize(Size(2000, 2000))
    text.insertTextContent(text.End, graphic, False)

    print("graphics before:", list(doc.GraphicObjects.ElementNames))
    result = file_cleanup.fix_common_file_errors(doc, remove_images=True)
    print("result:", result)
    print("graphics after:", list(doc.GraphicObjects.ElementNames))
    assert result["images_found"] == 1
    assert result["images_removed"] == 1
    assert list(doc.GraphicObjects.ElementNames) == []
    doc.close(False)
    print("IMAGE REMOVAL TEST PASSED")
