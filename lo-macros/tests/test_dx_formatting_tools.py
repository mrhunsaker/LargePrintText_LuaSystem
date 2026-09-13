import sys
sys.path.insert(0, ".")
sys.path.insert(0, "..")
from uno_harness import LOSession
from vistatype import utils
from vistatype.dx import formatting_tools as dxft

with LOSession() as lo:
    doc = lo.open_blank_writer()
    text = doc.Text
    text.setString("y = 2 x + 1")
    para = next(iter(utils.iter_paragraphs(doc)))
    dxft.compress_linear_math(doc, target=para)
    result = para.getString()
    print("braille compress:", repr(result))
    assert result == "y = 2x+1", repr(result)  # spaces KEPT around '=' only

    doc2 = lo.open_blank_writer()
    vc = doc2.CurrentController.ViewCursor
    dxft.insert_dash(doc2, "em_dash")
    dxft.insert_prime(doc2, "single")
    dxft.insert_fraction(doc2, 1, 2)       # compact
    dxft.insert_fraction(doc2, 132, 65)    # not compact -> DBT coded (int example)
    print("inserted:", repr(doc2.Text.getString()))
    assert doc2.Text.getString() == "\u2014\u2032\u00BD[[*fs*]]132[[*fl*]]65[[*fe*]]"

    doc3 = lo.open_blank_writer()
    doc3.Text.setString("The fraction 132.65/567.9 appears here, not a date like 12/25.")
    para3 = next(iter(utils.iter_paragraphs(doc3)))
    dxft.encode_fractions_in_selection(doc3, target=para3)
    print("encoded:", para3.getString())
    assert "[[*fs*]]132.65[[*fl*]]567.9[[*fe*]]" in para3.getString()
    assert "[[*fs*]]12[[*fl*]]25[[*fe*]]" in para3.getString()  # caller's job to avoid dates, not the function's

    doc.close(False); doc2.close(False); doc3.close(False)
    print("DX FORMATTING TOOLS TEST PASSED")
