import sys
sys.path.insert(0, ".")
sys.path.insert(0, "..")
from uno_harness import LOSession
from vistatype import config

with LOSession() as lo:
    doc = lo.open_blank_writer()
    r1 = config.set_config_new_install(doc)
    print(r1, doc.CurrentController.ViewSettings.ShowNonprintingCharacters)
    assert doc.CurrentController.ViewSettings.ShowNonprintingCharacters is False

    r2 = config.set_config_large_print(doc)
    print(r2, doc.CurrentController.ViewSettings.ShowNonprintingCharacters)
    assert doc.CurrentController.ViewSettings.ShowNonprintingCharacters is True

    r3 = config.set_config_braille(doc)
    print(r3)
    doc.close(False)
    print("CONFIG TEST PASSED")
