"""UNO controller for the LP_VISTATYPE Writer dropdown."""
from __future__ import annotations
import sys
from pathlib import Path
import unohelper
from com.sun.star.frame import XSubToolbarController, XToolbarController
from com.sun.star.lang import XInitialization

EXTENSION_DIR = Path(__file__).resolve().parent
if str(EXTENSION_DIR) not in sys.path:
    sys.path.insert(0, str(EXTENSION_DIR))

IMPLEMENTATION_NAME = "org.mrhunsaker.vistatype.LPVISTATYPEController"
SERVICE_NAME = "com.sun.star.frame.ToolbarController"
SUBTOOLBAR_URL = "private:resource/toolbar/addon_org.mrhunsaker.vistatype.LPVISTATYPE"

class LPVISTATYPEController(unohelper.Base, XToolbarController, XSubToolbarController, XInitialization):
    def __init__(self, ctx):
        self.ctx = ctx
        self.frame = None
        self.command_url = None
        self.service_manager = None

    def initialize(self, arguments):
        for arg in arguments:
            if arg.Name == "Frame": self.frame = arg.Value
            elif arg.Name == "CommandURL": self.command_url = arg.Value
            elif arg.Name == "ServiceManager": self.service_manager = arg.Value

    def execute(self, key_modifier): return None
    def click(self): return None
    def doubleClick(self): return None
    def createPopupWindow(self): return None
    def createItemWindow(self, parent): return None
    def opensSubToolbar(self): return True
    def getSubToolbarName(self): return SUBTOOLBAR_URL
    def updateImage(self): return None

    def functionSelected(self, command):
        if self.frame is None:
            return
        doc = self.frame.getController().getModel()
        desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        from vistatype.toolbar.dispatch import run_lp_command
        run_lp_command(command, doc, desktop)

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    LPVISTATYPEController, IMPLEMENTATION_NAME, (SERVICE_NAME,)
)
