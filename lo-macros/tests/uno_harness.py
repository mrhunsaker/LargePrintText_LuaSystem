"""
Minimal helper for connecting to a headless soffice instance over UNO.
Used by all the test scripts in this directory so each one doesn't have to
re-implement the boilerplate connection/retry logic.
"""
import subprocess
import time
import os
import signal

import uno
from com.sun.star.beans import PropertyValue


SOFFICE_PORT = 2002
USER_INSTALLATION = "/tmp/lo_test_profile"


def start_soffice():
    os.makedirs(USER_INSTALLATION, exist_ok=True)
    proc = subprocess.Popen(
        [
            "soffice",
            "--headless",
            "--invisible",
            "--nocrashreport",
            "--nodefault",
            "--norestore",
            "--nologo",
            "--nofirststartwizard",
            f"--accept=socket,host=localhost,port={SOFFICE_PORT};urp;",
            f"-env:UserInstallation=file://{USER_INSTALLATION}",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    return proc


def connect(retries=30, delay=0.5):
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    last_exc = None
    for _ in range(retries):
        try:
            ctx = resolver.resolve(
                f"uno:socket,host=localhost,port={SOFFICE_PORT};"
                "urp;StarOffice.ComponentContext"
            )
            return ctx
        except Exception as e:  # noqa: BLE001
            last_exc = e
            time.sleep(delay)
    raise RuntimeError(f"Could not connect to soffice: {last_exc}")


def make_prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


class LOSession:
    """Context manager: starts soffice, connects, gives you the desktop."""

    def __enter__(self):
        self.proc = start_soffice()
        self.ctx = connect()
        smgr = self.ctx.ServiceManager
        self.desktop = smgr.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx
        )
        return self

    def open_blank_writer(self):
        return self.desktop.loadComponentFromURL(
            "private:factory/swriter", "_blank", 0, (make_prop("Hidden", True),)
        )

    def open(self, path):
        url = "file://" + os.path.abspath(path)
        return self.desktop.loadComponentFromURL(
            url, "_blank", 0, (make_prop("Hidden", True),)
        )

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            self.desktop.terminate()
        except Exception:  # noqa: BLE001
            pass
        try:
            self.proc.send_signal(signal.SIGTERM)
            self.proc.wait(timeout=5)
        except Exception:  # noqa: BLE001
            self.proc.kill()
