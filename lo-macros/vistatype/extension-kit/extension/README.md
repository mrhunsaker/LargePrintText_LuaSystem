# LP_VISTATYPE extension payload

Adds a Writer-only `LP_VISTATYPE` dropdown to the existing VISTATYPE Python package.

The dropdown uses LibreOffice's complex toolbar controller architecture: `Addons.xcu`
declares a `DropdownButton`, while a Python `XSubToolbarController` supplies the
sub-toolbar and receives the selected command in `functionSelected()`.

The existing `vistatype.entrypoints` macros are not replaced. The toolbar dispatches
directly to the existing LP modules with an explicit Writer document because a UNO
toolbar controller does not receive the Python macro provider's `XSCRIPTCONTEXT`.

This payload is source-complete but must be tested against the LibreOffice version
you deploy. The controller/XCU layer is intentionally isolated so any version-specific
registration adjustment does not require rewriting the LP logic.
