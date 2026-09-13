# LP_VISTATYPE LibreOffice implementation kit

## Install into the repository

Copy this bundle into `lo-macros/vistatype/extension-kit/`.

Then run from that directory:

```bash
python3 scripts/validate_kit.py
python3 scripts/build_oxt.py
python3 scripts/install_oxt.py
```

The build script reuses the existing `lo-macros/vistatype` package and overlays only
the new toolbar integration files. Existing macro entrypoints remain unchanged.

## What is implemented

- Writer-only `LP_VISTATYPE` toolbar control.
- `DropdownButton` complex toolbar control.
- Python UNO toolbar controller using `XToolbarController`, `XSubToolbarController`, and `XInitialization`.
- Explicit-document dispatcher for all 13 existing LP commands.
- Cross-platform OXT build/install scripts.
- Static XML/package validation.

## Expected menu contents

1. Attach LP Template
2. File Fix Sequence
3. Selected File Cleanup
4. AutoTag Page Numbers
5. Manual Tag with $pg
6. Format Page Numbers
7. Validate $pg
8. Horizontal List → Vertical
9. Compress Linear Math
10. Toggle Space After Paragraph
11. Keep With Next Paragraph
12. Type Fill-In Line
13. Format Exercise Levels

## Testing boundary

This is an implementation artifact, not a claim of successful execution on every LibreOffice release.
The LibreOffice API documents `XSubToolbarController` specifically for toolbar button/sub-toolbar behavior,
and the SDK contains a complex-toolbar example. The first live test should be performed with the LibreOffice
version you intend to deploy. If controller registration differs on that version, the adjustment should be
confined to `Controller.xcu`/the controller layer.
