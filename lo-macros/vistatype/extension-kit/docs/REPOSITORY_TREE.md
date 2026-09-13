# Proposed repository layout

```text
lo-macros/vistatype/
├── existing Python package (unchanged)
└── extension-kit/
    ├── README.md
    ├── extension/
    │   ├── Addons.xcu
    │   ├── Controller.xcu
    │   ├── description.xml
    │   ├── lp_vistatype_controller.py
    │   ├── README.md
    │   ├── META-INF/manifest.xml
    │   └── vistatype/toolbar/
    │       ├── __init__.py
    │       ├── commands.py
    │       └── dispatch.py
    ├── scripts/
    │   ├── build_oxt.py
    │   ├── install_oxt.py
    │   └── validate_kit.py
    └── docs/REPOSITORY_TREE.md
```

The build script reuses the canonical parent `vistatype/` package; do not maintain a duplicate package.
