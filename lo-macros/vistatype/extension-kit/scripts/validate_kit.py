from pathlib import Path
import xml.etree.ElementTree as ET
R=Path(__file__).resolve().parents[1]
for rel in ["extension/Addons.xcu","extension/Controller.xcu","extension/description.xml","extension/META-INF/manifest.xml","extension/lp_vistatype_controller.py","extension/vistatype/toolbar/commands.py","extension/vistatype/toolbar/dispatch.py"]:
    p=R/rel
    if not p.exists(): raise SystemExit(f"Missing {rel}")
    if p.suffix==".xcu" or p.name=="description.xml" or p.name=="manifest.xml": ET.parse(p)
print("Kit validation passed.")
