#!/usr/bin/env python3

from pathlib import Path
import shutil
import tempfile
import zipfile


KIT = Path(__file__).resolve().parents[1]
VISTATYPE = KIT.parent
EXT = KIT / "extension"
OUT = KIT / "dist" / "vistatype-lp.oxt"


def main():
    if not (VISTATYPE / "entrypoints.py").exists():
        raise SystemExit(f"Expected canonical package at {VISTATYPE}")

    OUT.parent.mkdir(parents=True, exist_ok=True)

    if OUT.exists():
        OUT.unlink()

    with tempfile.TemporaryDirectory() as t:
        s = Path(t)

        for name in (
            "Addons.xcu",
            "Controller.xcu",
            "description.xml",
            "README.md",
            "lp_vistatype_controller.py",
        ):
            shutil.copy2(EXT / name, s / name)

        shutil.copytree(
            EXT / "META-INF",
            s / "META-INF",
        )

        shutil.copytree(
            VISTATYPE,
            s / "vistatype",
            ignore=shutil.ignore_patterns(
                "__pycache__",
                "*.pyc",
                "*.pyo",
                "extension-kit",
            ),
        )

        # The canonical package does not contain the toolbar directory.
        # Create it before overlaying the extension-specific toolbar modules.
        toolbar_dst = s / "vistatype" / "toolbar"
        toolbar_dst.mkdir(parents=True, exist_ok=True)

        for p in (EXT / "vistatype" / "toolbar").glob("*.py"):
            shutil.copy2(
                p,
                toolbar_dst / p.name,
            )

        with zipfile.ZipFile(
            OUT,
            "w",
            zipfile.ZIP_DEFLATED,
        ) as z:
            for p in sorted(s.rglob("*")):
                if p.is_file():
                    z.write(
                        p,
                        p.relative_to(s).as_posix(),
                    )

    print(f"Built {OUT}")


if __name__ == "__main__":
    main()
