#!/usr/bin/env bash
# Builds LargePrintTemplate-reference.odt from LargePrintTemplate.dotx.
#
# Pandoc's `--reference-doc` mechanism works by copying styles (and, for ODT,
# page layout / master page geometry) out of a reference file. Handing it a
# .dotx directly does not work reliably across Pandoc versions, so we first
# let a real copy of Writer do the OOXML->ODF style translation (which we've
# verified preserves style names, fonts, run-level character spacing, and
# page geometry), and use *that* .odt as the reference doc.
#
# Usage:
#   ./build_reference_odt.sh /path/to/LargePrintTemplate.dotx ./LargePrintTemplate-reference.odt
#
# Requires: soffice (LibreOffice) on PATH, headless-capable.

set -euo pipefail

SRC="${1:?Usage: build_reference_odt.sh <LargePrintTemplate.dotx> <out.odt>}"
OUT="${2:?Usage: build_reference_odt.sh <LargePrintTemplate.dotx> <out.odt>}"
OUTDIR="$(dirname "$(realpath "$OUT")")"
OUTNAME="$(basename "$OUT")"

TMPDIR="$(mktemp -d)"
trap 'rm -rf "$TMPDIR"' EXIT

# soffice names the output after the input; convert into a scratch dir then move.
soffice --headless --norestore --convert-to odt --outdir "$TMPDIR" "$SRC"

CONVERTED="$TMPDIR/$(basename "${SRC%.*}").odt"
if [ ! -f "$CONVERTED" ]; then
    echo "ERROR: expected converted file not found at $CONVERTED" >&2
    exit 1
fi

mkdir -p "$OUTDIR"
mv "$CONVERTED" "$OUTDIR/$OUTNAME"
echo "Wrote $OUTDIR/$OUTNAME"
echo
echo "Verifying key styles survived the conversion..."
python3 - "$OUTDIR/$OUTNAME" <<'PYEOF'
import sys, zipfile, re
path = sys.argv[1]
with zipfile.ZipFile(path) as z:
    styles = z.read("styles.xml").decode("utf-8")
    content = z.read("content.xml").decode("utf-8") if "content.xml" in z.namelist() else ""

expect = ["Print_20_Pg_20_Num", "Para_20_Black", "Words_20_Black",
          "Box_20_Black", "Box_20_Red", "List_20_Paragraph"]
for name in expect:
    ok = f'style:name="{name}"' in styles
    print(f'  [{"OK" if ok else "MISSING"}] {name}')

# "Normal" is ODF's "Standard" style; Word's display name survives as a comment
# only in some producers, so check the underlying properties instead.
m = re.search(r'<style:style style:name="Standard".*?</style:style>', styles, re.S)
if m and 'fo:font-size="18pt"' in m.group(0):
    print('  [OK] Normal (Standard): 18pt confirmed')
else:
    print('  [CHECK] Normal (Standard): could not confirm 18pt base size')

# Page geometry lives on whichever page-layout the "Standard" master page uses,
# not necessarily the first <style:page-layout> in the file.
mp = re.search(r'<style:master-page style:name="Standard"[^>]*style:page-layout-name="([^"]+)"', styles)
if mp:
    layout_name = mp.group(1)
    pl = re.search(r'<style:page-layout style:name="' + re.escape(layout_name) + r'">(.*?)</style:page-layout>', styles, re.S)
    if pl:
        body = pl.group(1)
        w = re.search(r'fo:page-width="([^"]+)"', body)
        h = re.search(r'fo:page-height="([^"]+)"', body)
        t = re.search(r'fo:margin-top="([^"]+)"', body)
        print(f'  Page size: {w.group(1) if w else "?"} x {h.group(1) if h else "?"}, top margin: {t.group(1) if t else "?"}')
PYEOF
