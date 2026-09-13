#!/usr/bin/env bash
# Rebuilds the reference doc and exercises each Lua filter against its
# fixture, printing plain-text output for a quick eyeball check. This is a
# smoke test, not an automated pass/fail suite (a couple of the checks --
# e.g. confirming the real "Print Pg Num" style properties took effect --
# need LibreOffice's object model to verify properly; see
# lo-macros/tests/test_reference_pages.py for that level of check).
set -euo pipefail
cd "$(dirname "$0")"
REF=../../reference-doc/LargePrintTemplate-reference.odt
FILTERS=..

if [ ! -f "$REF" ]; then
    echo "Building reference doc first..."
    (cd ../../reference-doc && ./build_reference_odt.sh \
        /mnt/user-data/uploads/LargePrintTemplate.dotx LargePrintTemplate-reference.odt)
fi

echo "=== fix-common-file-errors (smart quotes/dashes OFF, to see raw collapsing) ==="
pandoc fix-errors-sample.md -f markdown-smart -t plain \
    --lua-filter="$FILTERS/filters/lp-fix-common-file-errors.lua"
echo

echo "=== delete-multiple-paragraph-marks + fix-common-file-errors on sample.md ==="
pandoc sample.md -t plain \
    --lua-filter="$FILTERS/filters/lp-fix-common-file-errors.lua" \
    --lua-filter="$FILTERS/filters/lp-delete-multiple-paragraph-marks.lua" \
    --lua-filter="$FILTERS/filters/lp-horizontal-list-to-vertical.lua"
echo

echo "=== reference page number tag + format, converted all the way to ODT ==="
pandoc multipara.md -o /tmp/vistatype_pipeline_check.odt \
    --reference-doc="$REF" \
    --lua-filter="$FILTERS/filters/lp-fix-common-file-errors.lua" \
    --lua-filter="$FILTERS/filters/lp-delete-multiple-paragraph-marks.lua" \
    --lua-filter="$FILTERS/filters/lp-tag-reference-page-numbers.lua" \
    --lua-filter="$FILTERS/filters/lp-format-tagged-page-numbers.lua"
echo "wrote /tmp/vistatype_pipeline_check.odt -- open it in LibreOffice to eyeball the result."
