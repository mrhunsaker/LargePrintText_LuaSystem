#!/usr/bin/env bash
# Runs every test in this directory against a fresh headless LibreOffice
# instance each time (each test starts and stops its own soffice via
# uno_harness.LOSession, so they're independent and safe to run in any
# order). Requires `soffice` on PATH.
set -uo pipefail
cd "$(dirname "$0")"

pass=0
fail=0
for f in test_*.py; do
    echo "=== $f ==="
    if python3 "$f"; then
        pass=$((pass+1))
    else
        fail=$((fail+1))
        echo "!!! FAILED: $f"
    fi
    echo
done

echo "-----------------------------------"
echo "Passed: $pass   Failed: $fail"
[ "$fail" -eq 0 ]
