#!/usr/bin/env python3
"""
(Re)generate golden-file baselines for tests/test_golden.py.

Run this deliberately after a REVIEWED, INTENTIONAL pipeline output change:
    python3 tests/golden/_generate.py

Do NOT run this just to make a failing test pass without first reviewing the
diff between the old and new expected_TR.sgm — that diff IS the change you
are about to bless as "correct". `git diff tests/golden/` after running this
shows exactly what changed.
"""
import os
import subprocess
import sys

_GOLDEN_DIR = os.path.dirname(os.path.abspath(__file__))
_CONVERT = os.path.join(_GOLDEN_DIR, "_convert.py")

cases = sorted(
    d for d in os.listdir(_GOLDEN_DIR)
    if os.path.isdir(os.path.join(_GOLDEN_DIR, d)) and not d.startswith("_")
)

failures = []
for stem in cases:
    case_dir = os.path.join(_GOLDEN_DIR, stem)
    docx_path = os.path.join(case_dir, "input.docx")
    if not os.path.exists(docx_path):
        print(f"[SKIP] {stem}: no input.docx")
        continue
    expected_path = os.path.join(case_dir, "expected_TR.sgm")
    print(f"[GENERATE] {stem} ...")
    try:
        subprocess.run([sys.executable, _CONVERT, docx_path, expected_path], check=True)
    except subprocess.CalledProcessError as exc:
        print(f"[FAIL] {stem}: {exc}")
        failures.append(stem)

if failures:
    print(f"\n{len(failures)} case(s) FAILED to generate: {failures}")
    sys.exit(1)
print("\nDone. Review changes with `git diff tests/golden/` before committing.")
