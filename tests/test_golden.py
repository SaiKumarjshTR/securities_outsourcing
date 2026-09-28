"""
Golden-file regression suite for the DOCX -> SGML conversion pipeline.

Each tests/golden/<case>/ directory holds a checked-in input.docx and an
expected_TR.sgm snapshot. This test re-runs the real pipeline
(pipeline/batch_runner_deploy.py) on each input and asserts the output is
BYTE-IDENTICAL to the checked-in snapshot.

This is a *regression* safety net, not a correctness oracle: expected_TR.sgm
files are a snapshot of "last known accepted output", not a hand-verified
perfect answer (a couple of known upstream OCR defects are accepted as
out-of-scope — see business_feedback_issues.md in repo memory). If a test
fails after an intentional pipeline change, review the diff and, if the new
output is correct/better, re-bless the baseline with tests/golden/_generate.py.

Requires the same runtime environment as the rest of the pipeline (WSL2 venv
with anthropic/docx/chromadb installed) — run via:
    python3 -m pytest tests/test_golden.py -v
"""
import os
import subprocess
import sys

import pytest

_TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
_GOLDEN_DIR = os.path.join(_TESTS_DIR, "golden")
_CONVERT = os.path.join(_GOLDEN_DIR, "_convert.py")


def _discover_cases():
    if not os.path.isdir(_GOLDEN_DIR):
        return []
    return sorted(
        d for d in os.listdir(_GOLDEN_DIR)
        if os.path.isdir(os.path.join(_GOLDEN_DIR, d))
        and os.path.exists(os.path.join(_GOLDEN_DIR, d, "input.docx"))
    )


@pytest.mark.parametrize("case", _discover_cases())
def test_conversion_is_stable(case, tmp_path):
    case_dir = os.path.join(_GOLDEN_DIR, case)
    docx_path = os.path.join(case_dir, "input.docx")
    expected_path = os.path.join(case_dir, "expected_TR.sgm")
    assert os.path.exists(expected_path), (
        f"No expected_TR.sgm baseline for '{case}' — generate one with "
        f"tests/golden/_generate.py after reviewing its output."
    )

    out_path = tmp_path / f"{case}_TR.sgm"
    subprocess.run([sys.executable, _CONVERT, docx_path, str(out_path)], check=True)

    produced = out_path.read_text(encoding="utf-8")
    with open(expected_path, encoding="utf-8") as fh:
        expected = fh.read()

    assert produced == expected, (
        f"Conversion output changed for '{case}'. If this change is "
        f"intentional/correct, review the diff and re-bless with "
        f"tests/golden/_generate.py. Produced {len(produced)} chars, "
        f"expected {len(expected)} chars."
    )


def test_golden_cases_exist():
    """Guard against the suite silently collecting zero test cases."""
    assert _discover_cases(), f"No golden cases found under {_GOLDEN_DIR}"
