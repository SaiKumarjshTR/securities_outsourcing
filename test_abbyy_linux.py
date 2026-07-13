#!/usr/bin/env python3
"""
AbbyyLinuxConverter — diagnostic + integration test
Tests the Ubuntu ABBYY FREngine 12 as a drop-in replacement for FRS14Converter.

Usage:
    python3 test_abbyy_linux.py                          # diagnostics only
    python3 test_abbyy_linux.py <input.pdf>              # diag + convert one file
    python3 test_abbyy_linux.py --sample                 # use built-in sample PDFs
"""

import os
import sys
import time
import shutil
import subprocess
from pathlib import Path

# ── Config (matches what will go into batch_runner_deploy.py) ───────────────
ABBYY_BIN = "/opt/ABBYY/FREngine12/Bin"
ABBYY_CLI = "/opt/ABBYY/FREngine12/Samples/CommandLineInterface/CommandLineInterface"
ABBYY_ENV = {**os.environ, "LD_LIBRARY_PATH": ABBYY_BIN}

SAMPLE_PDFS = [
    "/mnt/c/Users/C303180/Downloads/input_ouput_samples/FineReader_Server_14.6.2_Results/11-349.pdf",
    "/mnt/c/Users/C303180/Downloads/input_ouput_samples/FineReader_Server_14.6.2_Results/23-101.pdf",
]

# ── Helpers ──────────────────────────────────────────────────────────────────
def ok(msg):    print(f"  ✅  {msg}")
def fail(msg):  print(f"  ❌  {msg}")
def info(msg):  print(f"  ℹ️   {msg}")
def section(t): print(f"\n{'─'*60}\n  {t}\n{'─'*60}")

# ── AbbyyLinuxConverter (the class to be embedded in batch_runner_deploy.py) ─
class AbbyyLinuxConverter:
    """
    Ubuntu ABBYY FREngine 12 — drop-in replacement for FRS14Converter.

    Uses the local ABBYY FineReader Engine 12 CLI installed at
    /opt/ABBYY/FREngine12 on Ubuntu/WSL.

    Same interface as FRS14Converter:
        initialize()
        convert_pdf_to_docx(pdf_path, docx_path) -> bool
        cleanup()
    """

    CLI  = ABBYY_CLI
    LIB  = ABBYY_BIN

    def __init__(self):
        self._env = {**os.environ, "LD_LIBRARY_PATH": self.LIB}

    def initialize(self):
        print("\n[ABBYY-Linux] Checking FREngine 12...")
        if not os.path.isfile(self.CLI):
            raise RuntimeError(f"ABBYY CLI not found: {self.CLI}")
        if not os.path.isdir(self.LIB):
            raise RuntimeError(f"ABBYY Bin dir not found: {self.LIB}")

        # Quick smoke-test: run --help (exits 0)
        r = subprocess.run(
            [self.CLI, "--help"],
            env=self._env, capture_output=True, timeout=15
        )
        if r.returncode not in (0, 1):   # --help may return 1 on some builds
            raise RuntimeError(f"ABBYY CLI smoke-test failed (rc={r.returncode})")

        # Check LicensingService
        lsvc = subprocess.run(
            ["pgrep", "-f", "LicensingService"],
            capture_output=True
        )
        if lsvc.returncode != 0:
            raise RuntimeError("ABBYY LicensingService is not running")

        print(f"   OK — ABBYY FREngine 12 ready  ({self.CLI})")

    def _run_cli(self, pdf_path: str, out_path: str, fmt: str, timeout: int = 600) -> bool:
        """Run ABBYY CLI to convert pdf_path → out_path in the given format."""
        # Write to a Linux temp path first (fast), then copy to final dest
        # This avoids the slow WSL→Windows filesystem write for large files.
        use_tmp = str(out_path).startswith("/mnt/")
        tmp_out = f"/tmp/abbyy_out_{os.getpid()}.{fmt.lower()}" if use_tmp else out_path

        cmd = [
            self.CLI,
            "-if",  pdf_path,
            "-rl",  "English",
            "-pam", "DocumentConversion",
            "-f",   fmt,
            "-of",  tmp_out,
        ]
        try:
            r = subprocess.run(cmd, env=self._env, capture_output=True,
                               text=True, timeout=timeout)
            if r.returncode != 0:
                print(f"   [ABBYY] CLI error (rc={r.returncode}): {r.stderr[-300:]}")
                return False
            if not os.path.exists(tmp_out) or os.path.getsize(tmp_out) == 0:
                print(f"   [ABBYY] CLI exited 0 but output missing/empty: {tmp_out}")
                return False
            if use_tmp:
                os.makedirs(Path(out_path).parent, exist_ok=True)
                shutil.copy2(tmp_out, out_path)
                os.unlink(tmp_out)
            size_kb = os.path.getsize(out_path) / 1024
            print(f"   OK — {fmt}: {size_kb:.1f} KB → {Path(out_path).name}")
            return True
        except subprocess.TimeoutExpired:
            print(f"   [ABBYY] Timeout ({timeout}s) converting {Path(pdf_path).name}")
            return False
        except Exception as e:
            print(f"   [ABBYY] Exception: {e}")
            return False

    def convert_pdf_to_docx(self, pdf_path: str, docx_path: str,
                             timeout: int = 600) -> bool:
        print(f"\n[ABBYY-Linux] Converting PDF → DOCX...")
        print(f"   Input: {Path(pdf_path).name}")
        return self._run_cli(pdf_path, docx_path, "DOCX", timeout)

    def convert_pdf_to_html(self, pdf_path: str, html_path: str,
                             timeout: int = 600) -> bool:
        """HTML export for footnote-anchor detection (same role as FRS14)."""
        print(f"   [ABBYY-Linux] Exporting HTML for footnote detection...")
        return self._run_cli(pdf_path, html_path, "HTMLUnicodeDefaults", timeout)

    def cleanup(self):
        pass  # No resources to release


# ── Diagnostics ──────────────────────────────────────────────────────────────
def run_diagnostics() -> bool:
    print("\n" + "═"*60)
    print("  AbbyyLinuxConverter — Integration Diagnostic")
    print("═"*60)

    all_ok = True

    section("1. Binary check")
    for label, path in [
        ("CLI binary",        ABBYY_CLI),
        ("Bin directory",     ABBYY_BIN),
        ("LicensingService",  "/usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService"),
        ("CodeMeterLin",      "/usr/sbin/CodeMeterLin"),
    ]:
        if os.path.exists(path):
            ok(f"{label}: {path}")
        else:
            fail(f"{label} NOT FOUND: {path}")
            all_ok = False

    section("2. Services")
    for svc, pattern in [("LicensingService", "LicensingService"), ("CodeMeterLin", "CodeMeterLin")]:
        r = subprocess.run(["pgrep", "-a", "-f", pattern], capture_output=True, text=True)
        if r.returncode == 0:
            pid = r.stdout.split()[0]
            ok(f"{svc} running (PID {pid})")
        else:
            fail(f"{svc} NOT running")
            all_ok = False

    section("3. Network / SSL")
    r = subprocess.run(
        ["curl", "-s", "-o", "/dev/null", "-w", "%{http_code} %{ssl_verify_result}",
         "--max-time", "8", "https://account.abbyy.com"],
        capture_output=True, text=True
    )
    if r.returncode == 0:
        parts = r.stdout.strip().split()
        http_code, ssl_ok = parts[0], parts[1]
        if ssl_ok == "0":
            ok(f"account.abbyy.com reachable (HTTP {http_code}, SSL verified)")
        else:
            fail(f"SSL verification failed for account.abbyy.com (ssl_result={ssl_ok})")
            all_ok = False
    else:
        fail("curl to account.abbyy.com failed")
        all_ok = False

    section("4. AbbyyLinuxConverter.initialize()")
    try:
        conv = AbbyyLinuxConverter()
        conv.initialize()
        ok("initialize() passed")
    except Exception as e:
        fail(f"initialize() raised: {e}")
        all_ok = False

    section("5. FRS14Converter compatibility check")
    required = ["initialize", "convert_pdf_to_docx", "convert_pdf_to_html", "cleanup"]
    conv = AbbyyLinuxConverter()
    for method in required:
        if hasattr(conv, method) and callable(getattr(conv, method)):
            ok(f"Method present: {method}()")
        else:
            fail(f"Method MISSING: {method}()")
            all_ok = False

    section("Summary")
    if all_ok:
        ok("All checks passed — AbbyyLinuxConverter ready for pipeline")
    else:
        fail("Some checks failed — fix before replacing FRS14Converter")

    return all_ok


# ── Conversion test ───────────────────────────────────────────────────────────
def run_conversion_test(pdf_paths: list) -> bool:
    section("Conversion Test")
    out_dir = "/tmp/abbyy_linux_test"
    os.makedirs(out_dir, exist_ok=True)

    conv = AbbyyLinuxConverter()
    try:
        conv.initialize()
    except Exception as e:
        fail(f"initialize() failed: {e}")
        return False

    results = []
    for pdf in pdf_paths:
        if not os.path.exists(pdf):
            fail(f"Input not found: {pdf}")
            results.append(False)
            continue

        stem = Path(pdf).stem
        docx_out = os.path.join(out_dir, f"{stem}_linux_abbyy.docx")

        info(f"Converting: {Path(pdf).name}  ({os.path.getsize(pdf)/1024:.0f} KB)")
        t0 = time.time()
        success = conv.convert_pdf_to_docx(pdf, docx_out)
        elapsed = time.time() - t0

        if success:
            size_kb = os.path.getsize(docx_out) / 1024
            ok(f"Done in {elapsed:.1f}s → {docx_out} ({size_kb:.0f} KB)")
            results.append(True)
        else:
            fail(f"Conversion failed for {Path(pdf).name}")
            results.append(False)

    conv.cleanup()

    passed = sum(results)
    print(f"\n  Conversion: {passed}/{len(results)} passed")
    print(f"  Output dir: {out_dir}")
    return all(results)


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    diag_ok = run_diagnostics()

    if "--diag" in sys.argv or (len(sys.argv) == 1 and not "--sample" in sys.argv):
        sys.exit(0 if diag_ok else 1)

    if "--sample" in sys.argv:
        pdfs = [p for p in SAMPLE_PDFS if os.path.exists(p)]
        if not pdfs:
            fail("No sample PDFs found — check SAMPLE_PDFS paths in script")
            sys.exit(1)
    else:
        pdfs = [sys.argv[1]]

    conv_ok = run_conversion_test(pdfs)
    sys.exit(0 if (diag_ok and conv_ok) else 1)


if __name__ == "__main__":
    main()
