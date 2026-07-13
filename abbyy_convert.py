#!/usr/bin/env python3
"""
ABBYY FineReader Engine 12 — PDF to DOCX Converter with Diagnostics
Usage:
    python3 abbyy_convert.py <input.pdf> [output.docx]
    python3 abbyy_convert.py --diag           # diagnostics only, no conversion
"""

import subprocess
import sys
import os
import shutil
import time
import argparse
from pathlib import Path

# ── Config ──────────────────────────────────────────────────────────────────
ABBYY_BIN      = "/opt/ABBYY/FREngine12/Bin"
CLI_BIN        = f"{ABBYY_BIN}/../Samples/CommandLineInterface/CommandLineInterface"
LM_BIN         = f"{ABBYY_BIN}/LicenseManager.Console"
LICENSE_TOKEN  = "/var/lib/ABBYY/SDK/12/Licenses/SWAT12411007447140836499.ABBYY.ActivationToken"
ABBYY_ACCOUNT  = "account.abbyy.com"
LICENSING_PORT = 3023

ENV = {**os.environ, "LD_LIBRARY_PATH": ABBYY_BIN}

# ── Helpers ──────────────────────────────────────────────────────────────────
def ok(msg):   print(f"  [OK]  {msg}")
def fail(msg): print(f"  [!!]  {msg}")
def info(msg): print(f"  [--]  {msg}")
def section(title): print(f"\n{'─'*60}\n  {title}\n{'─'*60}")

# ── Diagnostics ──────────────────────────────────────────────────────────────
def check_binaries():
    section("1. Binary / installation check")
    bins = {
        "FREngine Bin dir":        ABBYY_BIN,
        "CLI converter":           CLI_BIN,
        "LicenseManager.Console":  LM_BIN,
        "LicensingService":        "/usr/local/bin/ABBYY/SDK/12/Licensing/LicensingService",
        "CodeMeter daemon":        "/usr/sbin/CodeMeterLin",
    }
    all_ok = True
    for label, path in bins.items():
        if os.path.exists(path):
            ok(f"{label}: {path}")
        else:
            fail(f"{label} NOT FOUND: {path}")
            all_ok = False
    return all_ok


def check_services():
    section("2. Service status")
    all_ok = True

    # LicensingService
    r = subprocess.run(["pgrep", "-a", "-f", "LicensingService"], capture_output=True, text=True)
    if r.returncode == 0:
        pid_line = r.stdout.strip().split("\n")[0]
        ok(f"LicensingService running  (PID {pid_line.split()[0]})")
    else:
        fail("LicensingService is NOT running")
        all_ok = False

    # CodeMeter
    r = subprocess.run(["pgrep", "-a", "-f", "CodeMeterLin"], capture_output=True, text=True)
    if r.returncode == 0:
        pid_line = r.stdout.strip().split("\n")[0]
        ok(f"CodeMeterLin running      (PID {pid_line.split()[0]})")
    else:
        fail("CodeMeterLin is NOT running")
        all_ok = False

    # Port 3023 — use ss if available, fall back to /proc/net/tcp
    port_open = False
    try:
        r = subprocess.run(["ss", "-tlnp"], capture_output=True, text=True)
        port_open = f":{LICENSING_PORT}" in r.stdout
    except FileNotFoundError:
        # ss not installed — check /proc/net/tcp directly
        hex_port = format(LICENSING_PORT, "04X")
        try:
            with open("/proc/net/tcp") as f:
                port_open = any(f":{hex_port}" in line for line in f)
        except Exception:
            pass
    if port_open:
        ok(f"LicensingService listening on port {LICENSING_PORT}")
    else:
        fail(f"Nothing listening on port {LICENSING_PORT}")
        all_ok = False

    return all_ok


def check_network():
    section("3. Network / SSL check")
    all_ok = True

    # DNS
    r = subprocess.run(["getent", "hosts", ABBYY_ACCOUNT], capture_output=True, text=True)
    if r.returncode == 0:
        ip = r.stdout.split()[0]
        ok(f"DNS {ABBYY_ACCOUNT} → {ip}")
    else:
        fail(f"DNS lookup failed for {ABBYY_ACCOUNT}")
        all_ok = False

    # HTTPS / SSL
    r = subprocess.run(
        ["curl", "-s", "-o", "/dev/null", "-w", "%{http_code} %{ssl_verify_result}",
         "--max-time", "10", f"https://{ABBYY_ACCOUNT}"],
        capture_output=True, text=True
    )
    if r.returncode == 0:
        http_code, ssl_result = r.stdout.strip().split()
        if ssl_result == "0":
            ok(f"SSL to {ABBYY_ACCOUNT} verified OK  (HTTP {http_code})")
        else:
            fail(f"SSL verification FAILED for {ABBYY_ACCOUNT}  (curl ssl_verify_result={ssl_result})")
            all_ok = False
    else:
        fail(f"curl to {ABBYY_ACCOUNT} failed: {r.stderr.strip()}")
        all_ok = False

    return all_ok


def check_license():
    section("4. License token check")
    all_ok = True

    if os.path.exists(LICENSE_TOKEN):
        size = os.path.getsize(LICENSE_TOKEN)
        mtime = time.strftime("%Y-%m-%d %H:%M", time.localtime(os.path.getmtime(LICENSE_TOKEN)))
        ok(f"Token file found: {LICENSE_TOKEN}")
        info(f"  size={size} bytes, modified={mtime}")
    else:
        fail(f"Token file NOT found: {LICENSE_TOKEN}")
        all_ok = False
        return all_ok

    # Check TR/Zscaler certs exist in /usr/local/share/ca-certificates/
    tr_certs = [
        "/usr/local/share/ca-certificates/TR_RootCA2.crt",
        "/usr/local/share/ca-certificates/TR_ZscalerIntermediate.crt",
    ]
    found = [c for c in tr_certs if os.path.exists(c)]
    if len(found) == len(tr_certs):
        ok(f"TR/Zscaler CA certs installed: {len(found)}/{len(tr_certs)} found")
    else:
        missing = [c for c in tr_certs if c not in found]
        fail(f"TR/Zscaler CA certs missing: {missing} — SSL to ABBYY may fail")
        all_ok = False

    return all_ok


def check_shm():
    """ABBYY requires /dev/shm ≥ 1GB for POSIX shared memory. Kubernetes default is 64MB."""
    section("5. /dev/shm size check (ABBYY requires ≥ 1 GB)")
    try:
        r = subprocess.run(["df", "-B1024", "/dev/shm"], capture_output=True, text=True)
        lines = r.stdout.strip().split("\n")
        if len(lines) < 2:
            fail("Cannot parse df output for /dev/shm")
            return False
        parts = lines[1].split()
        size_kb = int(parts[1])
        size_mb = size_kb // 1024
        if size_kb >= 1048576:
            ok(f"/dev/shm = {size_mb} MB  (≥ 1 GB required ✓)")
            return True
        else:
            fail(f"/dev/shm = {size_mb} MB — ABBYY requires ≥ 1 GB (1024 MB)")
            info("This will cause silent OCR failures in Kubernetes (Plexus default = 64 MB)")
            info("Fix A — docker run:      --shm-size=1g")
            info("Fix B — Kubernetes pod:  volumes: [{name: dshm, emptyDir: {medium: Memory, sizeLimit: 1Gi}}]")
            info("                         volumeMounts: [{mountPath: /dev/shm, name: dshm}]")
            info("Fix C — Plexus UI:       contact infra to configure shm for the deployment")
            return False
    except Exception as e:
        fail(f"Cannot check /dev/shm: {e}")
        return False


def run_diagnostics():
    print("\n" + "═"*60)
    print("  ABBYY FineReader Engine 12 — Diagnostics")
    print("═"*60)

    results = {
        "Binaries":  check_binaries(),
        "Services":  check_services(),
        "Network":   check_network(),
        "License":   check_license(),
        "SHM Size":  check_shm(),
    }

    section("Summary")
    all_pass = True
    for name, passed in results.items():
        if passed:
            ok(name)
        else:
            fail(name)
            all_pass = False

    if all_pass:
        print("\n  ✔  All checks passed — ABBYY is ready\n")
    else:
        print("\n  ✗  Some checks failed — see details above\n")

    return all_pass


# ── Conversion ───────────────────────────────────────────────────────────────
def convert_pdf_to_docx(input_pdf: str, output_docx: str) -> bool:
    section("PDF → DOCX conversion")

    input_path  = Path(input_pdf).resolve()
    output_path = Path(output_docx).resolve()

    if not input_path.exists():
        fail(f"Input file not found: {input_path}")
        return False

    size_mb = input_path.stat().st_size / (1024 * 1024)
    info(f"Input  : {input_path}  ({size_mb:.2f} MB)")
    info(f"Output : {output_path}")
    info(f"Engine : {CLI_BIN}")

    cmd = [
        CLI_BIN,
        "-if", str(input_path),   # input file
        "-rl", "English",         # recognition language
        "-pam", "DocumentConversion",  # page analysis mode — best for PDFs
        "-f",  "DOCX",            # export format
        "-of", str(output_path),  # output file
    ]

    info(f"Command: {' '.join(cmd)}\n")

    t0 = time.time()
    try:
        result = subprocess.run(
            cmd,
            env=ENV,
            capture_output=False,   # stream output live
            text=True,
        )
    except FileNotFoundError:
        fail(f"CLI binary not found: {CLI_BIN}")
        return False

    elapsed = time.time() - t0

    print()
    if result.returncode == 0:
        if output_path.exists():
            out_mb = output_path.stat().st_size / (1024 * 1024)
            ok(f"Conversion succeeded in {elapsed:.1f}s")
            ok(f"Output written: {output_path}  ({out_mb:.2f} MB)")
            return True
        else:
            fail("Process exited 0 but output file was not created")
            return False
    else:
        fail(f"Conversion failed — exit code {result.returncode}")
        return False


# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="ABBYY FREngine 12 — diagnostics + PDF→DOCX converter"
    )
    parser.add_argument("input",  nargs="?", help="Input PDF path")
    parser.add_argument("output", nargs="?", help="Output DOCX path (optional)")
    parser.add_argument("--diag", action="store_true", help="Run diagnostics only")
    args = parser.parse_args()

    if args.diag or args.input is None:
        ok_diag = run_diagnostics()
        if args.diag or args.input is None:
            sys.exit(0 if ok_diag else 1)

    # Derive output path if not given
    input_pdf = args.input
    if args.output:
        output_docx = args.output
    else:
        stem = Path(input_pdf).stem
        output_docx = str(Path(input_pdf).parent / f"{stem}_abbyy.docx")

    # Always run diagnostics before conversion so issues are visible
    run_diagnostics()

    # Convert
    success = convert_pdf_to_docx(input_pdf, output_docx)

    if success:
        # Copy back to Windows Downloads only when running on WSL2 host (not inside container)
        win_downloads = "/mnt/c/Users/C303180/Downloads"
        output_path = Path(output_docx)
        if os.path.isdir(win_downloads) and not str(output_path).startswith(win_downloads):
            dest = Path(win_downloads) / output_path.name
            shutil.copy2(output_path, dest)
            ok(f"Copied to Windows: {dest}")
        elif not os.path.isdir(win_downloads):
            info("Windows /mnt/c/ not available (container?) — skipping Windows copy")

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
