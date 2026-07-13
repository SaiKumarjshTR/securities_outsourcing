"""
abbyy_linux_bridge.py — ABBYY FREngine 12 HTTP bridge (Ubuntu/Linux)

PURPOSE
-------
Ubuntu replacement for frs14_bridge.py (Windows FRS14 bridge).
Runs on the Ubuntu host. Exposes the same HTTP API so the Plexus
Docker container can convert PDFs to DOCX without any Windows dependency.

ENDPOINTS
---------
  GET  /health    — liveness check, returns {"status":"ok","abbyy":"FREngine12"}
  POST /convert   — accept PDF, return converted file (DOCX or HTML)

POST /convert fields:
  file            — PDF file (multipart/form-data)
  output_format   — 'DOCX' (default) or 'HTML'

STARTUP (run once on Ubuntu host)
----------------------------------
  python3 pipeline/abbyy_linux_bridge.py --host 0.0.0.0 --port 7091

  Or as a background service:
  nohup python3 pipeline/abbyy_linux_bridge.py --host 0.0.0.0 --port 7091 \
      > /var/log/abbyy_bridge.log 2>&1 &

CLIENT CONFIGURATION (Docker / Plexus)
---------------------------------------
  export FRS14_SERVER_URL=http://<ubuntu-host-ip>:7091

  In Dockerfile:
  ENV FRS14_SERVER_URL=http://172.25.16.1:7091   # WSL host gateway

REQUIREMENTS
------------
  pip install flask
  ABBYY FREngine 12 installed at /opt/ABBYY/FREngine12
"""

import os
import shutil
import subprocess
import tempfile
import argparse
import threading
from pathlib import Path

from flask import Flask, request, jsonify, Response

# ── Config ────────────────────────────────────────────────────────────────────
ABBYY_CLI = os.getenv('ABBYY_CLI', '/opt/ABBYY/FREngine12/Samples/CommandLineInterface/CommandLineInterface')
ABBYY_LIB = os.getenv('ABBYY_LIB', '/opt/ABBYY/FREngine12/Bin')
ABBYY_ENV = {**os.environ, 'LD_LIBRARY_PATH': ABBYY_LIB}
ABBYY_TIMEOUT = int(os.getenv('ABBYY_TIMEOUT', '600'))

# Serialize conversions (ABBYY CLI is not thread-safe for concurrent runs)
_convert_lock = threading.Lock()

FORMAT_MAP = {
    'DOCX': {
        'abbyy_fmt': 'DOCX',
        'ext':       'docx',
        'mimetype':  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    },
    'HTML': {
        'abbyy_fmt': 'HTMLUnicodeDefaults',
        'ext':       'html',
        'mimetype':  'text/html; charset=utf-8',
    },
}

app = Flask(__name__)


def _check_abbyy() -> bool:
    """Quick check that ABBYY CLI and LicensingService are available."""
    if not os.path.isfile(ABBYY_CLI):
        return False
    try:
        r = subprocess.run(['pgrep', '-f', 'LicensingService'], capture_output=True)
        return r.returncode == 0
    except (FileNotFoundError, OSError):
        # pgrep (procps) not installed — assume LicensingService is running
        return True


@app.route('/health', methods=['GET'])
def health():
    ok = _check_abbyy()
    return jsonify({
        'status':  'ok' if ok else 'degraded',
        'abbyy':   'FREngine12',
        'cli':     ABBYY_CLI,
        'license': 'LicensingService running' if ok else 'LicensingService NOT found',
    }), 200 if ok else 503


@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    output_format = request.form.get('output_format', 'DOCX').upper()
    if output_format not in FORMAT_MAP:
        return jsonify({'error': f'Unsupported format: {output_format}. Use DOCX or HTML'}), 400

    fmt = FORMAT_MAP[output_format]
    pdf_file = request.files['file']

    with tempfile.TemporaryDirectory(prefix='abbyy_bridge_') as tmpdir:
        # Save uploaded PDF to temp dir (Linux FS — fast I/O)
        pdf_path = os.path.join(tmpdir, 'input.pdf')
        out_path = os.path.join(tmpdir, f'output.{fmt["ext"]}')
        pdf_file.save(pdf_path)

        cmd = [
            ABBYY_CLI,
            '-if',  pdf_path,
            '-rl',  'English',
            '-pam', 'DocumentConversion',
            '-f',   fmt['abbyy_fmt'],
            '-of',  out_path,
        ]

        with _convert_lock:
            try:
                r = subprocess.run(
                    cmd, env=ABBYY_ENV,
                    capture_output=True, text=True,
                    timeout=ABBYY_TIMEOUT,
                )
            except subprocess.TimeoutExpired:
                return jsonify({'error': f'Conversion timeout ({ABBYY_TIMEOUT}s)'}), 504
            except Exception as e:
                return jsonify({'error': str(e)}), 500

        if r.returncode != 0:
            return jsonify({
                'error': f'ABBYY CLI failed (rc={r.returncode})',
                'detail': r.stderr[-500:],
            }), 500

        if not os.path.exists(out_path) or os.path.getsize(out_path) == 0:
            return jsonify({'error': 'Conversion produced empty output'}), 500

        result_bytes = Path(out_path).read_bytes()

    print(f'[bridge] Converted {pdf_file.filename} → {output_format} ({len(result_bytes)//1024} KB)')
    return Response(
        result_bytes,
        mimetype=fmt['mimetype'],
        headers={'Content-Disposition': f'attachment; filename="output.{fmt["ext"]}"'},
    )


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='ABBYY FREngine 12 HTTP bridge')
    parser.add_argument('--host', default='0.0.0.0')
    parser.add_argument('--port', type=int, default=7091)
    args = parser.parse_args()

    if not _check_abbyy():
        print(f'WARNING: ABBYY CLI not found or LicensingService not running')
        print(f'  CLI path: {ABBYY_CLI}')
        print(f'  Starting anyway...')
    else:
        print(f'ABBYY FREngine 12 bridge ready')

    print(f'Listening on http://{args.host}:{args.port}')
    app.run(host=args.host, port=args.port, threaded=True)
