"""
frs14_bridge.py -- ABBYY FineReader Server 14 HTTP bridge service.

PURPOSE
-------
Runs on the Windows FRS14 machine. Exposes HTTP endpoints so Ubuntu/Plexus
can convert PDFs to DOCX without Windows COM or file-system access.

Method: Hot folder (verified working). Ticket written before PDF drop.
See FRS14_WORK_SUMMARY.md section 4 for full protocol.

ENDPOINTS
---------
  GET  /health        -- liveness / readiness check
  POST /convert       -- accept a PDF, return the converted file

POST /convert fields:
  file            -- the PDF file (multipart/form-data)
  output_format   -- 'DOCX' (default) or 'HTML'

STARTUP (run once after each Windows reboot)
--------------------------------------------
  # No extra steps needed after reboot — hot folder processes automatically.
  # Services start in correct order and hot folder is active immediately.

  # Run the bridge
  python frs14_bridge.py --host 0.0.0.0 --port 7090

CLIENT CONFIGURATION
--------------------
  export FRS14_SERVER_URL=http://<windows-host-ip>:7090
"""

import os
import uuid
import shutil
import time
import argparse
import subprocess
import tempfile
import threading
from pathlib import Path

from flask import Flask, request, jsonify, Response

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
FRS_SERVER   = os.getenv('FRS14_SERVER_HOST', 'localhost')
FRS_WORKFLOW = os.getenv('FRS14_WORKFLOW', 'Default Workflow')
FRS_STAGING  = os.getenv('FRS14_STAGING', r'C:\FRS_Input')  # SYSTEM-accessible

# Hot folder paths (verified working per FRS14_WORK_SUMMARY.md)
WF_INPUT  = os.getenv(
    'FRS14_WF_INPUT',
    r'C:\Users\Public\ABBYY\ABBYY FineReader Server 14.0\Default Workflow\Input Folder',
)
FRS_OUT   = os.getenv('FRS14_OUTPUT', r'C:\FRS_Output')

POLL_TIMEOUT  = int(os.getenv('FRS14_POLL_TIMEOUT',  '600'))
POLL_INTERVAL = float(os.getenv('FRS14_POLL_INTERVAL', '5'))

# Serialize hot-folder drops so tickets don't collide
_hf_lock = threading.Lock()

FORMAT_MAP = {
    'DOCX': {
        'ext':      'docx',
        'mimetype': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    },
    'HTML': {
        'ext':      'html',
        'mimetype': 'text/html; charset=utf-8',
    },
}

os.makedirs(FRS_STAGING, exist_ok=True)
os.makedirs(FRS_OUT, exist_ok=True)
app = Flask(__name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _frs14_service_running() -> bool:
    try:
        out = subprocess.check_output(
            ['sc', 'query', 'ABBYY.Server.FineReaderServer.14.0'],
            text=True, stderr=subprocess.DEVNULL,
        )
        return 'RUNNING' in out
    except Exception:
        return True


def _convert_via_hotfolder(pdf_path: str, output_format: str) -> bytes:
    """Convert PDF using FRS14 hot folder method (verified working).

    Protocol per FRS14_WORK_SUMMARY.md section 4:
      1. Write XML ticket to WF_INPUT FIRST
      2. Copy PDF into WF_INPUT (FRS picks it up and processes it)
      3. Poll FRS_OUT for the output file

    Uses a lock so concurrent requests don't collide on ticket/PDF names.
    """
    ext_want = FORMAT_MAP[output_format]['ext']
    uid      = uuid.uuid4().hex[:16]
    pdf_name = uid + '.pdf'
    out_stem = uid

    ticket_path = Path(WF_INPUT) / (pdf_name + '.xml')
    wf_pdf_path = Path(WF_INPUT) / pdf_name
    staged_path = Path(FRS_STAGING) / pdf_name
    out_path    = Path(FRS_OUT) / f'{out_stem}.{ext_want}'

    # Build XML ticket (per FRS14_WORK_SUMMARY.md section 4)
    ticket = (
        f'<XmlTicket>'
        f'<InputFile Name="{pdf_name}"/>'
        f'<ExportParams>'
        f'<ExportFormat OutputFileFormat="{output_format}" OutputFlowType="SharedFolder">'
        f'<OutputLocation>{FRS_OUT}</OutputLocation>'
        f'</ExportFormat>'
        f'</ExportParams>'
        f'</XmlTicket>'
    )

    with _hf_lock:
        # Stage PDF to SYSTEM-accessible path
        shutil.copy2(pdf_path, staged_path)

        # 1. Write ticket FIRST
        ticket_path.write_text(ticket, encoding='utf-8')
        print(f'[FRS14] Ticket written: {ticket_path}')
        time.sleep(0.5)

        # 2. Drop PDF into workflow input folder
        shutil.copy2(str(staged_path), str(wf_pdf_path))
        print(f'[FRS14] PDF dropped: {wf_pdf_path}')

    # 3. Poll for output (outside lock so other requests can proceed)
    deadline = time.time() + POLL_TIMEOUT
    elapsed  = 0
    while time.time() < deadline:
        if out_path.exists() and out_path.stat().st_size > 1000:
            print(f'[FRS14] Output ready: {out_path}  ({out_path.stat().st_size:,} bytes)  elapsed={elapsed}s')
            data = out_path.read_bytes()
            # Clean up output file
            try:
                out_path.unlink()
            except OSError:
                pass
            return data
        time.sleep(POLL_INTERVAL)
        elapsed += POLL_INTERVAL

    # Timed out — clean up and raise
    for p in [staged_path, wf_pdf_path, ticket_path]:
        try:
            if p.exists():
                p.unlink()
        except OSError:
            pass
    raise RuntimeError(
        f'FRS14 hot folder timed out after {POLL_TIMEOUT}s. '
        f'Expected output: {out_path}. '
        f'Check FRS14 services and Default Workflow configuration.'
    )


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route('/health')
def health():
    running   = _frs14_service_running()
    wf_ok     = os.path.isdir(WF_INPUT)
    out_ok    = os.path.isdir(FRS_OUT)
    stage_ok  = os.path.isdir(FRS_STAGING)

    status = 'ok' if (running and wf_ok and out_ok and stage_ok) else 'degraded'
    return jsonify({
        'status':          status,
        'service':         'FRS14 bridge (hot folder)',
        'frs14_service':   'running' if running else 'not running',
        'wf_input_folder': 'ok' if wf_ok else f'missing: {WF_INPUT}',
        'output_folder':   'ok' if out_ok else f'missing: {FRS_OUT}',
        'staging_folder':  'ok' if stage_ok else f'missing: {FRS_STAGING}',
        'frs14_workflow':  FRS_WORKFLOW,
    }), 200 if status == 'ok' else 503


@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'No "file" field'}), 400
    f = request.files['file']
    if not f.filename:
        return jsonify({'error': 'Empty filename'}), 400

    fmt_key = request.form.get('output_format', 'DOCX').upper()
    if fmt_key not in FORMAT_MAP:
        return jsonify({'error': f'Unsupported format "{fmt_key}"'}), 400

    fmt_info = FORMAT_MAP[fmt_key]

    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False, dir=FRS_STAGING) as tmp:
        f.save(tmp)
        tmp_pdf = tmp.name

    try:
        content  = _convert_via_hotfolder(tmp_pdf, fmt_key)
        out_name = Path(f.filename).stem + '.' + fmt_info['ext']
        return Response(
            content,
            status=200,
            mimetype=fmt_info['mimetype'],
            headers={'Content-Disposition': f'attachment; filename="{out_name}"'},
        )
    except RuntimeError as e:
        print(f'[FRS14] Error: {e}')
        return jsonify({'error': str(e)}), 500
    except Exception as e:
        print(f'[FRS14] Unexpected: {e}')
        return jsonify({'error': str(e)}), 500
    finally:
        if os.path.exists(tmp_pdf):
            try:
                os.unlink(tmp_pdf)
            except OSError:
                pass


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='FRS14 HTTP bridge (COM API)')
    parser.add_argument('--host',  default='0.0.0.0')
    parser.add_argument('--port',  type=int, default=7090)
    parser.add_argument('--debug', action='store_true')
    args = parser.parse_args()

    print('=' * 60)
    print('FRS14 HTTP bridge  (hot folder)')
    print(f'  Listening : {args.host}:{args.port}')
    print(f'  WF input  : {WF_INPUT}')
    print(f'  Output    : {FRS_OUT}')
    print(f'  Staging   : {FRS_STAGING}')
    print(f'  Timeout   : {POLL_TIMEOUT}s')
    print('=' * 60)

    if args.debug:
        app.run(host=args.host, port=args.port, debug=True)
    else:
        try:
            from waitress import serve
            print('Using waitress (multi-threaded)')
            serve(app, host=args.host, port=args.port, threads=4)
        except ImportError:
            print('waitress not installed -- using Flask dev server')
            app.run(host=args.host, port=args.port, debug=False)
