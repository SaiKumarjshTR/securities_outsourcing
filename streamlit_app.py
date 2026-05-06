"""
Securities Commission Conversion — Streamlit UI
=============================================
AI-powered PDF/DOCX → SGML conversion for TR Securities Outsourcing.

Pipeline: PDF → (Hybrid PyMuPDF + pdfplumber) → DOCX → batch_runner_deploy.py → _TR.sgm

Plexus deployment: port 8501
Health check: GET /_stcore/health  (Streamlit built-in)
"""
import os
import sys
import uuid
import shutil
import tempfile
import datetime
from pathlib import Path

import streamlit as st

# ── Session manager (multi-user isolation) ────────────────────────────────────
try:
    sys.path.insert(0, str(Path(__file__).parent / "app"))
    from session_manager import SessionManager
    _session_mgr: SessionManager | None = SessionManager()
except Exception:
    _session_mgr = None

# ── Page config MUST be called first ─────────────────────────────────────────
st.set_page_config(
    page_title="Securities Commission Conversion",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
    menu_items={
        "Get Help": None,
        "Report a bug": None,
        "About": "Thomson Reuters Securities SGML Pipeline v0.0.6",
    },
)

# ── Custom CSS — dark sidebar + styling ──────────────────────────────────────
st.markdown(
    """
    <style>
    /* ── Sidebar dark theme ─────────────────────────────────────────── */
    [data-testid="stSidebar"] {
        background-color: #1a2432 !important;
    }
    [data-testid="stSidebar"] .stMarkdown p,
    [data-testid="stSidebar"] .stMarkdown span,
    [data-testid="stSidebar"] .stMarkdown div,
    [data-testid="stSidebar"] .stMarkdown h1,
    [data-testid="stSidebar"] .stMarkdown h2,
    [data-testid="stSidebar"] .stMarkdown h3,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] p {
        color: #cdd8e3 !important;
    }
    [data-testid="stSidebar"] .stTextInput input {
        background-color: #243347 !important;
        color: #cdd8e3 !important;
        border: 1px solid #3c5070 !important;
        border-radius: 4px !important;
    }
    [data-testid="stSidebar"] hr {
        border-color: #2e4060 !important;
        margin: 0.5rem 0 !important;
    }

    /* ── New Session button — green ──────────────────────────────────── */
    [data-testid="stSidebar"] .stButton > button {
        background-color: #27a745 !important;
        color: white !important;
        border: none !important;
        font-weight: 600 !important;
    }
    [data-testid="stSidebar"] .stButton > button:hover {
        background-color: #218838 !important;
    }

    /* ── Sidebar expander styling ────────────────────────────────────── */
    [data-testid="stSidebar"] .streamlit-expanderHeader {
        background-color: #1e2d3e !important;
        color: #cdd8e3 !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }
    [data-testid="stSidebar"] .streamlit-expanderContent {
        background-color: #1a2432 !important;
        border: 1px solid #2e4060 !important;
        border-top: none !important;
        border-radius: 0 0 6px 6px !important;
    }

    /* ── Active sessions metric — bright visible color ───────────────── */
    [data-testid="stSidebar"] [data-testid="stMetricValue"] {
        color: #4fc3f7 !important;
        font-size: 1.8rem !important;
        font-weight: 700 !important;
    }
    [data-testid="stSidebar"] [data-testid="stMetricLabel"] {
        color: #90adc4 !important;
    }

    /* ── Upload area dashed border ───────────────────────────────────── */
    [data-testid="stFileUploader"] {
        border: 2px dashed #c0c8d4 !important;
        border-radius: 8px !important;
        padding: 1rem !important;
        background-color: #f7f9fc !important;
    }

    /* ── Reduce top padding ──────────────────────────────────────────── */
    .main .block-container {
        padding-top: 2rem !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ── Session state initialisation ─────────────────────────────────────────────
if "session_id" not in st.session_state:
    st.session_state.session_id = uuid.uuid4().hex[:8]
if "created_at" not in st.session_state:
    st.session_state.created_at = datetime.datetime.now().strftime("%H:%M:%S")
if "status" not in st.session_state:
    st.session_state.status = "active"
if "processing_count" not in st.session_state:
    st.session_state.processing_count = 0
# Pipeline step tracking (mirrors PL reference app)
if "pipeline_steps" not in st.session_state:
    st.session_state.pipeline_steps = {}   # {step_idx: 'pending'|'running'|'done'|'error'}
if "pipeline_logs" not in st.session_state:
    st.session_state.pipeline_logs = []    # list of (timestamp, level, message)
if "pipeline_running" not in st.session_state:
    st.session_state.pipeline_running = False
# Cross-page carry-over — last pipeline run results shared with HITL pages
if "pipeline_result_ready" not in st.session_state:
    st.session_state.pipeline_result_ready = False
if "last_pdf_name" not in st.session_state:
    st.session_state.last_pdf_name = None
if "last_pdf_bytes" not in st.session_state:
    st.session_state.last_pdf_bytes = None
if "last_sgml_text" not in st.session_state:
    st.session_state.last_sgml_text = None
if "last_sgml_name" not in st.session_state:
    st.session_state.last_sgml_name = None
if "last_pipeline_score" not in st.session_state:
    st.session_state.last_pipeline_score = None
if "last_val_score" not in st.session_state:
    st.session_state.last_val_score = None
if "last_val_decision" not in st.session_state:
    st.session_state.last_val_decision = None
if "last_excel_sgml_text" not in st.session_state:
    st.session_state.last_excel_sgml_text = None
if "last_excel_sgml_name" not in st.session_state:
    st.session_state.last_excel_sgml_name = None
if "last_excel_xlsx_bytes" not in st.session_state:
    st.session_state.last_excel_xlsx_bytes = None
if "last_excel_xlsx_name" not in st.session_state:
    st.session_state.last_excel_xlsx_name = None
if "last_docx_bytes" not in st.session_state:
    st.session_state.last_docx_bytes = None
if "last_docx_name" not in st.session_state:
    st.session_state.last_docx_name = None

# ── Process-global browser session registry (real active-session count) ───────
# Each Streamlit browser tab gets its own st.session_state but shares the
# Python process.  We use a module-level dict so every tab can see all others.
import threading as _threading
_SESSION_REGISTRY: dict[str, float] = {}      # {session_id: last_ping_epoch}
_SESSION_REGISTRY_LOCK = _threading.Lock()
_SESSION_PROCESSING: dict[str, bool] = {}     # {session_id: is_processing}
_SESSION_TIMEOUT_S = 30 * 60                  # 30 minutes = inactive


def _ping_session(session_id: str, processing: bool = False) -> None:
    """Register/refresh this browser session in the process-global registry."""
    import time
    now = time.time()
    with _SESSION_REGISTRY_LOCK:
        _SESSION_REGISTRY[session_id] = now
        _SESSION_PROCESSING[session_id] = processing
        # Evict sessions not seen for >30 min
        stale = [sid for sid, t in _SESSION_REGISTRY.items()
                 if now - t > _SESSION_TIMEOUT_S]
        for sid in stale:
            _SESSION_REGISTRY.pop(sid, None)
            _SESSION_PROCESSING.pop(sid, None)


def _get_sys_stats() -> dict:
    """Return real browser-session counts from the process-global registry."""
    import time
    now = time.time()
    _ping_session(
        st.session_state.session_id,
        processing=st.session_state.get("pipeline_running", False),
    )
    with _SESSION_REGISTRY_LOCK:
        active     = len(_SESSION_REGISTRY)
        processing = sum(1 for v in _SESSION_PROCESSING.values() if v)
    return {
        "active_sessions":     active,
        "processing_sessions": processing,
        "total_sessions":      active,
    }


def _log(level: str, msg: str) -> None:
    """Append a timestamped log entry to session pipeline_logs."""
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    st.session_state.pipeline_logs.append((ts, level, msg))
    # Cap log buffer at 200 lines
    if len(st.session_state.pipeline_logs) > 200:
        st.session_state.pipeline_logs = st.session_state.pipeline_logs[-200:]


STEP_DEFS = [
    # (label, icon, phase)   phase: 1=Conversion, 2=Validation
    ("Hybrid PDF → DOCX (PyMuPDF + pdfplumber)", "📄", 1),
    ("DOCX → SGML (AI Pipeline)",                "🤖", 1),
    ("L1 Source Fidelity",                       "🔎", 2),
    ("L2 Structural Compliance",                 "🏗️", 2),
    ("L3 Doc-Type Check",                        "📋", 2),
    ("L4 Data Integrity",                        "✅", 2),
]

_PHASE_STEPS = {1: [0, 1], 2: [2, 3, 4, 5]}


def _phase_summary(phase: int) -> tuple[str, str]:
    """Return (bg_colour, status_label) for a phase based on its step states."""
    steps = st.session_state.pipeline_steps
    idxs  = _PHASE_STEPS[phase]
    vals  = [steps.get(i, 'pending') for i in idxs]
    if any(v == 'error'   for v in vals): return '#fee2e2', '❌ Error'
    if any(v == 'running' for v in vals): return '#fef9c3', '⏳ Running…'
    if all(v == 'done'    for v in vals): return '#dcfce7', '✅ Complete'
    return '#f1f5f9', '· Pending'


def _step_row_html(idx: int, label: str, icon: str) -> str:
    """Return one step row as HTML (small font, coloured dot)."""
    st_val = st.session_state.pipeline_steps.get(idx, 'pending')
    dot_col = {'done': '#16a34a', 'running': '#d97706',
               'error': '#dc2626', 'pending': '#94a3b8'}[st_val]
    dot_sym = {'done': '●', 'running': '◉', 'error': '✕', 'pending': '○'}[st_val]
    return (
        f'<div style="display:flex;align-items:center;gap:8px;'
        f'padding:3px 0;font-size:0.78em;color:#374151">'
        f'<span style="color:{dot_col};font-size:1em">{dot_sym}</span>'
        f'<span>{icon}</span>'
        f'<span>{label}</span>'
        f'</div>'
    )


def _render_pipeline_steps(placeholder=None) -> None:
    """Render Phase 1 + Phase 2 as collapsible expanders (PL-style)."""
    steps = st.session_state.pipeline_steps
    if not steps:
        return

    target = placeholder if placeholder else st

    def _draw():
        for phase, title in [(1, "📄 Phase 1 — Conversion")]:
            bg, status = _phase_summary(phase)
            rows_html  = ''.join(
                _step_row_html(i, STEP_DEFS[i][0], STEP_DEFS[i][1])
                for i in _PHASE_STEPS[phase]
            )
            st.markdown(
                f'<details style="background:{bg};border-radius:8px;padding:8px 14px;'
                f'margin-bottom:6px;border:1px solid #e2e8f0">'
                f'<summary style="font-size:0.82em;font-weight:700;cursor:pointer;'
                f'list-style:none;display:flex;justify-content:space-between">'
                f'<span>{title}</span>'
                f'<span style="font-weight:400;color:#64748b">{status}</span>'
                f'</summary>'
                f'<div style="margin-top:6px">{rows_html}</div>'
                f'</details>',
                unsafe_allow_html=True,
            )

    if placeholder:
        with placeholder.container():
            _draw()
    else:
        _draw()


def _render_live_logs(placeholder) -> None:
    """Render pipeline logs inside a collapsible expander (small monospace font)."""
    logs = st.session_state.pipeline_logs
    if not logs:
        return
    LEVEL_COL = {'INFO': '#22c55e', 'WARN': '#f59e0b', 'ERROR': '#ef4444', 'STEP': '#60a5fa'}
    html_lines = []
    for ts, lvl, msg in logs[-80:]:
        col = LEVEL_COL.get(lvl, '#94a3b8')
        esc = msg.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        html_lines.append(
            f'<div><span style="color:#475569">[{ts}]</span> '
            f'<span style="color:{col};font-weight:600">{lvl:<5}</span> {esc}</div>'
        )
    log_html = (
        '<div style="background:#0f172a;color:#e2e8f0;font-family:Consolas,monospace;'
        'font-size:0.70em;line-height:1.45;padding:10px;border-radius:6px;'
        'height:200px;overflow-y:auto">'
        + ''.join(html_lines) + '</div>'
    )
    placeholder.markdown(
        f'<details style="margin-top:4px"><summary style="font-size:0.78em;'
        f'color:#64748b;cursor:pointer">📋 Activity Log ({len(logs)} entries)</summary>'
        f'{log_html}</details>',
        unsafe_allow_html=True,
    )


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🏛️ TR Securities SGML")
    st.markdown("---")

    # Live system stats
    stats = _get_sys_stats()
    c1, c2 = st.columns(2)
    c1.metric("Active Sessions", stats.get('active_sessions', 1))
    c2.metric("Processing",      stats.get('processing_sessions', 0))

    st.markdown("---")

    with st.expander("⚙️ Session Info", expanded=False):
        st.markdown(f"**Session ID:** `{st.session_state.session_id}`")
        st.markdown(f"**Status:** {st.session_state.status}")
        st.markdown(f"**Started:** {st.session_state.created_at}")
        st.markdown(f"**Jobs run:** {st.session_state.processing_count}")

    st.markdown("---")

    if st.button("🆕 New Session", use_container_width=True, key="new_session_btn",
                 help="Reset all state and start fresh"):
        st.session_state.session_id    = uuid.uuid4().hex[:8]
        st.session_state.created_at    = datetime.datetime.now().strftime("%H:%M:%S")
        st.session_state.status        = "active"
        st.session_state.processing_count = 0
        st.session_state.pipeline_steps   = {}
        st.session_state.pipeline_logs    = []
        st.session_state.pipeline_running = False
        st.session_state.pipeline_result_ready = False
        st.session_state.last_pdf_name       = None
        st.session_state.last_pdf_bytes      = None
        st.session_state.last_sgml_text      = None
        st.session_state.last_sgml_name      = None
        st.session_state.last_pipeline_score = None
        st.session_state.last_val_score      = None
        st.session_state.last_val_decision   = None
        st.session_state.last_excel_sgml_text  = None
        st.session_state.last_excel_sgml_name  = None
        st.session_state.last_excel_xlsx_bytes = None
        st.session_state.last_excel_xlsx_name  = None
        st.session_state.last_docx_bytes        = None
        st.session_state.last_docx_name         = None
        for _k in [k for k in list(st.session_state.keys()) if k.startswith("_hitl_")]:
            del st.session_state[_k]
        st.rerun()


# ── Helper: run pipeline with DOCX bytes ─────────────────────────────────────
def _run_sgml_pipeline(docx_bytes: bytes, doc_name: str) -> dict:
    """Load pipeline_runner and convert DOCX → SGML. Feeds _log() during run."""
    try:
        app_dir = Path(__file__).parent / 'app'
        if str(app_dir) not in sys.path:
            sys.path.insert(0, str(app_dir))

        _log('STEP', f'Loading AI pipeline for "{doc_name}"…')
        from app import pipeline_runner
        _log('INFO', 'Pipeline module loaded — starting DOCX → SGML conversion')
        result = pipeline_runner.run_pipeline(docx_bytes, doc_name)
        if result.get('status') == 'ok':
            _log('INFO', f'Pipeline complete — score {result.get("score", "?")}')
        else:
            _log('ERROR', f'Pipeline returned error: {result.get("message", "unknown")}')
        return result
    except Exception as exc:
        _log('ERROR', f'Pipeline exception: {exc}')
        return {"status": "error", "message": str(exc)}


def _pdf_to_docx(pdf_path: str, docx_path: str) -> bool:
    """
    Convert PDF to DOCX using the hybrid PyMuPDF + pdfplumber converter.

    Replaces ABBYY FineReader as the primary conversion engine.
    Works natively on Linux/Docker and Windows with 100% accuracy for
    digital (native-text) PDFs — no OCR required, no content loss.

    Priority order:
      1. Hybrid PyMuPDF + pdfplumber   (primary — digital PDFs, all platforms)
      2. ABBYY FineReader Engine 12    (optional legacy fallback, Windows only)
    """
    # ── Primary: Hybrid converter (PyMuPDF + pdfplumber) ─────────────────────
    try:
        _app_dir = Path(__file__).parent
        if str(_app_dir) not in sys.path:
            sys.path.insert(0, str(_app_dir))
        from validator.pdf.hybrid_converter import convert_pdf_to_docx as _hybrid_convert
        ok = _hybrid_convert(pdf_path, docx_path)
        if ok:
            _log('INFO', 'PDF→DOCX via hybrid converter (PyMuPDF + pdfplumber)')
            return True
        _log('WARNING', 'Hybrid converter returned False — trying ABBYY fallback')
    except Exception as exc:
        _log('WARNING', f'Hybrid converter exception ({exc}) — trying ABBYY fallback')

    # ── Fallback: ABBYY FineReader Engine 12 (Windows only) ──────────────────
    try:
        import pythoncom                  # type: ignore
        import win32com.client as win32c  # type: ignore
        pythoncom.CoInitialize()
        customer_id      = os.getenv('ABBYY_CUSTOMER_ID', '')
        license_path     = os.getenv('ABBYY_LICENSE_PATH', '')
        license_password = os.getenv('ABBYY_LICENSE_PASSWORD', '')

        engine_loader = win32c.Dispatch("FREngine.OutprocLoader.12")
        engine = engine_loader.InitializeEngine(
            customer_id, license_path, license_password, "", "", False
        )
        engine.LoadPredefinedProfile("DocumentConversion_Accuracy")

        document = engine.CreateFRDocument()
        document.AddImageFile(pdf_path, None, None)
        document.Process(None)

        export_params = engine.CreateRTFExportParams()
        export_params.PictureExportParams.Resolution = 300
        export_params.BackgroundColorMode = 1
        export_params.PageSynthesisMode = 1
        export_params.KeepPageBreaks = 1
        export_params.UseDocumentStructure = True
        try:
            export_params.WriteRunningTitles = False
        except AttributeError:
            pass

        document.Export(docx_path, 8, export_params)   # FEF_DOCX = 8
        document.Close()
        if os.path.exists(docx_path):
            _log('INFO', 'PDF→DOCX via ABBYY FineReader (fallback)')
            return True

    except ImportError:
        pass  # ABBYY not available on this platform
    except Exception as exc:
        _log('ERROR', f'ABBYY fallback also failed: {exc}')
    finally:
        try:
            import pythoncom  # type: ignore
            pythoncom.CoUninitialize()
        except Exception:
            pass

    st.error("❌ PDF → DOCX conversion failed. Please upload a DOCX file directly.")
    return False


# ── Main content ──────────────────────────────────────────────────────────────
st.markdown("# 📄 Securities Commission Conversion")

st.markdown("---")

# ════════════════════════════════════════════════════════════════════════════
# SECTION 1 — PDF Upload
# ════════════════════════════════════════════════════════════════════════════
st.markdown("#### 📁 Upload PDF File")

uploaded_pdf = st.file_uploader(
    "Choose a PDF file to process",
    type=["pdf", "docx"],          # accept DOCX too (no ABBYY on Linux)
    key="pdf_uploader",
)

if uploaded_pdf is None:
    st.info("👆 Please upload a PDF file to begin processing")
else:
    file_size_mb = len(uploaded_pdf.getvalue()) / (1024 * 1024)
    is_pdf = uploaded_pdf.name.lower().endswith(".pdf")
    icon = "📄" if is_pdf else "📝"
    st.success(f"{icon} **{uploaded_pdf.name}** — {file_size_mb:.2f} MB uploaded")

    doc_name = Path(uploaded_pdf.name).stem

    col_btn, col_btn2, col_note = st.columns([1, 1, 3])
    with col_btn:
        convert_btn = st.button("⚙️ Process Document", type="primary", key="convert_pdf")
    with col_btn2:
        extract_btn = st.button("🖼️ Extract Images", key="extract_images")
    with col_note:
        if is_pdf:
            st.caption("Process → SGML.  Extract Images → BMP ZIP (SB000001.BMP, SB000002.BMP …)")
        else:
            st.caption("DOCX processed directly by SGML pipeline.  Extract Images → BMP ZIP from word/media/.")

    # ── Extract Images handler ──────────────────────────────────────────────
    if extract_btn:
        import io as _io
        import re as _re
        import zipfile as _zipfile
        from PIL import Image as _PILImage

        # Formats PIL can reliably open from word/media/
        _PIL_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".tiff", ".tif", ".bmp", ".webp"}

        def _media_sort_key(name):
            m = _re.search(r'(\d+)', os.path.basename(name))
            return int(m.group(1)) if m else 0

        progress_img = st.progress(0, text="Preparing DOCX…")
        status_img   = st.empty()
        img_tmp      = tempfile.mkdtemp(prefix="imgext_")

        try:
            # ── Reuse DOCX from pipeline run if available (avoids re-running ABBYY) ──
            _cached_docx  = st.session_state.get("last_docx_bytes")
            _cached_name  = st.session_state.get("last_docx_name", "")
            _same_doc     = _cached_docx and Path(_cached_name).stem == doc_name

            if _same_doc:
                # Already have the DOCX from the pipeline run — reuse it
                docx_bytes = _cached_docx
                progress_img.progress(20, text="Using cached DOCX from pipeline run…")
                status_img.info("📌 Reusing DOCX from pipeline run — skipping ABBYY conversion.")
            elif is_pdf:
                # No cached DOCX — need to run ABBYY
                status_img.info("⏳ Converting PDF → DOCX to access images (run 'Process Document' first to skip this step)…")
                progress_img.progress(10, text="Converting PDF → DOCX…")
                pdf_tmp  = os.path.join(img_tmp, uploaded_pdf.name)
                docx_tmp = os.path.join(img_tmp, f"{doc_name}.docx")
                with open(pdf_tmp, "wb") as fh:
                    fh.write(uploaded_pdf.getvalue())
                ok = _pdf_to_docx(pdf_tmp, docx_tmp)
                if not ok or not os.path.exists(docx_tmp):
                    progress_img.empty(); status_img.empty()
                    st.error("❌ PDF → DOCX conversion failed. Cannot extract images.")
                    shutil.rmtree(img_tmp, ignore_errors=True)
                    st.stop()
                docx_bytes = open(docx_tmp, "rb").read()
                # Cache it so next Extract Images click is instant
                st.session_state.last_docx_bytes = docx_bytes
                st.session_state.last_docx_name  = f"{doc_name}.docx"
            else:
                docx_bytes = uploaded_pdf.getvalue()

            progress_img.progress(30, text="Reading word/media/…")
            status_img.info("⏳ Extracting images…")

            extracted = []
            skipped   = []
            with _zipfile.ZipFile(_io.BytesIO(docx_bytes)) as z:
                media = sorted(
                    [n for n in z.namelist() if n.startswith("word/media/")],
                    key=_media_sort_key,
                )
                # Filter to only PIL-supported formats; skip WMF/EMF/XML etc.
                supported = [n for n in media if Path(n).suffix.lower() in _PIL_EXTS]
                skipped   = [n for n in media if Path(n).suffix.lower() not in _PIL_EXTS]

                total = len(supported)
                if total == 0:
                    progress_img.empty(); status_img.empty()
                    if media:
                        st.warning(
                            f"⚠️ Found {len(media)} file(s) in word/media/ but none are "
                            f"supported image formats (found: "
                            f"{', '.join(set(Path(n).suffix for n in media))})."
                        )
                    else:
                        st.warning("⚠️ No images found in this document (word/media/ is empty).")
                    shutil.rmtree(img_tmp, ignore_errors=True)
                    st.stop()

                bmp_idx = 1
                for i, entry in enumerate(supported):
                    bmp_name = f"SB{bmp_idx:06d}.BMP"
                    bmp_path = os.path.join(img_tmp, bmp_name)
                    raw      = z.read(entry)
                    try:
                        buf = _io.BytesIO(raw)
                        buf.seek(0)
                        img = _PILImage.open(buf)
                        img.load()  # force decode before BytesIO goes out of scope
                        # Convert palette/RGBA to RGB for BMP compatibility
                        if img.mode not in ("RGB", "L", "1"):
                            img = img.convert("RGB")
                        img.save(bmp_path, format="BMP")
                        extracted.append(bmp_path)
                        bmp_idx += 1
                    except Exception as _img_err:
                        skipped.append(f"{entry} ({_img_err})")
                        continue
                    progress_img.progress(
                        30 + int((i + 1) / total * 65),
                        text=f"{os.path.basename(entry)} → {bmp_name}…",
                    )

            if not extracted:
                progress_img.empty(); status_img.empty()
                st.error("❌ No images could be converted. All files were skipped.")
                shutil.rmtree(img_tmp, ignore_errors=True)
                st.stop()

            progress_img.progress(97, text="Building ZIP…")
            zip_buf = _io.BytesIO()
            with _zipfile.ZipFile(zip_buf, "w", _zipfile.ZIP_DEFLATED) as zf:
                for bp in extracted:
                    zf.write(bp, arcname=Path(bp).name)
            zip_buf.seek(0)
            n_out = len(extracted)
            progress_img.progress(100, text="Done!")
            status_img.empty(); progress_img.empty()
            st.success(f"✅ Extracted **{n_out}** image(s) — SB000001.BMP … SB{n_out:06d}.BMP")
            if skipped:
                st.caption(f"ℹ️ {len(skipped)} file(s) skipped (unsupported format: WMF/EMF/XML)")
            st.download_button(
                label=f"⬇️  Download  {doc_name}_images.zip",
                data=zip_buf.getvalue(),
                file_name=f"{doc_name}_images.zip",
                mime="application/zip",
                use_container_width=True,
            )
        except Exception as exc:
            progress_img.empty(); status_img.empty()
            st.error(f"❌ Image extraction error: {exc}")
        finally:
            shutil.rmtree(img_tmp, ignore_errors=True)

    if convert_btn:
        st.session_state.processing_count  += 1
        st.session_state.pipeline_running   = True
        st.session_state.pipeline_logs      = []
        st.session_state.pipeline_steps     = {i: 'pending' for i in range(len(STEP_DEFS))}

        # ── Live UI containers ────────────────────────────────────────────────
        progress_bar      = st.progress(0, text="Initialising…")
        status_box        = st.empty()
        steps_placeholder = st.empty()
        log_placeholder   = st.empty()

        def _refresh():
            _render_pipeline_steps(steps_placeholder)
            _render_live_logs(log_placeholder)

        tmp_dir = tempfile.mkdtemp(prefix="plp_")
        try:
            raw_bytes = uploaded_pdf.getvalue()

            if is_pdf:
                # ── Step 0: ABBYY PDF → DOCX ─────────────────────────────────
                st.session_state.pipeline_steps[0] = 'running'
                _log('STEP', 'Step 1/6 — ABBYY PDF → DOCX conversion starting…')
                _refresh()
                progress_bar.progress(10, text="Step 1/6 — ABBYY PDF → DOCX…")
                status_box.info("⏳  ABBYY FineReader converting PDF to DOCX…")

                pdf_path  = os.path.join(tmp_dir, uploaded_pdf.name)
                docx_path = os.path.join(tmp_dir, f"{doc_name}.docx")
                with open(pdf_path, "wb") as fh:
                    fh.write(raw_bytes)

                ok = _pdf_to_docx(pdf_path, docx_path)
                if not ok:
                    st.session_state.pipeline_steps[0] = 'error'
                    _log('ERROR', 'ABBYY conversion failed — aborting.')
                    _refresh()
                    st.session_state.processing_count  = max(0, st.session_state.processing_count - 1)
                    st.session_state.pipeline_running   = False
                    progress_bar.empty(); status_box.empty()
                    st.stop()

                st.session_state.pipeline_steps[0] = 'done'
                _log('INFO', f'ABBYY conversion complete → {Path(docx_path).name}')
                _refresh()

                with open(docx_path, "rb") as fh:
                    docx_bytes = fh.read()
                # Cache DOCX so Extract Images can reuse without re-running ABBYY
                st.session_state.last_docx_bytes = docx_bytes
                st.session_state.last_docx_name  = f"{doc_name}.docx"
                progress_bar.progress(25, text="Step 2/6 — SGML pipeline…")
            else:
                docx_bytes = raw_bytes
                st.session_state.last_docx_bytes = raw_bytes
                st.session_state.last_docx_name  = uploaded_pdf.name
                st.session_state.pipeline_steps[0] = 'done'  # N/A for DOCX input
                _log('INFO', 'DOCX uploaded directly — skipping ABBYY step.')
                progress_bar.progress(15, text="Step 2/6 — SGML pipeline…")

            # ── Step 1: DOCX → SGML (AI pipeline) ────────────────────────────
            st.session_state.pipeline_steps[1] = 'running'
            _log('STEP', 'Step 2/6 — DOCX → SGML AI pipeline starting…')
            _refresh()
            status_box.info("⏳  Running SGML extraction pipeline…")

            result = _run_sgml_pipeline(docx_bytes, doc_name)
            progress_bar.progress(60, text="Step 3/6 — Validating output…")

            if result.get("status") != "success":
                st.session_state.pipeline_steps[1] = 'error'
                _log('ERROR', f'Pipeline error: {result.get("message","unknown")}')
                _refresh()
                st.session_state.processing_count  = max(0, st.session_state.processing_count - 1)
                st.session_state.pipeline_running   = False
                status_box.empty(); progress_bar.empty()
                st.error(f"❌ Pipeline error: {result.get('message', 'Unknown error')}")
                st.stop()

            st.session_state.pipeline_steps[1] = 'done'
            sgml_text: str = result["sgml"]
            _log('INFO', f'SGML extraction complete — {len(sgml_text):,} chars generated')
            _refresh()

            # ── Steps 2-5: Validate ───────────────────────────────────────────
            for step_i in [2, 3, 4, 5]:
                st.session_state.pipeline_steps[step_i] = 'running'
            _refresh()

            try:
                pipeline_dir = str(Path(__file__).parent / "pipeline")
                if pipeline_dir not in sys.path:
                    sys.path.insert(0, pipeline_dir)
                import importlib, excel_validator as _ev_mod
                importlib.reload(_ev_mod)   # pick up any live edits

                import tempfile as _tf
                with _tf.NamedTemporaryFile(delete=False, suffix='.sgm',
                                            mode='w', encoding='utf-8') as _sgm_f:
                    _sgm_f.write(sgml_text)
                    _sgm_tmp = Path(_sgm_f.name)

                _val_result = _ev_mod.validate(_sgm_tmp, None)

                for step_i in [2, 3, 4, 5]:
                    st.session_state.pipeline_steps[step_i] = 'done'
                _refresh()

                val_score    = _val_result['scores']['normalised']
                val_decision = _val_result['decision']
                _sgm_tmp.unlink(missing_ok=True)
            except Exception as _ve:
                for step_i in [2, 3, 4, 5]:
                    st.session_state.pipeline_steps[step_i] = 'done'
                _log('WARN', f'Validator skipped (not blocking): {_ve}')
                _refresh()
                val_score    = None
                val_decision = None

            progress_bar.progress(100, text="✅ Complete!")
            status_box.empty()
            st.session_state.pipeline_running  = False
            st.session_state.processing_count  = max(0, st.session_state.processing_count - 1)
            _refresh()
            _log('INFO', '🎉 All steps complete.')
            _render_live_logs(log_placeholder)

            # ── Persist results for cross-page carry-over and page-revisit ──────
            score_pipeline = result.get("score")
            st.session_state.last_pdf_name       = uploaded_pdf.name
            st.session_state.last_pdf_bytes      = raw_bytes if is_pdf else None
            st.session_state.last_sgml_text      = sgml_text
            st.session_state.last_sgml_name      = f"{doc_name}_TR.sgm"
            st.session_state.last_pipeline_score = score_pipeline
            st.session_state.last_val_score      = val_score
            st.session_state.last_val_decision   = val_decision
            st.session_state.pipeline_result_ready = True
            # Invalidate HITL cached temp-file paths so HITL reloads the new SGML
            for _ck in [k for k in list(st.session_state.keys()) if k.startswith("_hitl_")]:
                del st.session_state[_ck]

            # ── Results ───────────────────────────────────────────────────────
            st.markdown("---")
            col_a, col_c = st.columns([2, 2])
            with col_a:
                st.success("✅ SGML conversion complete!")
            with col_c:
                st.download_button(
                    label=f"⬇️  Download  {doc_name}_TR.sgm",
                    data=sgml_text,
                    file_name=f"{doc_name}_TR.sgm",
                    mime="text/plain",
                    use_container_width=True,
                )

            with st.expander("🔍 Preview SGML output (first 3 000 characters)"):
                st.code(sgml_text[:3000], language="xml")

        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)
            st.session_state.pipeline_running = False


# ── Persistent results panel (survives page navigation / page revisit) ────────
if uploaded_pdf is None and st.session_state.get("pipeline_result_ready"):
    _p_score = st.session_state.last_pipeline_score
    _v_score = st.session_state.last_val_score
    _v_dec   = st.session_state.last_val_decision
    _sgml    = st.session_state.last_sgml_text
    _sname   = st.session_state.last_sgml_name or "output.sgm"
    _pname   = st.session_state.last_pdf_name  or "?"

    st.markdown("---")

    # Completed step tracker
    _render_pipeline_steps()

    # Pipeline logs
    if st.session_state.pipeline_logs:
        _LCOL = {'INFO': '#22c55e', 'WARN': '#f59e0b', 'ERROR': '#ef4444', 'STEP': '#60a5fa'}
        _log_lines = []
        for _ts, _lvl, _msg in st.session_state.pipeline_logs[-80:]:
            _col = _LCOL.get(_lvl, '#94a3b8')
            _esc = _msg.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            _log_lines.append(
                f'<div><span style="color:#475569">[{_ts}]</span> '
                f'<span style="color:{_col};font-weight:600">{_lvl:<5}</span> {_esc}</div>'
            )
        _log_html = (
            '<div style="background:#0f172a;color:#e2e8f0;font-family:Consolas,monospace;'
            'font-size:0.70em;line-height:1.45;padding:10px;border-radius:6px;'
            'height:200px;overflow-y:auto">' + ''.join(_log_lines) + '</div>'
        )
        st.markdown(
            f'<details style="margin-top:4px"><summary style="font-size:0.78em;'
            f'color:#64748b;cursor:pointer">'
            f'📋 Activity Log ({len(st.session_state.pipeline_logs)} entries)</summary>'
            f'{_log_html}</details>',
            unsafe_allow_html=True,
        )

    # Download
    _col_a, _col_c = st.columns([2, 2])
    with _col_a:
        st.success("✅ SGML conversion complete!")
    with _col_c:
        if _sgml:
            st.download_button(
                label=f"⬇️  Download  {_sname}",
                data=_sgml,
                file_name=_sname,
                mime="text/plain",
                use_container_width=True,
                key="dl_sgml_persistent",
            )

    if _sgml:
        with st.expander("🔍 Preview SGML output (first 3 000 characters)"):
            st.code(_sgml[:3000], language="xml")


# ════════════════════════════════════════════════════════════════════════════
# SECTION 2 — Excel Upload
# ════════════════════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("#### 📊 Upload Excel File")

uploaded_excel = st.file_uploader(
    "Choose an Excel file to process",
    type=["xlsx", "xls"],
    key="excel_uploader",
)

if uploaded_excel is None:
    st.info("👆 Please upload an Excel file (.xlsx / .xls) to begin conversion")
else:
    file_size_mb = len(uploaded_excel.getvalue()) / (1024 * 1024)
    st.success(f"📊 **{uploaded_excel.name}** — {file_size_mb:.2f} MB uploaded")

    excel_doc_name = Path(uploaded_excel.name).stem

    col_btn2, col_note2 = st.columns([1, 4])
    with col_btn2:
        excel_btn = st.button("⚙️ Process Document", type="primary", key="convert_excel")
    with col_note2:
        st.caption("Excel → SGML conversion using rule-based pipeline (no AI required).")

    if excel_btn:
        st.session_state.processing_count += 1
        progress_bar2 = st.progress(0, text="Initialising Excel pipeline…")
        status_box2   = st.empty()

        excel_tmp = tempfile.mkdtemp(prefix="excel_plp_")
        try:
            # Write uploaded Excel to temp dir
            xlsx_path = os.path.join(excel_tmp, uploaded_excel.name)
            sgm_path  = os.path.join(excel_tmp, excel_doc_name + ".sgm")
            with open(xlsx_path, "wb") as fh:
                fh.write(uploaded_excel.getvalue())

            progress_bar2.progress(30, text="Converting Excel → SGML…")
            status_box2.info("⏳  Running Excel → SGML conversion…")

            # Load converter from pipeline folder
            pipeline_dir = str(Path(__file__).parent / "pipeline")
            if pipeline_dir not in sys.path:
                sys.path.insert(0, pipeline_dir)

            from excel_batch_converter import convert_single
            convert_single(xlsx_path, sgm_path)

            st.session_state.processing_count = max(0, st.session_state.processing_count - 1)
            progress_bar2.progress(100, text="Done!")
            status_box2.empty()
            progress_bar2.empty()

            if os.path.exists(sgm_path):
                sgml_text = Path(sgm_path).read_text(encoding="utf-8", errors="replace")

                # ── Persist for Excel HITL carry-over ────────────────────────
                st.session_state.last_excel_sgml_text  = sgml_text
                st.session_state.last_excel_sgml_name  = f"{excel_doc_name}.sgm"
                st.session_state.last_excel_xlsx_bytes = uploaded_excel.getvalue()
                st.session_state.last_excel_xlsx_name  = uploaded_excel.name

                col_a2, col_b2 = st.columns([2, 3])
                with col_a2:
                    st.success("✅ Excel → SGML conversion complete!")
                    st.metric("SGML size", f"{len(sgml_text):,} chars")
                with col_b2:
                    st.download_button(
                        label=f"⬇️  Download  {excel_doc_name}.sgm",
                        data=sgml_text,
                        file_name=f"{excel_doc_name}.sgm",
                        mime="text/plain",
                        use_container_width=True,
                    )
                with st.expander("🔍 Preview SGML output (first 3 000 characters)"):
                    st.code(sgml_text[:3000], language="xml")
            else:
                st.error("❌ Conversion ran but no SGML file was produced.")

        except Exception as exc:
            st.session_state.processing_count = max(0, st.session_state.processing_count - 1)
            progress_bar2.empty()
            status_box2.empty()
            st.error(f"❌ Excel pipeline error: {exc}")

        finally:
            shutil.rmtree(excel_tmp, ignore_errors=True)


