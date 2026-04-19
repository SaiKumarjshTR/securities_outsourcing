"""
Securities Commission Conversion — Streamlit UI
=============================================
AI-powered PDF/DOCX → SGML conversion for TR Securities Outsourcing.

Pipeline: PDF → (pdf2docx) → DOCX → batch_runner_standalone.py → _TR.sgm

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
if "active_sessions" not in st.session_state:
    st.session_state.active_sessions = 1
if "processing_count" not in st.session_state:
    st.session_state.processing_count = 0


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    with st.expander("⚙ Session Info", expanded=False):
        st.markdown(f"**Session ID:** `{st.session_state.session_id}`")
        st.markdown(f"**Status:** {st.session_state.status}")
        st.markdown(f"**Created:** {st.session_state.created_at}")

    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

    with st.expander("📊 System Status", expanded=False):
        st.metric("Active Sessions", st.session_state.active_sessions)
        st.metric("Processing", st.session_state.processing_count)

    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

    with st.expander("🆕 New Session", expanded=False):
        st.markdown("Start a fresh session and reset all state.")
        if st.button("Start New Session", use_container_width=True, key="new_session_btn"):
            st.session_state.session_id = uuid.uuid4().hex[:8]
            st.session_state.created_at = datetime.datetime.now().strftime("%H:%M:%S")
            st.session_state.status = "active"
            st.session_state.processing_count = 0
            st.rerun()


# ── Helper: run pipeline with DOCX bytes ─────────────────────────────────────
def _run_sgml_pipeline(docx_bytes: bytes, doc_name: str) -> dict:
    """Load pipeline_runner and convert DOCX → SGML."""
    try:
        # Ensure app/ is importable
        app_dir = Path(__file__).parent
        if str(app_dir) not in sys.path:
            sys.path.insert(0, str(app_dir))

        from app import pipeline_runner
        return pipeline_runner.run_pipeline(docx_bytes, doc_name)
    except Exception as exc:
        return {"status": "error", "message": str(exc)}


def _pdf_to_docx(pdf_path: str, docx_path: str) -> bool:
    """Convert PDF to DOCX via pdf2docx. Returns True on success."""
    try:
        from pdf2docx import Converter as PDFConverter  # type: ignore
        cv = PDFConverter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        return True
    except ImportError:
        st.warning(
            "pdf2docx is not installed. "
            "Please upload a DOCX file instead, or install pdf2docx."
        )
        return False
    except Exception as exc:
        st.error(f"PDF → DOCX conversion failed: {exc}")
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

        def _media_sort_key(name):
            m = _re.search(r'(\d+)', os.path.basename(name))
            return int(m.group(1)) if m else 0

        progress_img = st.progress(0, text="Reading DOCX…")
        status_img   = st.empty()
        img_tmp      = tempfile.mkdtemp(prefix="imgext_")

        try:
            # For PDF: convert to DOCX first so we can read word/media/
            if is_pdf:
                status_img.info("⏳ Converting PDF → DOCX to access images…")
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
            else:
                docx_bytes = uploaded_pdf.getvalue()

            progress_img.progress(30, text="Reading word/media/…")
            status_img.info("⏳ Extracting images…")

            extracted = []
            with _zipfile.ZipFile(_io.BytesIO(docx_bytes)) as z:
                media = sorted(
                    [n for n in z.namelist() if n.startswith("word/media/")],
                    key=_media_sort_key,
                )
                total = len(media)
                if total == 0:
                    progress_img.empty(); status_img.empty()
                    st.warning("⚠️ No images found in this document (word/media/ is empty).")
                    shutil.rmtree(img_tmp, ignore_errors=True)
                    st.stop()

                for i, entry in enumerate(media):
                    bmp_name = f"SB{i + 1:06d}.BMP"   # SB000001, SB000002, …
                    bmp_path = os.path.join(img_tmp, bmp_name)
                    raw      = z.read(entry)
                    img      = _PILImage.open(_io.BytesIO(raw))
                    img.save(bmp_path, format="BMP")   # preserve original mode (RGB / 1-bit)
                    extracted.append(bmp_path)
                    progress_img.progress(
                        30 + int((i + 1) / total * 65),
                        text=f"{os.path.basename(entry)} → {bmp_name}…",
                    )

            progress_img.progress(97, text="Building ZIP…")
            zip_buf = _io.BytesIO()
            with _zipfile.ZipFile(zip_buf, "w", _zipfile.ZIP_DEFLATED) as zf:
                for bp in extracted:
                    zf.write(bp, arcname=Path(bp).name)
            zip_buf.seek(0)
            progress_img.progress(100, text="Done!")
            status_img.empty(); progress_img.empty()
            st.success(f"✅ Extracted **{total}** image(s) — SB000001.BMP … SB{total:06d}.BMP")
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
        st.session_state.processing_count += 1
        progress_bar = st.progress(0, text="Initialising pipeline…")
        status_box = st.empty()

        tmp_dir = tempfile.mkdtemp(prefix="plp_")
        try:
            raw_bytes = uploaded_pdf.getvalue()

            if is_pdf:
                # Step 1: PDF → DOCX
                progress_bar.progress(20, text="Converting PDF → DOCX…")
                status_box.info("⏳  Converting PDF to DOCX (pdf2docx)…")
                pdf_path = os.path.join(tmp_dir, uploaded_pdf.name)
                docx_path = os.path.join(tmp_dir, f"{doc_name}.docx")
                with open(pdf_path, "wb") as fh:
                    fh.write(raw_bytes)
                ok = _pdf_to_docx(pdf_path, docx_path)
                if not ok:
                    st.session_state.processing_count = max(0, st.session_state.processing_count - 1)
                    progress_bar.empty()
                    status_box.empty()
                    st.stop()
                with open(docx_path, "rb") as fh:
                    docx_bytes = fh.read()
                progress_bar.progress(40, text="PDF converted — running SGML pipeline…")
            else:
                docx_bytes = raw_bytes
                progress_bar.progress(30, text="Running SGML pipeline…")

            # Step 2: DOCX → SGML
            status_box.info("⏳  Running SGML extraction pipeline…")
            result = _run_sgml_pipeline(docx_bytes, doc_name)

            st.session_state.processing_count = max(0, st.session_state.processing_count - 1)
            progress_bar.progress(100, text="Done!")
            status_box.empty()
            progress_bar.empty()

            if result.get("status") == "success":
                sgml_text: str = result["sgml"]
                score = result.get("score")

                col_a, col_b = st.columns([2, 3])
                with col_a:
                    st.success("✅ SGML conversion complete!")
                    if score is not None:
                        st.metric("Quality Score", f"{score:.1f}%")
                with col_b:
                    st.download_button(
                        label=f"⬇️  Download  {doc_name}_TR.sgm",
                        data=sgml_text,
                        file_name=f"{doc_name}_TR.sgm",
                        mime="text/plain",
                        use_container_width=True,
                    )

                with st.expander("🔍 Preview SGML output (first 3 000 characters)"):
                    st.code(sgml_text[:3000], language="xml")
            else:
                st.error(f"❌ Pipeline error: {result.get('message', 'Unknown error')}")

        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)


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


