"""
Pipeline runner — thin shim that loads batch_runner_deploy.py, patches
Windows-only imports, overrides hardcoded PATHS, then exposes a clean
`run_pipeline(docx_bytes, doc_name) -> dict` function used by the API.

Design notes
------------
* `win32com` / ABBYY is Windows-only and is NOT available in the Docker
  (Linux) container.  We inject a comprehensive no-op stub covering all
  sub-modules so the pipeline can import cleanly and falls back gracefully
  when `self.abbyy is None`.
* All hardcoded Windows paths in PATHS / RAG_CONFIG are overridden from the
  environment-driven config before the pipeline is executed.
* The pipeline script is loaded via `exec()` into an isolated namespace to
  avoid polluting the global module space.
* Loading is fully lazy — the pipeline script is NOT imported at startup,
  only on the first actual convert request. This keeps /health fast.
"""
import os
import sys
import types
import shutil
import tempfile
from pathlib import Path
from typing import Any, Dict


# ── Comprehensive win32com / pywin32 stub (Linux-safe) ──────────────────────
def _install_win32com_stub() -> None:
    """
    Inject no-op stubs for every win32/pywin32 module the pipeline touches.
    Must be called BEFORE any exec() of the pipeline script.
    """
    class _NullObj:
        """A do-nothing object that absorbs any attribute access or call."""
        def __init__(self, *a, **kw): pass
        def __getattr__(self, _): return self
        def __call__(self, *a, **kw): return self
        def __iter__(self): return iter([])
        def __bool__(self): return False

    stubs = [
        "win32com",
        "win32com.client",
        "win32com.server",
        "win32con",
        "win32api",
        "win32gui",
        "pywintypes",
        "pythoncom",
        "winerror",
    ]
    for name in stubs:
        if name not in sys.modules:
            mod = types.ModuleType(name)
            # Attach the null class as common entry points
            mod.Dispatch = _NullObj       # type: ignore[attr-defined]
            mod.CoInitialize = _NullObj   # type: ignore[attr-defined]
            mod.CoUninitialize = _NullObj # type: ignore[attr-defined]
            sys.modules[name] = mod

    # Make win32com.client accessible as attribute
    sys.modules["win32com"].client = sys.modules["win32com.client"]  # type: ignore[attr-defined]


_install_win32com_stub()

# ── Load the monolithic pipeline script ──────────────────────────────────────
_PIPELINE_SCRIPT = Path(__file__).parent.parent / "pipeline" / "batch_runner_deploy.py"

_pipeline_ns: Dict[str, Any] = {}
_initialized = False


def _load_pipeline() -> None:
    global _initialized
    if _initialized:
        return
    if not _pipeline_script_exists():
        raise RuntimeError(
            f"Pipeline script not found: {_PIPELINE_SCRIPT}\n"
            "Ensure batch_runner_deploy.py is present in the pipeline/ directory."
        )
    src = _PIPELINE_SCRIPT.read_text(encoding="utf-8-sig", errors="replace")
    code = compile(src, str(_PIPELINE_SCRIPT), "exec")
    exec(code, _pipeline_ns)  # noqa: S102
    _initialized = True


def _pipeline_script_exists() -> bool:
    return _PIPELINE_SCRIPT.exists()


# ── Public API ────────────────────────────────────────────────────────────────
def is_pipeline_available() -> bool:
    return _pipeline_script_exists()


def run_pipeline(docx_bytes: bytes, doc_name: str) -> Dict[str, Any]:
    """
    Convert a DOCX file (supplied as raw bytes) to SGML.

    Parameters
    ----------
    docx_bytes : bytes
        Raw bytes of the `.docx` file to convert.
    doc_name : str
        The document identifier / stem (e.g. ``"51-737"``).

    Returns
    -------
    dict with keys:
        status  – "success" | "error"
        sgml    – SGML string (on success)
        message – error detail (on failure)
        score   – pipeline confidence score (float, 0–100)
    """
    _load_pipeline()

    from app import config  # Import here to avoid circular imports

    # Create an isolated temp directory for this request
    os.makedirs(config.TEMP_DIR, exist_ok=True)  # ensure base dir exists (not created by Docker)
    tmp_dir = tempfile.mkdtemp(prefix=f"sgml_{doc_name}_", dir=config.TEMP_DIR)
    os.makedirs(tmp_dir, exist_ok=True)

    try:
        # Write the DOCX to the temp directory
        docx_path = os.path.join(tmp_dir, f"{doc_name}.docx")
        with open(docx_path, "wb") as fh:
            fh.write(docx_bytes)

        sgml_path = os.path.join(tmp_dir, f"{doc_name}.sgm")

        # Override hardcoded PATHS in the pipeline namespace
        _pipeline_ns["PATHS"]["input_pdf"] = os.path.join(tmp_dir, f"{doc_name}.pdf")
        _pipeline_ns["PATHS"]["output_dir"] = tmp_dir
        _pipeline_ns["PATHS"]["keying_rules"] = config.KEYING_RULES_PATH

        # Override RAG config
        _pipeline_ns["RAG_CONFIG"]["enabled"] = config.RAG_ENABLED
        _pipeline_ns["RAG_CONFIG"]["persist_dir"] = config.RAG_PERSIST_DIR

        # Override SYSTEM_CONFIG
        _pipeline_ns["SYSTEM_CONFIG"]["use_llm"] = config.USE_LLM
        _pipeline_ns["SYSTEM_CONFIG"]["extract_images"] = config.EXTRACT_IMAGES
        _pipeline_ns["SYSTEM_CONFIG"]["max_tokens"] = config.MAX_TOKENS

        # Instantiate and run
        PipelineClass = _pipeline_ns["CompletePipeline"]
        pipe = PipelineClass()
        pipe.initialize()

        # Pass a fake pdf_path — the pipeline will fall back to the DOCX we placed
        fake_pdf = os.path.join(tmp_dir, f"{doc_name}.pdf")
        result = pipe.convert(fake_pdf)

        if result.get("status") == "error":
            return {"status": "error", "message": result.get("message", "Pipeline error")}

        # Read output SGML
        if not os.path.exists(sgml_path):
            return {"status": "error", "message": "Pipeline ran but no SGML output found"}

        sgml_content = Path(sgml_path).read_text(encoding="utf-8", errors="replace")
        return {
            "status": "success",
            "sgml": sgml_content,
            "score": result.get("score"),
        }

    except Exception as exc:
        return {"status": "error", "message": str(exc)}

    finally:
        # Always clean up temp directory
        shutil.rmtree(tmp_dir, ignore_errors=True)
