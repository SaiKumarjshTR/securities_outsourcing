"""
sync_to_deploy.py
=================
Transforms final_scripts/batch_runner_standalone.py
into the production-safe pipeline/batch_runner_deploy.py.

Transformations applied:
  1. Strip notebook-only cells (CELL 16+) — contain hardcoded Windows test paths
  2. Replace CELL 0 imports block with deploy-safe imports (win32com guarded)
  3. Replace CELL 1 config block with env-var driven config (no hardcoded paths)
  4. Remove all remaining # ====== CODE CELL N ====== markers
  5. Also copies standalone reference into deployment/pipeline/

Usage: python sync_to_deploy.py
"""

import os
import re
import shutil
from pathlib import Path
from datetime import datetime


# ---------------------------------------------------------------------------
# File paths
# ---------------------------------------------------------------------------
STANDALONE_SRC = (
    r"C:\Users\C303180\OneDrive - Thomson Reuters Incorporated\Desktop\TR"
    r"\securities-outsourcing-samples\final_scripts\batch_runner_standalone.py"
)
DEPLOY_OUT = (
    r"C:\Users\C303180\OneDrive - Thomson Reuters Incorporated\Desktop\TR"
    r"\sgml-pipeline-deployment\pipeline\batch_runner_deploy.py"
)
STANDALONE_OUT = (
    r"C:\Users\C303180\OneDrive - Thomson Reuters Incorporated\Desktop\TR"
    r"\sgml-pipeline-deployment\pipeline\batch_runner_standalone.py"
)


# ---------------------------------------------------------------------------
# Production-safe replacement blocks
# ---------------------------------------------------------------------------

DEPLOY_HEADER_IMPORTS = """\
# batch_runner_deploy.py - Linux/Docker safe deployment build.
# Generated from batch_runner_standalone.py by sync_to_deploy.py.
# All Windows-specific paths replaced with env-var driven config.
# win32com (ABBYY) guarded with try/except - returns None on Linux.
# Notebook-only cells (CELL 16+) stripped - not needed in production.
#
import os
import re
import json
import time
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
from datetime import datetime
from dataclasses import dataclass, field
from collections import defaultdict

# ABBYY - Windows-only (ABBYY FineReader Engine 12).
# On Linux/Docker this is stubbed via pipeline_runner._install_win32com_stub().
try:
    import win32com.client as win32c
    _WIN32_AVAILABLE = True
except ImportError:
    win32c = None  # type: ignore[assignment]
    _WIN32_AVAILABLE = False

# DOCX
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# Image processing
from PIL import Image

# Anthropic
from anthropic import Anthropic
import requests

# RAG (ChromaDB vector database)
import chromadb

print("\\u2705 All imports successful!")
print(f"\\U0001f4c5 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
"""


DEPLOY_CONFIG_BLOCK = """\
# ---------------------------------------------------------------------------
# DEPLOYMENT CONFIGURATION - all values from environment variables.
# pipeline_runner.py overrides PATHS/RAG_CONFIG after exec() — these defaults
# are used only when running the script directly (e.g. local testing).
# ---------------------------------------------------------------------------

# ABBYY - disabled in container (no Windows COM server on Linux).
ABBYY_CONFIG = {
    'customer_id':      os.getenv('ABBYY_CUSTOMER_ID', 'FFvrEyp5Gz8sXSwP98N9'),
    'license_path':     os.getenv('ABBYY_LICENSE_PATH', ''),    # Windows-only, N/A in Docker
    'license_password': os.getenv('ABBYY_LICENSE_PASSWORD', '/80HjebrjO2bzpJUiJ/DwQ=='),
}

# Thomson Reuters AI Platform
WORKSPACE_ID    = os.getenv('WORKSPACE_ID', 'Saikumar3Y0Z')
TR_AUTH_URL     = os.getenv('TR_AUTH_URL', 'https://aiplatform.gcs.int.thomsonreuters.com/v1/anthropic/token')
ANTHROPIC_MODEL = os.getenv('ANTHROPIC_MODEL', 'claude-sonnet-4-5-20250929')
OPUS_MODEL      = os.getenv('OPUS_MODEL', 'claude-opus-4-20250514')

# Runtime paths - overridden per-request by pipeline_runner.run_pipeline()
_DEFAULT_TEMP = os.getenv('TEMP_DIR', '/tmp/sgml_pipeline')
os.makedirs(_DEFAULT_TEMP, exist_ok=True)   # ensure base temp dir exists

PATHS = {
    'input_pdf':    os.path.join(_DEFAULT_TEMP, 'input.pdf'),
    'output_dir':   _DEFAULT_TEMP,
    'keying_rules': os.getenv('KEYING_RULES_PATH', '/app/data/COMPLETE_KEYING_RULES_UPDATED.txt'),
}

# System config - can be overridden via env vars
SYSTEM_CONFIG = {
    'use_llm':                 os.getenv('USE_LLM', 'true').lower() == 'true',
    'llm_batch_size':          int(os.getenv('LLM_BATCH_SIZE', '10')),
    'extract_tables':          True,
    'extract_images':          os.getenv('EXTRACT_IMAGES', 'false').lower() == 'true',
    'apply_inline_formatting': True,
    'max_tokens':              int(os.getenv('MAX_TOKENS', '16000')),
    'temperature':             0.0,
    'image_dpi':               300,
}

# RAG Configuration - vendor SGMs auto-discovered from VENDOR_SGMS_DIR
_VENDOR_SGMS_DIR = os.getenv('VENDOR_SGMS_DIR', '/app/data/vendor_sgms')

def _discover_vendor_sgms(vendor_dir: str) -> list:
    # Auto-discover all .sgm files in the vendor SGMs directory.
    p = Path(vendor_dir)
    if not p.exists():
        return []
    return sorted(str(f) for f in p.rglob('*.sgm'))

RAG_CONFIG = {
    'enabled':     os.getenv('RAG_ENABLED', 'true').lower() == 'true',
    'persist_dir': os.getenv('RAG_PERSIST_DIR', '/app/data/chroma_db'),
    'n_rules':     int(os.getenv('RAG_N_RULES', '12')),
    'n_examples':  int(os.getenv('RAG_N_EXAMPLES', '25')),
    'vendor_sgms': _discover_vendor_sgms(_VENDOR_SGMS_DIR),
}

AGENT_CONFIG = {
    'orchestrator':       {'model': OPUS_MODEL, 'temperature': 0.0, 'max_tokens': 4096},
    'specialized_agents': {'model': OPUS_MODEL, 'temperature': 0.0, 'max_tokens': 4096},
    'validator':          {'model': OPUS_MODEL, 'temperature': 0.0, 'max_tokens': 2048},
    'enable_parallel':    True,
    'agent_routing_rules': {
        'has_bold_formatting':    ['BOLD_AGENT'],
        'has_italic_formatting':  ['EM_AGENT'],
        'has_bullet_or_indent':   ['ITEM_AGENT'],
        'is_heading_style':       ['BLOCK_AGENT'],
        'contains_email':         ['EM_AGENT'],
        'contains_act_reference': ['EM_AGENT'],
        'is_regular_paragraph':   ['STRUCTURE_AGENT'],
    }
}

print("\\u2705 Configuration loaded!")
print(f"\\U0001f4c5 Output: {PATHS['output_dir']}")
print(f"\\U0001f916 Model (Sonnet fallback): {ANTHROPIC_MODEL}")
print(f"\\U0001f9e0 RAG: {'ENABLED' if RAG_CONFIG['enabled'] else 'DISABLED'}")
print(f"\\U0001f916 Agentic Model (Opus primary): {OPUS_MODEL}")
"""


# ---------------------------------------------------------------------------
# Build function
# ---------------------------------------------------------------------------

def build_deploy(text: str) -> str:
    lines = text.split('\n')

    # Step 1: Truncate at CELL 16 (notebook-only test harness, hardcoded paths)
    for i, line in enumerate(lines):
        if re.match(r'^# ={6} CODE CELL 16 ={6}', line):
            lines = lines[:i]
            print(f"  Truncated at line {i+1} (CODE CELL 16)")
            break
    text = '\n'.join(lines)

    # Step 2 + 3: Replace CELL 0 AND CELL 1 together in one shot.
    # Find from start of "# ====== CODE CELL 0 ======" marker
    # to just before "# ====== CODE CELL 2 ======" marker.
    # This covers: imports (win32com), ABBYY_CONFIG, PATHS, RAG_CONFIG, AGENT_CONFIG.
    pat_cells01 = re.compile(
        r'# ={6} CODE CELL 0 ={6}\n'
        r'.*?'
        r'(?=# ={6} CODE CELL 2 ={6})',
        re.DOTALL,
    )
    m = pat_cells01.search(text)
    if not m:
        raise ValueError(
            "CELL 0 / CELL 2 boundary not found - check standalone file format"
        )
    replacement = DEPLOY_HEADER_IMPORTS + '\n' + DEPLOY_CONFIG_BLOCK + '\n'
    text = text[:m.start()] + replacement + text[m.end():]

    # Step 4: Remove all remaining CODE CELL markers
    text = re.sub(r'^# ={6} CODE CELL \d+ ={6}\n', '', text, flags=re.MULTILINE)

    return text.rstrip('\n') + '\n'


# ---------------------------------------------------------------------------
# Sanity checks
# ---------------------------------------------------------------------------

def check(deploy: str) -> bool:
    tests = [
        ("No hardcoded C:\\Users\\C303180 paths",
         r"C:\Users\C303180" not in deploy),
        ("win32com guarded with try/except",
         "try:\n    import win32com" in deploy),
        ("ABBYY_CUSTOMER_ID via os.getenv",
         "os.getenv('ABBYY_CUSTOMER_ID'" in deploy),
        ("WORKSPACE_ID via os.getenv",
         "os.getenv('WORKSPACE_ID'" in deploy),
        ("KEYING_RULES_PATH via os.getenv",
         "os.getenv('KEYING_RULES_PATH'" in deploy),
        ("_post_fix_vendor_footnote_injection present",
         "_post_fix_vendor_footnote_injection" in deploy),
        ("_post_fix_replace_deficient_table present",
         "_post_fix_replace_deficient_table" in deploy),
        ("No CODE CELL markers remaining",
         "# ====== CODE CELL " not in deploy),
        ("class CompletePipeline present",
         "class CompletePipeline" in deploy),
        ("Notebook TEST_BASE stripped",
         "TEST_BASE" not in deploy),
        ("vendor_sgms uses _discover_vendor_sgms",
         "_discover_vendor_sgms" in deploy),
    ]
    ok = True
    for desc, result in tests:
        status = "OK  " if result else "FAIL"
        if not result:
            ok = False
        print(f"  [{status}] {desc}")
    return ok


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("=" * 60)
    print("sync_to_deploy.py")
    print("=" * 60)

    print(f"\nReading: {STANDALONE_SRC}")
    src = open(STANDALONE_SRC, encoding='utf-8', errors='ignore').read()
    print(f"  {len(src):,} chars  |  {src.count(chr(10))+1:,} lines")

    print("\nBuilding deploy version...")
    deploy = build_deploy(src)
    print(f"  {len(deploy):,} chars  |  {deploy.count(chr(10))+1:,} lines")

    print("\nRunning sanity checks...")
    ok = check(deploy)

    if not ok:
        print("\n[ABORT] Checks failed — output NOT written. Fix sync_to_deploy.py and retry.")
        return 1

    print(f"\nWriting deploy: {DEPLOY_OUT}")
    with open(DEPLOY_OUT, 'w', encoding='utf-8', newline='\n') as f:
        f.write(deploy)
    print("  Written OK")

    print(f"\nCopying standalone reference: {STANDALONE_OUT}")
    shutil.copy2(STANDALONE_SRC, STANDALONE_OUT)
    print("  Copied OK")

    print("\n[SUCCESS] Sync complete. Ready to commit.")
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
