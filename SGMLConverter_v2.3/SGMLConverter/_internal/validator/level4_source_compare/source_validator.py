"""
level4_source_compare/source_validator.py
──────────────────────────────────────────
Level 4: Source Comparison Validator  (30 points total)

Runs ONLY when a source PDF is provided. Compares the pipeline-generated
SGML against the source document across 6 validation dimensions:

  D2  Tagging accuracy        — font/style → tag correctness  (5 pts)
  D3  Text accuracy           — paragraph-level text diff      (8 pts)
  D4  Completeness            — count-based: tables, images,   (7 pts)
                                 footnotes, sections, pages
  D5  Ordering / sequence     — section + paragraph order      (4 pts)
  D6  Encoding & characters   — smart quotes, dashes, accents  (3 pts)
  D7  Metadata accuracy       — title, date, doc-number, lang  (3 pts)
  ─────────────────────────────────
  Total                                                        30 pts

Architecture decisions:
  • PyMuPDF (fitz) for text + font metadata — fast, free, good for single-column
  • pdfplumber for table structure — more reliable than fitz for grid detection
  • All deterministic — no GPT/API calls, consistent results
  • Graceful degradation — each sub-check has its own try/except;
    one failure doesn't block others
  • Source PDF is always provided in production — all 6 dimensions run

Limitations (by design):
  • Cannot validate SEMANTIC tag choice (BLOCK2 vs PART) — needs human judgment
  • Multi-column PDF layouts may produce ordering false-positives — flagged as WARNING
  • Scanned PDFs (no text layer) silently skip text diff — logged as warning
"""

from __future__ import annotations

import json
import os
import re
import unicodedata
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from typing import Optional

# ── Optional dependencies ─────────────────────────────────────────────────────
try:
    import fitz  # PyMuPDF
    _FITZ_OK = True
except ImportError:
    _FITZ_OK = False

try:
    import pdfplumber
    _PLUMBER_OK = True
except ImportError:
    _PLUMBER_OK = False

try:
    import requests as _requests
    import httpx as _httpx
    from anthropic import Anthropic as _Anthropic
    _LLM_DEPS_OK = True
except ImportError:
    _LLM_DEPS_OK = False

# ── LLM configuration (TR internal platform) ─────────────────────────────────
# Set _LLM_ENABLED = False to disable LLM augmentation and fall back to
# deterministic-only validation.
_LLM_ENABLED = True
_TR_AUTH_URL  = "https://aiplatform.gcs.int.thomsonreuters.com/v1/anthropic/token"
_TR_WORKSPACE_ID = "Saikumar3Y0Z"
_OPUS_MODEL   = "claude-opus-4-20250514"
_LLM_TIMEOUT  = 180          # seconds — Opus can be slow on long docs
_LLM_MAX_PARAS = 120         # max paragraphs sent per Opus call (token budget)
_LLM_MAX_CELLS = 80          # max table cells sent per Opus call

# Lazy-initialised singleton client (None until first LLM call succeeds)
_llm_client: "Optional[_Anthropic]" = None


def _get_llm_client() -> "Optional[_Anthropic]":
    """Return a cached Anthropic client authenticated via the TR internal platform.

    Returns None when LLM is disabled, dependencies are missing, or auth fails.
    Falls back gracefully — every call site checks for None before using the client.
    """
    global _llm_client, _LLM_ENABLED
    if not _LLM_ENABLED or not _LLM_DEPS_OK:
        return None
    if _llm_client is not None:
        return _llm_client
    try:
        resp = _requests.post(
            _TR_AUTH_URL,
            json={"workspace_id": _TR_WORKSPACE_ID, "model_name": _OPUS_MODEL},
            timeout=20,
            verify=False,
        )
        if resp.status_code != 200:
            _LLM_ENABLED = False   # stop retrying for this run
            return None
        data = resp.json()
        token = data.get("anthropic_api_key") or data.get("token", "")
        if not token:
            _LLM_ENABLED = False
            return None
        _llm_client = _Anthropic(api_key=token, http_client=_httpx.Client(verify=False))
        return _llm_client
    except Exception:
        _LLM_ENABLED = False
        return None


# ── D3 LLM: paragraph alignment + mutation/truncation detection ───────────────
_LLM_PARA_PROMPT = """\
You are a legal-document text-accuracy auditor.

I have extracted paragraphs from a SOURCE PDF and converted them into an SGML file.
Your job is to compare each PDF paragraph against the SGML paragraphs and report any
accuracy issues.

For EACH PDF paragraph:
1. Find the MOST LIKELY matching SGML paragraph (semantic match, not just keyword).
2. If a match exists, check whether:
   a. Any words are DELETED from the START of the SGML paragraph vs. the PDF paragraph.
   b. Any words are DELETED or INSERTED anywhere INSIDE the paragraph.
3. If NO matching SGML paragraph exists at all, set "found": false.

IMPORTANT rules:
- Minor punctuation / whitespace differences are NOT issues.
- Different paragraph numbering (e.g. "(a)" vs "(b)") is NOT a word mutation.
- Legal citations that differ only in format (s. 3(2) vs s.3(2)) are NOT issues.
- Structural reformatting (heading merged into body) is NOT an issue.
- Only flag SUBSTANTIVE word-level additions or deletions that change meaning.
- "deleted_start": ONLY set this if 5 or more substantive body-text words are missing from
  the very beginning of the matched SGML paragraph vs the PDF paragraph. Do NOT set it for:
  list markers like "(a)", "(b)", "1.", "i.", bullet characters;
  section/subsection labels like "3.1", "Section 2", "PART 2";
  short preamble labels of ≤4 words; or any structural prefix.
  If fewer than 5 substantive words are missing from the start, set deleted_start to null.
- "deleted_end": ONLY set this if 5 or more substantive body-text words are missing from
  the very END of the matched SGML paragraph. Do NOT flag footnote reference numbers,
  page numbers, or short trailing labels. If fewer than 5 words, set deleted_end to null.
- "mutations": list of {{"deleted": ["word1","word2"], "inserted": ["word3"]}} for mid-para
  word-level changes that are clearly substantive (not formatting/punctuation).

CRITICAL: Your entire response must be ONLY the raw JSON array starting with '[' and ending with ']'.
Do NOT include any explanation, commentary, or markdown fences. Start your response with '[' immediately.

PDF PARAGRAPHS (numbered 1..N):
{pdf_block}

SGML PARAGRAPHS (numbered 1..M):
{sgml_block}

Your response (raw JSON array only, start with '['):
[
  {{
    "pdf_idx": 1,
    "found": true,
    "best_sgml_idx": 3,
    "deleted_start": null,
    "deleted_end": null,
    "mutations": [],
    "confidence": 0.97
  }},
  ...
]
"""

_LLM_TABLE_PROMPT = """\
You are a legal-document table-accuracy auditor.

I extracted the following TABLE CELLS from a SOURCE PDF.
I then extracted the SGML document text (which may encode the table as body paragraphs
or as <TBLCELL> elements rather than a visual table).

Your job: for EACH table cell from the PDF, determine whether its content appears
anywhere in the SGML text (it may be in <TBLCELL>, <P>, <ITEM>, or any other tag).

Rules:
- Minor punctuation / number formatting differences (e.g. "$1,000" vs "1000") are OK.
- If the cell content is found anywhere in the SGML, mark found: true.
- If it is completely absent or the numbers/words are significantly different, mark found: false.
- Empty cells, single-letter cells, and pure-number cells under 4 digits are always found: true.

CRITICAL: Your entire response must be ONLY the raw JSON array starting with '[' and ending with ']'.
Do NOT include any explanation, commentary, or markdown. Start immediately with '['.

PDF TABLE CELLS (numbered 1..N):
{cells_block}

SGML TEXT (first 8000 characters):
{sgml_text}

Your response (raw JSON array only):
[
  {{"cell_idx": 1, "text": "...", "found": true}},
  ...
]
"""


def _extract_json_from_llm_response(text: str) -> str:
    """Extract a JSON array or object from LLM response text.

    Handles cases where Opus prepends explanatory text before the JSON block,
    or wraps it in markdown fences. Returns the raw JSON string or raises ValueError.
    """
    # 1. Strip markdown fences (```json ... ``` or ``` ... ```)
    text = re.sub(r"```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```", "", text)
    text = text.strip()

    # 2. Try to find a JSON array first (most common for our prompts)
    arr_match = re.search(r"\[.*\]", text, re.DOTALL)
    if arr_match:
        candidate = arr_match.group(0).strip()
        try:
            json.loads(candidate)
            return candidate
        except json.JSONDecodeError:
            pass

    # 3. Try a JSON object
    obj_match = re.search(r"\{.*\}", text, re.DOTALL)
    if obj_match:
        candidate = obj_match.group(0).strip()
        try:
            json.loads(candidate)
            return candidate
        except json.JSONDecodeError:
            pass

    # 4. Last resort: try the whole cleaned text
    json.loads(text)   # will raise if invalid
    return text


def _llm_align_paragraphs(
    pdf_paras: list[str],
    sgml_paras: list[str],
) -> "list[dict] | None":
    """Call Opus to align PDF paragraphs to SGML paragraphs and detect mutations.

    Returns a list of dicts (one per PDF paragraph) or None on failure.
    Batches automatically if > _LLM_MAX_PARAS paragraphs.
    """
    client = _get_llm_client()
    if client is None:
        return None

    def _fmt_block(paras: list[str], label: str) -> str:
        return "\n".join(f"[{i+1}] {p[:400]}" for i, p in enumerate(paras))

    all_results: list[dict] = []
    batch_size = _LLM_MAX_PARAS

    # Limit SGML paras to avoid token overflow; take first _LLM_MAX_PARAS*1.5
    sgml_capped = sgml_paras[:int(_LLM_MAX_PARAS * 1.5)]
    sgml_block  = _fmt_block(sgml_capped, "SGML")

    # Process PDF paras in batches
    for batch_start in range(0, len(pdf_paras), batch_size):
        batch = pdf_paras[batch_start : batch_start + batch_size]
        pdf_block = _fmt_block(batch, "PDF")
        prompt = _LLM_PARA_PROMPT.format(pdf_block=pdf_block, sgml_block=sgml_block)
        try:
            msg = client.messages.create(
                model=_OPUS_MODEL,
                max_tokens=4096,
                timeout=_LLM_TIMEOUT,
                messages=[{"role": "user", "content": prompt}],
            )
            raw = msg.content[0].text.strip()
            # Robust JSON extraction — handles preamble text and markdown fences
            json_str = _extract_json_from_llm_response(raw)
            batch_results: list[dict] = json.loads(json_str)
            # Adjust pdf_idx offset for batching
            for item in batch_results:
                item["pdf_idx"] = item.get("pdf_idx", 1) + batch_start
            all_results.extend(batch_results)
        except Exception:
            # On any parse/API error for this batch: mark all as unchecked
            for i in range(len(batch)):
                all_results.append({
                    "pdf_idx": batch_start + i + 1,
                    "found": True,
                    "best_sgml_idx": None,
                    "deleted_start": None,
                    "deleted_end": None,
                    "mutations": [],
                    "confidence": 0.0,
                    "_error": True,
                })
    return all_results if all_results else None


def _llm_verify_table_cells(
    pdf_cells: list[str],
    sgml_text: str,
) -> "list[dict] | None":
    """Call Opus to verify that each PDF table cell appears in the SGML.

    Returns list of {cell_idx, text, found} dicts or None on failure.
    """
    client = _get_llm_client()
    if client is None:
        return None

    cells_capped = pdf_cells[:_LLM_MAX_CELLS]
    cells_block  = "\n".join(f"[{i+1}] {c[:200]}" for i, c in enumerate(cells_capped))
    sgml_snippet = sgml_text[:8000]
    prompt = _LLM_TABLE_PROMPT.format(cells_block=cells_block, sgml_text=sgml_snippet)
    try:
        msg = client.messages.create(
            model=_OPUS_MODEL,
            max_tokens=2048,
            timeout=_LLM_TIMEOUT,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = msg.content[0].text.strip()
        json_str = _extract_json_from_llm_response(raw)
        return json.loads(json_str)
    except Exception:
        return None


# ── Character maps for D6 encoding checks (no PDF needed) ────────────────────
# Unicode characters that MUST be encoded as SGML entities in Carswell SGML
UNICODE_TO_ENTITY: dict[str, str] = {
    "\u2018": "&lsquo;",   # '  left single quote
    "\u2019": "&rsquo;",   # '  right single quote
    "\u201c": "&ldquo;",   # "  left double quote
    "\u201d": "&rdquo;",   # "  right double quote
    "\u2013": "&ndash;",   # –  en dash
    "\u2014": "&mdash;",   # —  em dash
    "\u00a0": "&nbsp;",    # non-breaking space
    "\u00e9": "&eacute;",  # é
    "\u00e8": "&egrave;",  # è
    "\u00ea": "&ecirc;",   # ê
    "\u00e0": "&agrave;",  # à
    "\u00f4": "&ocirc;",   # ô
    "\u00c9": "&Eacute;",  # É
    "\u00b0": "&deg;",     # °
    "\u00b7": "&middot;",  # ·
    "\u00d7": "&times;",   # ×
    "\u00b1": "&plusmn;",  # ±
    "\u00a9": "&copy;",    # ©
    "\u2022": "&bull;",    # •
    "\u2026": "&hellip;",  # …
    "\u20ac": "&euro;",    # €
    "\u2265": "&ge;",      # ≥
    "\u2264": "&le;",      # ≤
    "\u00bd": "&frac12;",  # ½
    "\u00bc": "&frac14;",  # ¼
}

# Characters where bare hyphen is used instead of proper dash entity
_DASH_CONTEXT_RE = re.compile(
    r"(?<=[a-z\d])\s+-\s+(?=[a-z\d])",  # word - word  (likely em/en dash)
    re.IGNORECASE,
)

# Legal citation patterns that are corruption-prone
_LEGAL_CITATION_RE = re.compile(
    r"\b(?:s\.|ss\.|art\.|para\.|cl\.|sch\.)\s*\d+(?:\(\w+\))*",
    re.IGNORECASE,
)

# Number pattern: integers, decimals, currency, percentages
_NUMBER_RE = re.compile(r"\$[\d,]+(?:\.\d+)?|\d+(?:,\d{3})*(?:\.\d+)?%?")

# ── Contact-detail patterns (D4-g/h) ─────────────────────────────────────────
# Email addresses (RFC 5321 simplified)
_EMAIL_RE_L4 = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", re.I)

# HTTP/HTTPS URLs and bare www. URLs
# Character class excludes whitespace, SGML angle brackets, quotes, and common
# URL-terminating bracket characters so URLs embedded in prose are captured cleanly.
_URL_RE_L4 = re.compile(r'(?:https?://|www\.)[^\s<>()\[\]{}"]+', re.I)

# North-American phone numbers: (NXX) NXX-XXXX / NXX-NXX-XXXX / +1-NXX-NXX-XXXX
# 3+3+4 digit structure naturally avoids regulation numbers (NN-NNN = 2+3 digits).
_PHONE_RE_L4 = re.compile(
    r'\b(?:\+?1[-. ]?)?'               # optional country code  +1 / 1
    r'(?:\(\d{3}\)|\d{3})'             # area code: (416) or 416
    r'[-. ]'                            # separator (required)
    r'\d{3}'                            # exchange (3 digits)
    r'[-. ]'                            # separator (required)
    r'\d{4}\b'                          # subscriber (4 digits — never matches NN-NNN)
)

# Canadian postal codes (A1A 1A1 format)
_POSTAL_CODE_RE = re.compile(r'\b[A-Z]\d[A-Z][ \t]*\d[A-Z]\d\b')

# Legitimate omission patterns — text in PDF not expected in SGML
_OMIT_PATTERNS = [
    re.compile(r"^\d{1,4}$"),                              # bare page numbers
    re.compile(r"^[Pp]age\s+\d{1,4}"),                    # "Page 1", "Page 12"
    re.compile(r"^table\s+of\s+contents", re.I),           # TOC header
    re.compile(r"copyright\s+©?\s*20\d{2}", re.I),         # copyright lines
    re.compile(r"^\s*(continued|suite)\s*$", re.I),        # continuation marks
    re.compile(r"^(home|trading|français|sign in)", re.I),  # web chrome
    re.compile(r"thomson\s+reuters", re.I),                 # TR branding
    # Standalone date lines ("June 9, 2023", "September 27, 2016")
    re.compile(
        r"^(?:january|february|march|april|may|june|july|august|september|"
        r"october|november|december|janvier|février|mars|avril|mai|juin|"
        r"juillet|août|septembre|octobre|novembre|décembre)\s+\d{1,2},?\s+\d{4}$",
        re.IGNORECASE,
    ),
    # Header/cover lines that are also bold (e.g. organisation names)
    re.compile(r"^\d{7}$"),  # internal reference numbers ("6217407")
    # TOC dot-leader entries: "Registrants ........................................................."
    re.compile(r"\.{5,}"),
    # Footnote/endnote lines: "19 See NI 33-109..." or "21 See subsection..."
    re.compile(r"^\d{1,3}\s+(?:See|Ibid|supra|infra|note\b)", re.IGNORECASE),
    # URL-only or source-only footnote lines: "5 Budget 2024: ... - Canada.ca"
    re.compile(r"^\d{1,3}\s+.{5,}\.(?:ca|com|gov|org|net)(?:[/\s]|$)", re.IGNORECASE),
    # OSC Bulletin page-citation running headers: "(2025), 48 OSCB 9737"
    # These appear in bold at the top of every page but are NOT body content.
    re.compile(r"^\(\d{4}\),\s+\d+\s+OSCB\s+\d+"),
    # OSC Bulletin section markers: "B. Ontario Securities Commission",
    # "B.1: Notices", "B.5: Rules and Policies" — bulletin nav headers, not doc content
    re.compile(r"^B\.\d*[:\s]"),
    # Legal party separator in court/regulatory documents: "- and -"
    re.compile(r"^[-\u2013]\s+and\s+[-\u2013]$"),
    # Date fragments split across PDF lines: "12, 2025", "10, 2025"
    re.compile(r"^\d{1,2},\s+\d{4}$"),
    # Parenthetical date references: "(as of September 17, 2025)"
    re.compile(r"^\(as\s+of\b", re.IGNORECASE),
    # Standard statutory header phrase in legal documents
    re.compile(r"^made\s+under\s+the\b", re.IGNORECASE),
    # Continuation fragments from page headers/footers (start with comma)
    re.compile(r"^,"),
    # Statistical table column headers: "Q3 2024", "Q4 2023"
    re.compile(r"^Q[1-4]\s+\d{4}$", re.IGNORECASE),
    # Percentage change table headers
    re.compile(r"^%\s+change\b", re.IGNORECASE),
    # Metadata date lines: "Date: 20250417"
    re.compile(r"^[Dd]ate:\s+\d"),
    # Standalone section/part labels from TOC (PDF renders in bold/italic)
    # e.g. "Part 3", "Part 7", "Section 2" — already encoded inside <TI> full text
    re.compile(r"^(?:part|section|chapter)\s+\d+[a-z]?\s*$", re.IGNORECASE),
    # Standalone annex/appendix labels: "Annex A", "Appendix B"
    re.compile(r"^(?:annex|appendix)\s+[a-zA-Z]\s*$", re.IGNORECASE),
    # Roman numeral TOC sub-entries: "i.  tick test", "iv.  short sale circuit breaker"
    re.compile(r"^(?:i{1,3}|iv|vi{0,3}|ix|xi{0,2}|xiv|xv)\.\s+\S", re.IGNORECASE),
    # Institution letterhead names (appear on every page header, not SGML body)
    re.compile(r"(?:securities\s+commission|securities\s+authority)\s*$", re.IGNORECASE),
    # CIRO / CIPF regulatory org names appearing as page-header elements
    re.compile(r"^(?:canadian\s+investment\s+regulatory\s+organization|cipf)\s*$", re.IGNORECASE),
    # PDF track-changes / annotation artifacts
    re.compile(r"\bstrikethrough\b", re.IGNORECASE),
    # Metadata form-field labels in regulatory submission forms
    re.compile(
        r"^(?:document\s+(?:type|no\.?|number|date|title)|effective\s+date|reference\s+no\.?)"
        r"\s*[:\s]*$",
        re.IGNORECASE,
    ),
    # Part/Section label WITH inline title text — heading already covered by <TI> check
    # e.g. "Part 1  Definitions", "Part 3 Effective Date"
    re.compile(
        r"^(?:part|section|article|division|schedule|item)\s+\d+[a-zA-Z]?\s+"
        r"(?:[-\u2013\u2014]|definitions?|purposes?|interpretation|effective\s+date|"
        r"general|application|exemption|transitional|repeal|hearings?|enforcement|amendments?|provisions?)",
        re.IGNORECASE,
    ),
    # FLI/FOFI forward-looking information abbreviation fragments
    re.compile(r"^fli[;,\s].*fofi", re.IGNORECASE),
    # Spaced-out decorative cover typography (e.g. "2 0 2 5", "5 5 / 2 0 2 5")
    # PyMuPDF extracts large-format year/issue numbers letter-by-letter
    re.compile(r"^(?:\S{1,2}\s+){2,}\S{1,2}$"),
    # OSC Bulletin section category labels: "Rules and Policies", "Notices"
    # These appear as bold navigation headers in the bulletin, not SGML body content
    re.compile(r"^(?:rules\s+and\s+policies|notices\s+and\s+news\s+releases)$", re.IGNORECASE),
    # Table row-number + pipe separator artifact from PDF table rendering
    # e.g. "1 |", "2 |", "25 |" — row labels extracted by PyMuPDF, not in SGML body
    re.compile(r"^\d+\s*\|", re.IGNORECASE),
    # Part/Section N – <any title> — italic TOC entries where the heading itself
    # is covered by <TI> but PDF TOC renders it italic with a dash separator
    re.compile(
        r"^(?:part|section|article|division|schedule|item)\s+\d+[a-zA-Z]?\s*[\u2013\u2014-]",
        re.IGNORECASE,
    ),
    # ── Exchange / regulator website footer boilerplate ───────────────────────
    # TMX / TSX / MX website navigation bars extracted as single merged paragraph
    # e.g. "Contact Us Terms of Use Privacy Policy Fraud Prevention TMX Group..."
    re.compile(r"contact\s+us\s+terms\s+of\s+use\s+privacy\s+policy", re.IGNORECASE),
    # TMX/TSX disclaimer: "The views, opinions and advice of any third party..."
    re.compile(r"views,?\s+opinions\s+and\s+advice\s+of\s+any\s+third\s+party", re.IGNORECASE),
    # CIRO footer: "You can find the Canadian Investment Regulatory Organization (CIRO) at CIRO.ca"
    re.compile(r"you\s+can\s+find\s+the\s+canadian\s+investment\s+regulatory", re.IGNORECASE),
    # MX / CDCC contact-info lines: "Market Operations Toll-free:" / "Derivatives Operations Toll-free:"
    re.compile(r"(?:market|derivatives)\s+operations\s+toll.free", re.IGNORECASE),
    # MX trademark line: "Canadian Derivatives Exchange is an official mark of Bourse de Montréal Inc."
    re.compile(r"canadian\s+derivatives\s+(?:exchange|clearing)\s+(?:is\s+an\s+official|corporation)", re.IGNORECASE),
    # CIRO internal distribution lists: "Distribute internally to: Corporate Finance, Credit..."
    re.compile(r"distribute\s+internally\s+to\s*:", re.IGNORECASE),
    # CIRO distribution-list continuation: "Corporate Finance, Credit, Institutional, Internal Audit..."
    re.compile(r"corporate\s+finance,\s+credit,\s+institutional,\s+internal\s+audit", re.IGNORECASE),
    # Government filing stamps: "Filed OCT 30 2025 Province of Saskatchewan Order in Council"
    re.compile(r"\bfiled\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b", re.IGNORECASE),
    # TSX Ops Notice / Production Alert header lines
    re.compile(r"(?:ops\s+notice|production\s+alert)\s+\d{4}", re.IGNORECASE),
    # Form blank lines: "___ Name of brokerage firm: ___..." (long underscore runs)
    re.compile(r"_{5,}"),
    # MX / TMX website capital formation nav: "Capital Formation Contact Us Terms of Use"
    re.compile(r"capital\s+formation\s+contact\s+us", re.IGNORECASE),
    # TMX copyright / trademark tail: "TMX Group Limited and its affiliates"
    re.compile(r"tmx\s+group\s+limited\s+and\s+its\s+affiliates", re.IGNORECASE),
    # TMX/MX website general disclaimer: "on this site or the content of any third party sites"
    re.compile(r"on\s+this\s+site\s+or\s+the\s+content\s+of\s+any\s+third\s+party\s+sites", re.IGNORECASE),
    # MX contact footer: "If you require additional information regarding this notice, please contact"
    re.compile(r"if\s+you\s+require\s+additional\s+information\s+regarding\s+this\s+notice", re.IGNORECASE),
    # TMX Production Alert / Ops Notice (without year requirement)
    re.compile(r"\btmx\s+production\s+alert\b", re.IGNORECASE),
    # Saskatchewan filing stamp with spaced OCR: "Province of Saskatchewan Order in Council"
    re.compile(r"province\s+of\s+saskatchewan\s+order\s+in\s+council", re.IGNORECASE),
    # Quebec AMF bulletin citation lines: "Bulletin de l'Autorité : 2023-06-01, Vol."
    re.compile(r"bulletin\s+de\s+l.autorit", re.IGNORECASE),
    # CIRO popup / website frame text: "Close this popup Welcome to CIRO"
    re.compile(r"close\s+this\s+popup", re.IGNORECASE),
    # Phase boilerplate in CIRO consultations: "Phase 5 of the Dealer Consolidated Rule Project"
    # (appears in PDF sidebar/banner, not in SGML body)
    re.compile(r"\bphase\s+\d+\s+of\s+the\s+dealer\s+consolidated\s+rule\s+project", re.IGNORECASE),
    # French regulatory authority contact blocks — appear in PDF footer/signature
    # sections of bilingual documents but are not encoded in the English SGML body.
    re.compile(r"autorit[e\xe9]\s+des\s+march[e\xe9]s", re.IGNORECASE),
    re.compile(r"direction\s+de\s+l[''\u2019]\s*encadrement", re.IGNORECASE),
    # Provincial minister signature blocks (PEI, etc.)
    re.compile(r"\bminister\s+of\s+justice\s+and\s+public\s+safety\b", re.IGNORECASE),
    # Amending instrument directives: "Part X ... is amended by the addition of the following"
    # These appear in the PDF amending text but are not body paragraphs in the SGML.
    re.compile(r"\bis\s+amended\s+by\s+the\s+addition\s+of\s+the\s+following\b", re.IGNORECASE),
    # CIRO track-changes / blackline revision markers
    # "insert end insert" / "end insert" are markup anchors in blacklined PDFs.
    re.compile(r"\bend\s+insert\b|\binsert\s+end\b", re.IGNORECASE),
    re.compile(r"\bblacklined\s+to\b", re.IGNORECASE),
    # CIRO website navigation text in PDF frame: "Welcome to CIRO.ca"
    re.compile(r"\bwelcome\s+to\s+ciro\.ca\b", re.IGNORECASE),
    # CIRO consultation heading: "Withdrawal of Proposed Amendments Respecting..."
    re.compile(r"\bwithdrawal\s+of\s+proposed\s+amendments\b", re.IGNORECASE),
    # CIRO PDF navigation/TOC frame row: "Description of Non-Material Changes"
    re.compile(r"\bdescription\s+of\s+non.material\s+changes\b", re.IGNORECASE),
    # CIRO regulatory topic sub-heading list (appears in document header navigation)
    re.compile(r"\binternal\s+investigation\s+and\s+client\s+complaint\b", re.IGNORECASE),
    # CIRO fee model navigation row: "Fee Model (clean) proposal publication"
    # Appears in CIRO blacklined consultation PDFs as a navigation/contents link.
    re.compile(r"\bfee\s+model\s*\(clean\)", re.IGNORECASE),
    # CIRO Rulebook navigation link: "Rulebook connection: IDPC Rules, UMIR Type..."
    # Appears in CIRO consultation PDFs as a document-header navigation block.
    re.compile(r"\brulebook\s+connection\b", re.IGNORECASE),
    # CIRO consultation navigation tab: tail fragment after fee model filter
    re.compile(r"\bproposal\s+publication\s*\)", re.IGNORECASE),
    # CIRO navigation tab: "Summary of Comments Received" section link
    re.compile(r"\bsummary\s+of\s+comments\s+received\b", re.IGNORECASE),
    # CIRO Rulebook navigation chapter: "Trading Desk" section name
    re.compile(r"\btrading\s+desk\b", re.IGNORECASE),
    # CIRO document navigation breadcrumb: "Type: Rules Bulletin"
    re.compile(r"\bType:\s+Rules\s+Bulletin\b", re.IGNORECASE),
    # CIRO consultation navigation tab: "Transitional Clarifications"
    re.compile(r"\btransitional\s+clarifications\b", re.IGNORECASE),
    # CIRO navigation cascade: "Comments Received The Investment Industry Regulatory..."
    re.compile(r"\bcomments\s+received\s+the\s+investment\s+industry\b", re.IGNORECASE),
    # Standalone "Terms of Use" / "Privacy Policy" — PDF website footer chrome, not SGML body
    re.compile(r"^terms\s+of\s+use$", re.IGNORECASE),
    re.compile(r"^privacy\s+policy$", re.IGNORECASE),
    # Single bullet/dash fragment lines: '• use of encryption;', 'o availability of...'
    # These are PDF bullet-list items where the bullet char is the first word
    re.compile(r"^[\u2022\u25e6\u25aa\u25b8\u2192\u25cb\u25cf\u25a0\u25a1\uf0b7o]\s+\S"),
    # TSX operational notice metadata fields — never encoded in SGML body text
    re.compile(r"^symbol\s+affected\s*:", re.IGNORECASE),
    re.compile(r"^division\s*:", re.IGNORECASE),
    # Hyperlink anchor text: "is available here." / "available here"
    re.compile(r"\bavailable\s+here\b", re.IGNORECASE),
    # Weekday + date lines: "Saturday, October 25, 2025." (PDF operational event dates)
    re.compile(
        r"^(?:monday|tuesday|wednesday|thursday|friday|saturday|sunday),\s+",
        re.IGNORECASE,
    ),
    # Document cross-reference line: "Companion Policy 11-332" (not SGML body content)
    re.compile(r"^companion\s+policy\b", re.IGNORECASE),
    # Dash-prefix PDF bullet items (SGML encodes as <ITEM> without the leading dash)
    # e.g. "- (ATI - TSXV)", "- Exec Report -", "- A (Trade)"
    re.compile(r"^-\s+\S"),
    # Regulatory rule/policy labels with instrument numbers
    # e.g. "RULE 11-803 (Amendment)", "BC POLICY 15-601 HEARINGS"
    re.compile(r"^rule\s+\d[\d-]", re.IGNORECASE),
    re.compile(r"^bc\s+policy\s+\d", re.IGNORECASE),
    # TMX operational department / section label (appears as PDF header, not SGML body)
    re.compile(r"^tmx\s+equity\s+trading\b", re.IGNORECASE),
    # Appendix with numeric reference: "Appendix 1", "Appendix 2" (extends letter-only match)
    re.compile(r"^appendix\s+\d", re.IGNORECASE),
    ]

# ── Font-name bold/italic detection (Gap 6 fix) ─────────────────────────────
# PyMuPDF font flags (bit4=bold) are unreliable for many PDFs that encode bold
# via the font name (e.g. "TimesNewRoman-Bold", "Helvetica-BoldOblique").
# These regexes complement the flags-based check.
_BOLD_FONT_NAME_RE = re.compile(
    r"(?i)(?:bold|heavy|black|demi|semibold|extrabold|ultra)"
)
_ITALIC_FONT_NAME_RE = re.compile(
    r"(?i)(?:italic|oblique|slanted)"
)


def _is_bold_font_name(font_name: str) -> bool:
    """Return True if the font name indicates bold weight."""
    return bool(_BOLD_FONT_NAME_RE.search(font_name))


def _is_italic_font_name(font_name: str) -> bool:
    """Return True if the font name indicates italic/oblique style."""
    return bool(_ITALIC_FONT_NAME_RE.search(font_name))


# ── Result dataclass ──────────────────────────────────────────────────────────
@dataclass
class L4Result:
    score: float = 0.0
    max_score: float = 30.0

    # Sub-scores (each dimension)
    tagging_score: float = 0.0      # D2: 0–5
    text_score: float = 0.0         # D3: 0–8
    completeness_score: float = 0.0 # D4: 0–7
    ordering_score: float = 0.0     # D5: 0–4
    encoding_score: float = 0.0     # D6: 0–3
    metadata_score: float = 0.0     # D7: 0–3

    issues: list[dict] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    # Diagnostics
    pdf_available: bool = False
    pdf_text_extractable: bool = False
    text_coverage: float = 0.0          # fraction of PDF paragraphs found in SGML
    missing_paragraphs: list[str] = field(default_factory=list)
    encoding_violations: list[str] = field(default_factory=list)
    sequence_violations: list[str] = field(default_factory=list)
    metadata_mismatches: list[str] = field(default_factory=list)

    # diff_generator fields — populated by check_* functions so the HITL
    # diff engine can produce line-specific, actionable fix suggestions
    d2_untagged_bold: list[str] = field(default_factory=list)
    d2_untagged_italic: list[str] = field(default_factory=list)
    d2_untagged_headings: list[str] = field(default_factory=list)
    d5_inverted_pairs: list[tuple] = field(default_factory=list)  # [(sgml_h_before, sgml_h_after), ...]
    d7_expected_lang: str = ""        # language the PDF suggests (for D7 LANG fix)
    d7_pdf_doc_number: str = ""       # doc number extracted from PDF (for D7 N-tag fix)
    pdf_headings: list[str] = field(default_factory=list)  # PDF headings (for D3 placement heuristic)

    # GAP 1: Two-stage validation — separate ABBYY extraction errors from pipeline errors
    docx_available: bool = False
    abbyy_missing_paragraphs: list[str] = field(default_factory=list)   # in PDF but not DOCX
    pipeline_missing_paragraphs: list[str] = field(default_factory=list) # in DOCX but not SGML
    # GAP confidence details — each dict: {text, confidence, method}
    abbyy_missing_paragraph_details: list[dict] = field(default_factory=list)
    pipeline_missing_paragraph_details: list[dict] = field(default_factory=list)
    # GAP 5: Table cell-level coverage — DOCX cells not found in SGML
    d4_missing_table_cells: list[dict] = field(default_factory=list)
    # D3-d: Paragraphs present in SGML but with leading text deleted
    truncated_paragraphs: list[str] = field(default_factory=list)

    # ── New comprehensive content-verification fields ─────────────────────────
    # D3-e: Paragraphs present in SGML but with words changed/added/deleted
    inline_changed_paragraphs: list[dict] = field(default_factory=list)

    # ── LLM-confirmed issues (subset of above — only from Opus analysis) ──────
    # These are RELIABLE — the deterministic lists above include FPs.
    # The escalation rule in validator_main.py uses ONLY these fields.
    llm_confirmed_truncations: int = 0    # D3-d confirmed by LLM
    llm_confirmed_mutations: int = 0      # D3-e confirmed by LLM
    # D3-f: Short lines (5–7 words) from PDF not found in SGML
    missing_short_lines: list[str] = field(default_factory=list)
    # D4-g: Contact details — bidirectional comparison (PDF ↔ SGML)
    missing_emails: list[str] = field(default_factory=list)
    extra_emails: list[str] = field(default_factory=list)
    missing_phones: list[str] = field(default_factory=list)
    extra_phones: list[str] = field(default_factory=list)
    missing_urls: list[str] = field(default_factory=list)
    extra_urls: list[str] = field(default_factory=list)
    missing_postal_codes: list[str] = field(default_factory=list)
    # D4-h: PDF table cells (direct pdfplumber extraction — no DOCX required)
    pdf_direct_table_cells_missing: list[dict] = field(default_factory=list)
    # D4-fn: Empty footnote body count — <FREEFORM> blocks inside <FOOTNOTE>
    # that contain no text (strong signal of deliberate content deletion).
    empty_footnote_bodies: int = 0


def _add_issue(result: L4Result, dimension: str, severity: str, description: str,
               location: str = "", impact: str = "") -> None:
    result.issues.append({
        "level": "L4",
        "category": dimension,
        "severity": severity,
        "description": description,
        "location": location,
        "impact": impact,
    })


# ── Entity decoding ──────────────────────────────────────────────────────────
# Map named SGML/HTML entities → Unicode characters.
# Used by _norm() so that SGML "l&rsquo;article" and PDF "l'article" compare equal.
_ENTITY_CHAR_MAP: dict[str, str] = {
    # Quotation marks / apostrophes
    "rsquo": "\u2019", "lsquo": "\u2018",
    "rdquo": "\u201d", "ldquo": "\u201c",
    "apos": "'",       "quot": '"',
    # Dashes
    "ndash": "\u2013", "mdash": "\u2014", "minus": "\u2212",
    # Spaces
    "nbsp": "\u00a0",  "ensp": "\u2002", "emsp": "\u2003",
    # XML built-ins
    "amp": "&", "lt": "<", "gt": ">",
    # French / accented Latin
    "eacute": "\u00e9", "egrave": "\u00e8", "ecirc": "\u00ea", "euml": "\u00eb",
    "Eacute": "\u00c9", "Egrave": "\u00c8", "Ecirc": "\u00ca",
    "agrave": "\u00e0", "acirc": "\u00e2", "auml": "\u00e4", "aring": "\u00e5",
    "Agrave": "\u00c0", "Acirc": "\u00c2",
    "ugrave": "\u00f9", "ucirc": "\u00fb", "uuml": "\u00fc",
    "icirc": "\u00ee", "iuml": "\u00ef",
    "ocirc": "\u00f4", "ouml": "\u00f6",
    "ccedil": "\u00e7", "Ccedil": "\u00c7",
    "oelig": "\u0153",  "OElig": "\u0152",
    "szlig": "\u00df",
    # Symbols
    "deg": "\u00b0",   "middot": "\u00b7", "times": "\u00d7", "plusmn": "\u00b1",
    "copy": "\u00a9",  "reg": "\u00ae",    "trade": "\u2122",
    "bull": "\u2022",  "hellip": "\u2026", "euro": "\u20ac",
    "ge": "\u2265",    "le": "\u2264",     "ne": "\u2260",
    "frac12": "\u00bd","frac14": "\u00bc", "frac34": "\u00be",
    "sect": "\u00a7",  "para": "\u00b6",   "dagger": "\u2020",
}


def _decode_sgml_entities(text: str) -> str:
    """Replace named SGML entities with their Unicode characters."""
    def _replace(m: re.Match) -> str:
        return _ENTITY_CHAR_MAP.get(m.group(1), " ")
    return re.sub(r"&([a-zA-Z][a-zA-Z0-9]*);", _replace, text)


# ── Text normalisation ────────────────────────────────────────────────────────
# Ligature map: PDF-extracted ligature chars → plain ASCII equivalents.
# pdfplumber and PyMuPDF sometimes return these as single Unicode chars;
# vendor SGML always has plain text. Without this, D3 n-gram matching fails
# on PDFs that use ligature glyphs (common in professionally typeset docs).
_LIGATURE_MAP: dict[str, str] = {
    "\uFB00": "ff",   # ﬀ
    "\uFB01": "fi",   # ﬁ
    "\uFB02": "fl",   # ﬂ
    "\uFB03": "ffi",  # ﬃ
    "\uFB04": "ffl",  # ﬄ
    "\uFB05": "st",   # ﬅ (long-s t)
    "\uFB06": "st",   # ﬆ
}


def _norm(text: str) -> str:
    """Decode entities → strip tags → normalise ligatures/quotes/dashes → lowercase."""
    text = _decode_sgml_entities(text)             # "&rsquo;" → "'"
    text = re.sub(r"<[^>]+>", " ", text)           # strip SGML tags
    text = unicodedata.normalize("NFC", text)
    # Ligature normalisation: ﬁ→fi, ﬂ→fl, ﬀ→ff, etc.
    # Must run before lowercasing so replacement chars are clean ASCII.
    for _lig, _plain in _LIGATURE_MAP.items():
        text = text.replace(_lig, _plain)
    # Normalise typographic variants to ASCII so PDF and SGML compare equal
    text = text.replace("\u2018", "'").replace("\u2019", "'")   # ' ' → '
    text = text.replace("\u201c", '"').replace("\u201d", '"')   # " " → "
    text = text.replace("\u2013", "-").replace("\u2014", "-")   # en/em dash → -
    text = text.replace("\u2212", "-")                          # minus sign → -
    text = text.replace("\u2011", "-")                          # non-breaking hyphen → -
    text = text.replace("\u00a0", " ")                          # nbsp → space
    text = text.lower()
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _is_omittable(text: str) -> bool:
    """Return True if this PDF text is legitimately absent from SGML."""
    t = text.strip()
    return any(p.search(t) for p in _OMIT_PATTERNS)


# ── PDF extraction (PyMuPDF) ──────────────────────────────────────────────────
@dataclass
class _PDFData:
    paragraphs: list[str] = field(default_factory=list)
    headings: list[str] = field(default_factory=list)
    bold_spans: list[str] = field(default_factory=list)
    italic_spans: list[str] = field(default_factory=list)
    table_count: int = 0
    image_count: int = 0
    footnote_count: int = 0
    page_count: int = 0
    first_page_text: str = ""
    language_hint: str = "EN"      # detected from character frequencies
    doc_title: str = ""
    doc_date: str = ""
    doc_number: str = ""
    two_column: bool = False        # Gap 9: True if 2-column layout detected
    ok: bool = True
    error: str = ""
    # Extended fields for comprehensive content verification
    raw_lines: list[str] = field(default_factory=list)         # all body lines (for short-line check)
    link_uris: list[str] = field(default_factory=list)         # PDF annotation hyperlink URIs
    pdf_table_cells: list[str] = field(default_factory=list)   # cell text from pdfplumber
    all_text: str = ""                                         # full unfiltered text from ALL pages (for contact search)


def _detect_language(text: str) -> str:
    """Detect document language using word-frequency ratio.

    CSA notices always include AMF/Quebec contact details with French accented
    characters (Autorite, marches, Quebec) even when the main document is English.
    A simple accented-char count falsely flags these as French documents.

    Strategy: count unambiguously French content words vs English content words.
    Only flag FR when French words strongly dominate (>30 words AND >2x English).
    This prevents false positives from French contact blocks in English CSA docs.
    """
    _fr_pat = (r'\b(et|des|les|pour|avec|dans|que|qui|une|par|sur|aux|est|sont|'
               r'cette|ce|ces|mais|donc|ni|aussi|comme|bien|'
               r'nous|vous|ils|elles|leur|leurs|dont|depuis|sans|entre|toute|'
               r'autre|autres|peut|doit|doivent|sera|seront|toutes|plusieurs|'
               r'chaque|selon|afin|lors|ainsi)\b')
    _en_pat = (r'\b(the|and|or|of|in|to|for|with|this|that|from|are|is|has|have|'
               r'be|by|an|at|its|it|as|on|not|which|will|shall|may|any|all|such|'
               r'under|been|were|also|where|when|would|should|could|their|these|'
               r'those|between|section|pursuant|including|required|must|each|'
               r'provide|person|persons|following|applies|means|make|report|within)\b')
    fr_count = len(re.findall(_fr_pat, text, re.I))
    en_count = len(re.findall(_en_pat, text, re.I))
    # Require strong French dominance to flag as FR —
    # prevents false positives from a French contact block in an English document
    if fr_count > 30 and en_count == 0:
        return "FR"
    if fr_count > 30 and en_count > 0 and (fr_count / en_count) > 2.0:
        return "FR"
    return "EN"


def _extract_doc_number(text: str) -> str:
    """Extract regulatory document number like NI 31-103, OSC Rule 14-501, etc."""
    # Normalise newlines so multiline spans like 'Notice\n11-326' are joined
    text = re.sub(r"\s+", " ", text)
    patterns = [
        r"\b(NI|MI|CSA|OSC|MSC|ASC|BCSC|AMF)\s+\d{2}-\d{3}\b",
        r"\b(?:National|Multilateral)\s+Instrument\s+(\d{2}-\d{3})\b",
        r"\b(?:Rule|Policy|Notice|Bulletin|Guideline)\s+(\d{2}-\d{3})\b",
        r"\b(\d{2}-\d{3})\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return re.sub(r"\s+", " ", m.group(0)).strip()
    return ""


def _extract_doc_date(text: str) -> str:
    """Extract date from first-page text. Returns YYYYMMDD or empty string."""
    # Match formats: "April 15, 2026", "15 April 2026", "2026-04-15", "April 2026"
    months = {
        "january": "01", "february": "02", "march": "03", "april": "04",
        "may": "05", "june": "06", "july": "07", "august": "08",
        "september": "09", "october": "10", "november": "11", "december": "12",
        "jan": "01", "feb": "02", "mar": "03", "apr": "04",
        "jun": "06", "jul": "07", "aug": "08", "sep": "09",
        "oct": "10", "nov": "11", "dec": "12",
        # French
        "janvier": "01", "février": "02", "mars": "03", "avril": "04",
        "mai": "05", "juin": "06", "juillet": "07", "août": "08",
        "septembre": "09", "octobre": "10", "novembre": "11", "décembre": "12",
    }
    # "April 15, 2026" or "April 2026"
    m = re.search(
        r"\b(" + "|".join(months) + r")\s+(\d{1,2})(?:,\s+|\s+)(\d{4})\b",
        text, re.IGNORECASE
    )
    if m:
        mo = months[m.group(1).lower()]
        day = m.group(2).zfill(2)
        yr = m.group(3)
        return f"{yr}{mo}{day}"
    # "15 April 2026"
    m = re.search(
        r"\b(\d{1,2})\s+(" + "|".join(months) + r")\s+(\d{4})\b",
        text, re.IGNORECASE
    )
    if m:
        day = m.group(1).zfill(2)
        mo = months[m.group(2).lower()]
        yr = m.group(3)
        return f"{yr}{mo}{day}"
    # ISO: "2026-04-15"
    m = re.search(r"\b(\d{4})-(\d{2})-(\d{2})\b", text)
    if m:
        return f"{m.group(1)}{m.group(2)}{m.group(3)}"
    # Labelled bare YYYYMMDD: "Date: 20250417" or "Date:  20250417"
    m = re.search(r"\bDate:\s*(\d{8})\b", text, re.IGNORECASE)
    if m:
        return m.group(1)
    return ""


def _extract_pdf_data(pdf_path: str) -> _PDFData:
    """Extract structured data from PDF using PyMuPDF + optional pdfplumber."""
    data = _PDFData()

    if not _FITZ_OK:
        data.ok = False
        data.error = "PyMuPDF (fitz) not installed"
        return data

    # Cap pages processed for text/font extraction to avoid hanging on very large PDFs.
    # Structural metadata (page count, image count, table count) still uses the full doc.
    _MAX_EXTRACT_PAGES = 80

    try:
        doc = fitz.open(pdf_path)
        data.page_count = len(doc)

        # Track repeated lines (headers/footers)
        line_freq: dict[str, int] = {}
        all_lines: list[str] = []

        for page_idx, page in enumerate(doc):
            if page_idx >= _MAX_EXTRACT_PAGES:
                break  # skip text/font extraction for pages beyond cap
            _page_h = page.rect.height  # GAP 3: for geometric header/footer filtering
            blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]

            # Extract PDF hyperlink annotations (external URIs only)
            try:
                for link in page.get_links():
                    if link.get("kind") == 2:  # kind=2 → external URI
                        uri = link.get("uri", "").strip()
                        if uri and uri not in data.link_uris:
                            data.link_uris.append(uri)
            except Exception:
                pass  # get_links() may fail on some PDFs

            for block in blocks:
                if block.get("type") != 0:
                    # Type 1 = image block
                    if block.get("type") == 1:
                        data.image_count += 1
                    continue
                # GAP 3: Skip blocks in header zone (top 10%) or footer zone (bottom 8%).
                # These are running headers/footers that ABBYY already excludes from DOCX.
                # Geometry-based filtering is more reliable than 100+ regex patterns.
                if _page_h > 0:
                    _bbox = block.get("bbox", (0, 0, 0, _page_h))
                    if _bbox[1] / _page_h < 0.10 or _bbox[3] / _page_h > 0.92:
                        continue
                for line in block.get("lines", []):
                    spans = line.get("spans", [])
                    if not spans:
                        continue
                    line_text = "".join(s["text"] for s in spans).strip()
                    if not line_text:
                        continue
                    all_lines.append(line_text)
                    line_freq[line_text] = line_freq.get(line_text, 0) + 1

                    # Bold / italic detection: flags AND font-name (Gap 6 fix)
                    # Flags alone miss PDFs that encode bold via font name only.
                    for span in spans:
                        flags = span.get("flags", 0)
                        font_name = span.get("font", "")
                        text = span["text"].strip()
                        if not text or len(text) < 3:
                            continue
                        is_bold = bool(flags & 16) or _is_bold_font_name(font_name)
                        is_italic = bool(flags & 2) or _is_italic_font_name(font_name)
                        if is_bold:
                            data.bold_spans.append(text)
                        if is_italic:
                            data.italic_spans.append(text)

        # Remove repeated lines (appear on ≥ 50% of pages, min 3 pages) — headers/footers
        threshold = max(3, data.page_count * 0.5)
        repeated = {ln for ln, cnt in line_freq.items() if cnt >= threshold}

        # Post-filter bold/italic span lists: remove running header/footer strings.
        # Bold/italic spans are accumulated during the page loop BEFORE we compute
        # `repeated`, so they contain header/footer text (e.g. italic running title,
        # bold date lines that appear on every page). Filter them out now.
        data.bold_spans = [s for s in data.bold_spans if s not in repeated]
        data.italic_spans = [s for s in data.italic_spans if s not in repeated]

        # Build page-1 text for metadata extraction (first 3 pages)
        first_pages_text = ""
        for page_idx in range(min(3, data.page_count)):
            first_pages_text += doc[page_idx].get_text()
        data.first_page_text = first_pages_text[:3000]

        # Build full-document text (all pages, no cap) for contact detail extraction
        _all_pages_text = first_pages_text  # reuse already-built first 3 pages
        for page_idx in range(3, data.page_count):
            _all_pages_text += doc[page_idx].get_text()
        data.all_text = _all_pages_text

        # Detect font sizes to identify headings
        # Collect (font_size, text) tuples from first _MAX_EXTRACT_PAGES pages
        size_text: list[tuple[float, str]] = []
        for page in list(doc)[:_MAX_EXTRACT_PAGES]:
            blocks = page.get_text("dict")["blocks"]
            for block in blocks:
                if block.get("type") != 0:
                    continue
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        t = span["text"].strip()
                        if t and len(t) > 3:
                            size_text.append((span.get("size", 0), t))

        # Body font = most common size
        if size_text:
            from collections import Counter
            size_counts = Counter(round(s, 0) for s, _ in size_text)
            body_size = size_counts.most_common(1)[0][0]
            heading_threshold = body_size * 1.15  # 15% larger = heading

            for size, text in size_text:
                if text in repeated:
                    continue
                if size >= heading_threshold and len(text) > 5:
                    data.headings.append(text)

        # Build clean paragraphs (remove repeated, omittable, short lines)
        # PyMuPDF extracts line-by-line; join consecutive lines into paragraphs
        # A line is a continuation if it doesn't end a sentence and the next is
        # short enough to be a wrapped line (< 80 chars).
        raw_lines: list[str] = []
        for ln in all_lines:
            if ln in repeated:
                continue
            if _is_omittable(ln):
                continue
            if len(ln.split()) < 2:
                continue
            raw_lines.append(ln)

        # Store raw_lines for short-line content check (D3-f)
        data.raw_lines = raw_lines

        # Merge continuation lines into paragraphs
        merged: list[str] = []
        buf = ""
        for ln in raw_lines:
            if not buf:
                buf = ln
            else:
                # Heuristic: join if previous line doesn't end in sentence-terminator
                # or current line starts lowercase / looks like a continuation
                prev_ends_sentence = buf.rstrip().endswith((".", "?", "!", ":", ";"))
                curr_starts_upper = ln[0].isupper() if ln else True
                is_short_prev = len(buf) < 70  # previous line was short (wrapped)
                if not prev_ends_sentence or (is_short_prev and not curr_starts_upper):
                    buf = buf.rstrip() + " " + ln
                else:
                    merged.append(buf)
                    buf = ln
        if buf:
            merged.append(buf)

        data.paragraphs = [ln for ln in merged if len(ln.split()) >= 4]

        # Footnote heuristic: lines with superscript-style numbering at start
        footnote_re = re.compile(r"^\d{1,3}\s+\S")
        data.footnote_count = sum(1 for ln in all_lines if footnote_re.match(ln))

        # Language detection from first page
        data.language_hint = _detect_language(first_pages_text)

        # Metadata extraction from first-page text
        data.doc_number = _extract_doc_number(first_pages_text)
        data.doc_date = _extract_doc_date(first_pages_text)

        # Table count via pdfplumber (more reliable).
        # Guard: skip pdfplumber for large PDFs (>400 KB) or many pages (>60)
        # where it can take several minutes — fall back to fitz heuristic.
        _pdf_size_kb = os.path.getsize(pdf_path) / 1024 if os.path.exists(pdf_path) else 0
        _use_plumber = _PLUMBER_OK and _pdf_size_kb <= 400 and data.page_count <= 60
        if _use_plumber:
            try:
                import pdfplumber as _plumber
                with _plumber.open(pdf_path) as plumb:
                    for pg in plumb.pages:
                        tables = pg.extract_tables()
                        if tables:
                            # Only count genuinely multi-column tables.
                            # pdfplumber detects single-column bordered callout boxes
                            # ("Practice Point", "Example", etc.) as tables — these
                            # are styled text blocks in the PDF, not real data tables.
                            # Vendor SGML correctly does not tag them as <TABLE>.
                            # Filter: require ≥2 columns to count as a real table.
                            for tbl in tables:
                                if not tbl:
                                    continue
                                max_cols = max((len(r) for r in tbl if r), default=0)
                                if max_cols < 2:
                                    continue  # skip single-col bordered boxes
                                data.table_count += 1
                                # Collect cell text only from real multi-column tables
                                for row in tbl:
                                    if not row:
                                        continue
                                    for cell in row:
                                        if cell:
                                            ct = str(cell).strip()
                                            if len(ct.split()) >= 2:
                                                data.pdf_table_cells.append(ct)
            except Exception:
                # Fallback: rough table detection from fitz line geometry
                data.table_count = _estimate_table_count_fitz(doc)
        else:
            data.table_count = _estimate_table_count_fitz(doc)

        # Full-doc image count (faster than text extraction — uses xref list)
        if data.image_count == 0:
            try:
                for _pg in doc:
                    data.image_count += len(_pg.get_images(full=False))
            except Exception:
                pass

        # Gap 9: detect 2-column layout BEFORE closing doc
        data.two_column = _detect_two_column_layout(doc)

        doc.close()

    except Exception as exc:
        data.ok = False
        data.error = str(exc)

    return data


def _detect_two_column_layout(doc) -> bool:
    """
    Detect if the PDF uses a 2-column text layout (Gap 9 fix).

    Heuristic: sample the first 5 pages; if the median text-line width is
    less than 55 % of the usable page width, the document is almost certainly
    laid out in two (or more) columns.  In that case D5 ordering results are
    unreliable because PyMuPDF reads text left-to-right across the full page
    width, mixing columns.
    """
    if not _FITZ_OK:
        return False

    line_widths: list[float] = []
    page_widths: list[float] = []

    for page_idx, page in enumerate(doc):
        if page_idx >= 5:  # sample first 5 pages only
            break
        pw = page.rect.width
        if pw <= 0:
            continue
        page_widths.append(pw)

        blocks = page.get_text("dict")["blocks"]
        for block in blocks:
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue
                line_text = "".join(s["text"] for s in spans).strip()
                if len(line_text) < 20:  # skip very short / label lines
                    continue
                bbox = line.get("bbox", [0, 0, 0, 0])
                line_w = bbox[2] - bbox[0]
                if line_w > 0:
                    line_widths.append(line_w)

    if not line_widths or not page_widths:
        return False

    avg_page_width = sum(page_widths) / len(page_widths)
    # Typical margins: ~10 % each side → usable width ≈ 80 % of page width
    usable_width = avg_page_width * 0.80

    line_widths_sorted = sorted(line_widths)
    median_lw = line_widths_sorted[len(line_widths_sorted) // 2]

    ratio = median_lw / usable_width if usable_width > 0 else 1.0
    # 2-column threshold: median text line < 55 % of usable page width
    return ratio < 0.55


def _estimate_table_count_fitz(doc) -> int:
    """Rough table count using horizontal line density heuristic (fitz fallback)."""
    count = 0
    for page in doc:
        paths = page.get_drawings()
        h_lines = [p for p in paths if p.get("rect") and
                   abs(p["rect"].height) < 3 and p["rect"].width > 50]
        if len(h_lines) >= 4:
            count += 1
    return count


# ── SGML text extraction ──────────────────────────────────────────────────────
def _extract_sgml_text(sgml: str) -> dict:
    """Extract text content, tag counts, and metadata from raw SGML."""
    # Strip all tags to get plain text, decoding entities → Unicode
    # (same normalisation as _norm() so SGML blob matches PDF text in D3)
    text_only = re.sub(r"<[^>]+>", " ", sgml)
    text_only = _decode_sgml_entities(text_only)
    text_only = unicodedata.normalize("NFC", text_only)
    text_only = text_only.replace("\u2018", "'").replace("\u2019", "'")
    text_only = text_only.replace("\u201c", '"').replace("\u201d", '"')
    text_only = text_only.replace("\u2013", "-").replace("\u2014", "-")
    text_only = text_only.replace("\u2212", "-")
    text_only = text_only.replace("\u00a0", " ")
    text_only = re.sub(r"\s+", " ", text_only).strip()

    # Extract paragraphs from all paragraph-like elements so word-level diff
    # has full coverage: P/P1/P2 tags PLUS ITEM, CLAUSE, LB, NOTE, BLOCK text
    paragraphs = []
    _para_tag_re = re.compile(
        r"<(?:P\d*|ITEM|CLAUSE|LB\d*|NOTE|PARA|BLOCKQUOTE)[^>]*>(.*?)"
        r"</(?:P\d*|ITEM|CLAUSE|LB\d*|NOTE|PARA|BLOCKQUOTE)>",
        re.DOTALL | re.IGNORECASE,
    )
    for pc in _para_tag_re.findall(sgml):
        clean = _norm(re.sub(r"<[^>]+>", " ", pc))
        if len(clean.split()) >= 4:
            paragraphs.append(clean)

    # Headings (TI tags)
    headings = [_norm(h) for h in re.findall(r"<TI[^>]*>(.*?)</TI>", sgml, re.DOTALL)]

    # Counts
    table_count = len(re.findall(r"<TABLE[\s>]", sgml))
    fn_count = len(re.findall(r"<FN[\s>]|<FOOTNOTE[\s>]", sgml))
    graphic_count = len(re.findall(r"<GRAPHIC\s", sgml))

    # POLIDOC metadata
    polidoc_m = re.search(r"<POLIDOC([^>]*)>", sgml)
    attrs = {}
    if polidoc_m:
        for am in re.finditer(r'(\w+)="([^"]*)"', polidoc_m.group(1)):
            attrs[am.group(1)] = am.group(2)

    # Section labels from N tags near TI tags (document numbering)
    sections = re.findall(r"<N[^>]*>(.*?)</N>", sgml, re.DOTALL)

    # GAP 5: extract SGML table cell text for cell-by-cell comparison
    sgml_table_cells: list[str] = []
    for _cell_content in re.findall(r"<TBLCELL[^>]*>(.*?)</TBLCELL>", sgml, re.DOTALL):
        _ct = _norm(re.sub(r"<[^>]+>", " ", _cell_content))
        if _ct and len(_ct.split()) >= 2:
            sgml_table_cells.append(_ct)

    # Extract footnote text for content accuracy check
    sgml_footnote_paras: list[str] = []
    for _fn_content in re.findall(
        r"<(?:FN|FOOTNOTE)[^>]*>(.*?)</(?:FN|FOOTNOTE)>", sgml, re.DOTALL | re.IGNORECASE
    ):
        _fc = _norm(re.sub(r"<[^>]+>", " ", _fn_content))
        if len(_fc.split()) >= 4:
            sgml_footnote_paras.append(_fc)

    # Extract SGML hyperlinks from XREF, EXTREF, A HREF attributes
    sgml_hrefs: list[str] = []
    for _href in re.findall(r'(?:HREF|href|URI|uri)="([^"]+)"', sgml):
        _href = _href.strip()
        if _href and _href not in sgml_hrefs:
            sgml_hrefs.append(_href)

    return {
        "text": text_only,
        "paragraphs": paragraphs,
        "headings": headings,
        "sections": [_norm(s) for s in sections],
        "table_count": table_count,
        "fn_count": fn_count,
        "graphic_count": graphic_count,
        "attrs": attrs,
        "table_cells": sgml_table_cells,   # GAP 5: TBLCELL text
        "footnote_paras": sgml_footnote_paras,  # footnote body text
        "sgml_hrefs": sgml_hrefs,               # hyperlink href values
    }


# ─────────────────────────────────────────────────────────────────────────────
# D6: Encoding & character accuracy (runs WITHOUT source PDF)
# ─────────────────────────────────────────────────────────────────────────────
def check_encoding(raw_sgml: str, result: L4Result) -> None:
    """
    D6 — 3 pts: Detect Unicode characters that must be encoded as SGML entities.

    Checks text content inside tags for raw Unicode that should be an entity.
    Does NOT require the source PDF — runs on SGML alone.
    """
    score = 3.0

    # Extract text content (inside tags, after stripping tags)
    # We need to check the raw content, not the tag attributes
    # Strip attribute values and tag markup, keep text content
    text_content = re.sub(r"<[^>]+>", "\x00", raw_sgml)  # replace tags with null
    # text_content now has text nodes separated by nulls

    violations: list[str] = []
    char_counts: dict[str, int] = {}

    for char, entity in UNICODE_TO_ENTITY.items():
        count = text_content.count(char)
        if count > 0:
            char_counts[char] = count
            violations.append(f"Raw U+{ord(char):04X} ({entity}) found {count}× — use {entity}")

    # Check for bare hyphen used as dash in mid-sentence
    dash_misuse = len(_DASH_CONTEXT_RE.findall(text_content))
    if dash_misuse > 0:
        violations.append(f"Bare hyphen used as dash {dash_misuse}× — use &ndash; or &mdash;")

    # Check for straight quotes in running text (not in tag attributes)
    # Tag attributes already handled — check content only
    straight_dq = text_content.count('"')
    if straight_dq > 5:  # allow a few in quoted material
        violations.append(f"Straight double-quotes {straight_dq}× — use &ldquo;/&rdquo;")

    if violations:
        n = len(violations)
        pts = min(2.0, n * 0.4)
        score -= pts
        result.encoding_violations = violations  # store all — no cap
        severity = "major" if pts >= 1.0 else "minor"
        _add_issue(result, "encoding", severity,
                   f"D6 — {n} encoding violation(s): raw Unicode instead of SGML entities. "
                   f"First: {violations[0]}",
                   impact=f"-{pts:.1f} pts")

    result.encoding_score = max(0.0, score)


# ─────────────────────────────────────────────────────────────────────────────
# D2: Tagging accuracy (requires PDF)
# ─────────────────────────────────────────────────────────────────────────────
def check_tagging(pdf: _PDFData, sgml_data: dict, raw_sgml: str, result: L4Result,
                  docx_data: "dict | None" = None) -> None:
    """
    D2 — 5 pts: Validate that PDF formatting is reflected by correct SGML tags.

    Checks:
    - Bold text in PDF → <BOLD> or <EM> tag in SGML
    - Italic text in PDF → <EM> or <ITALIC> tag in SGML
    - Headings in PDF (larger font) → <TI> in SGML
    - Tables in PDF → <TABLE> count approximately matches
    - Images in PDF → <GRAPHIC> count approximately matches

    GAP 4: When docx_data is provided, use DOCX bold/italic runs as the
    authoritative source instead of PDF font-flag spans. python-docx run.bold
    and run.italic read directly from OOXML w:rPr elements — far more reliable
    than PyMuPDF font-flag inference.
    """
    score = 5.0

    # Pre-compute <LINE> TOC text — reused by D2-a, D2-b, D2-c to avoid flagging
    # TOC entries that render bold/italic in the PDF but are correctly encoded as <LINE>.
    _line_texts = {
        _norm(re.sub(r"<[^>]+>", " ", ln))
        for ln in re.findall(r"<LINE[^>]*>(.*?)</LINE>", raw_sgml, re.DOTALL)
        if len(ln.split()) >= 2
    }

    # Pre-compute full SGML body plain-text blob (tags stripped) — used in D2-a and D2-b
    # as the final coverage fallback: if the text exists anywhere in SGML, it's present;
    # the vendor just chose a different tag (or no tag) which may be acceptable.
    _sgml_body_blob = _norm(re.sub(r"<[^>]+>", " ", raw_sgml))

    # D2-a: Bold spans from PDF should appear wrapped in BOLD/EM in SGML
    bold_tagged = re.findall(r"<(?:BOLD|EM)[^>]*>(.*?)</(?:BOLD|EM)>", raw_sgml, re.DOTALL)
    bold_tagged_text = {_norm(t) for t in bold_tagged}

    # GAP-7 FIX: Amendment instruments encode quoted replacement text in <QUOTE> tags.
    # This text renders bold in the PDF ("replacing X with Y") but is NOT wrapped in
    # <BOLD> in SGML — the <QUOTE> tag itself implies the content. Add to covered set.
    for qt in re.findall(r"<QUOTE[^>]*>(.*?)</QUOTE>", raw_sgml, re.DOTALL):
        bold_tagged_text.add(_norm(qt))

    # Also collect TI heading text — bold spans that are part of headings
    # are legitimately not wrapped in <BOLD> (the heading tag itself implies bold)
    ti_texts = {_norm(h) for h in sgml_data["headings"]}

    # GAP 4: prefer DOCX bold runs (explicit OOXML markup) over PDF font-flag spans.
    # Fall back to PDF spans if DOCX unavailable or has no bold runs detected.
    _bold_source = (docx_data.get("bold_runs") or []) if (docx_data and docx_data.get("bold_runs")) else pdf.bold_spans

    untagged_bold = []
    for span in _bold_source:  # check all bold spans — no sampling cap
        norm_span = _norm(span)
        if len(norm_span.split()) < 2:
            continue  # skip single words, too noise-prone
        if _is_omittable(span):
            continue
        in_bold_em = any(norm_span in bt or bt in norm_span for bt in bold_tagged_text)
        in_heading = any(norm_span in th or th in norm_span for th in ti_texts)
        # Fragment match: PDF splits long bold runs into short line-spans; match against longer SGML text
        if not in_bold_em:
            in_bold_em = any(
                norm_span in bt or SequenceMatcher(None, norm_span, bt).ratio() >= 0.85
                for bt in bold_tagged_text if len(bt) >= len(norm_span)
            )
        if not in_bold_em and not in_heading:
            in_heading = any(
                norm_span in th or SequenceMatcher(None, norm_span, th).ratio() >= 0.85
                for th in ti_texts if len(th) >= len(norm_span)
            )
        # <LINE> TOC coverage: bold TOC entries encoded as <LINE> in SGML, not <BOLD>
        if not in_bold_em and not in_heading:
            in_bold_em = any(norm_span in ln or ln in norm_span for ln in _line_texts)
        # GAP-4 FIX: If bold span text appears anywhere in SGML plain text, the vendor
        # has the content — they just encoded it without <BOLD> (e.g. short fragments
        # from multi-line bold blocks, provision labels, commission names in body text).
        if not in_bold_em and not in_heading:
            in_bold_em = norm_span in _sgml_body_blob
        if not in_bold_em and not in_heading:
            untagged_bold.append(span[:60])

    # Always store for diff_generator (even if no issue)
    result.d2_untagged_bold = untagged_bold

    if untagged_bold:
        ratio = len(untagged_bold) / max(1, len([s for s in _bold_source if len(s.split()) >= 2]))
        pts = min(1.5, ratio * 3.0)
        score -= pts
        severity = "major" if ratio > 0.3 else "minor"
        _add_issue(result, "tagging_accuracy", severity,
                   f"D2 — {len(untagged_bold)} bold text span(s) from PDF not wrapped in "
                   f"<BOLD> or <EM> in SGML. Examples: {untagged_bold[:3]}",
                   impact=f"-{pts:.1f} pts")

    # D2-b: Italic spans from PDF → <EM> or <ITALIC>
    # Important: text inside <TI> heading tags is already styled — it does NOT
    # need a separate <EM> wrapper. Exclude heading text from the italic check
    # to avoid false positives (e.g. '<TI>Trade Execution</TI>' is correct;
    # flagging it as "missing <EM>" is wrong).
    italic_tagged = re.findall(r"<(?:EM|ITALIC)[^>]*>(.*?)</(?:EM|ITALIC)>", raw_sgml, re.DOTALL)
    italic_tagged_text = {_norm(t) for t in italic_tagged}
    # Add all TI heading text as implicitly covered (heading styling implies italic/bold)
    italic_tagged_text.update(ti_texts)  # ti_texts defined in D2-a

    # GAP-1 FIX: Text already in <BOLD> does NOT need <EM> too.
    # Bold-italic PDF spans (e.g. Phase 1/2/3 headings) are correctly encoded
    # as <BOLD> only in SGML. Adding bold text to covered set prevents false positives.
    italic_tagged_text.update(bold_tagged_text)

    # GAP-2 FIX: Add all plain text from inside <FOOTNOTE> blocks to covered set.
    # Legislation Act/Bill names cited in footnotes are italic in PDF but the vendor
    # correctly puts them in <FOOTNOTE><FREEFORM><P> without wrapping in <EM>.
    _fn_blob = " ".join(
        _norm(re.sub(r"<[^>]+>", " ", fn))
        for fn in re.findall(r"<FOOTNOTE[^>]*>(.*?)</FOOTNOTE>", raw_sgml, re.DOTALL)
    )

    # (_line_texts pre-computed before D2-a above — reused here for italic TOC coverage)

    # (D2-b) _sgml_body_blob is pre-computed above, shared with D2-a.
    # GAP 4: prefer DOCX italic runs over PDF font-flag spans.
    _italic_source = (docx_data.get("italic_runs") or []) if (docx_data and docx_data.get("italic_runs")) else pdf.italic_spans

    untagged_italic = []
    for span in _italic_source:  # check all italic spans — no sampling cap
        norm_span = _norm(span)
        if len(norm_span.split()) < 2:
            continue
        # Direct match: span text appears inside an EM/ITALIC/TI/BOLD tag
        found = any(norm_span in it or it in norm_span for it in italic_tagged_text)
        # GAP-4 FIX: fragment match — span may be a PyMuPDF line-fragment of a longer EM span
        if not found:
            found = any(
                norm_span in em_full or SequenceMatcher(None, norm_span, em_full).ratio() >= 0.85
                for em_full in italic_tagged_text if len(em_full) >= len(norm_span)
            )
        # GAP-2 FIX: span may be inside a footnote (legislation names, citations)
        if not found:
            found = norm_span in _fn_blob
        # TOC FIX: span may be a <LINE> TOC entry (italic in PDF, no <EM> in SGML — by design)
        if not found:
            found = any(norm_span in ln or ln in norm_span for ln in _line_texts)
        # GAP-2b FIX: span text present in SGML body as plain text — vendor omitted <EM>
        # which is acceptable for legislation names, bibliography entries, Q&A headings.
        # Only flag if the text is completely absent from SGML.
        if not found:
            found = norm_span in _sgml_body_blob
        # PDF TOC entries often have trailing punctuation (comma, dash) not in SGML headings.
        # Strip trailing punctuation and retry body blob match.
        if not found:
            norm_stripped = norm_span.rstrip('- ,;\u2013\u2014')
            if len(norm_stripped.split()) >= 2:
                found = norm_stripped in _sgml_body_blob
        if not found and not _is_omittable(span):
            untagged_italic.append(span[:60])

    result.d2_untagged_italic = untagged_italic

    if untagged_italic:
        ratio = len(untagged_italic) / max(1, len([s for s in _italic_source if len(s.split()) >= 2]))
        pts = min(1.0, ratio * 2.0)
        score -= pts
        _add_issue(result, "tagging_accuracy", "minor",
                   f"D2 — {len(untagged_italic)} italic span(s) from PDF not wrapped in "
                   f"<EM> or <ITALIC> in SGML. Examples: {untagged_italic[:3]}",
                   impact=f"-{pts:.1f} pts")

    # D2-c: PDF headings (larger font) → <TI> in SGML
    # Some documents legitimately encode the document title/notice heading as
    # <P><BOLD>...</BOLD></P> rather than <TI>. Accept that as covered if the
    # heading text matches bold-tagged paragraph content.
    bold_para_texts: set[str] = set()
    for bp in re.findall(r"<(?:BOLD|EM)[^>]*>(.*?)</(?:BOLD|EM)>", raw_sgml, re.DOTALL):
        norm_bp = _norm(bp)
        if len(norm_bp.split()) >= 2:
            bold_para_texts.add(norm_bp)

    sgml_ti_texts_h = {_norm(h) for h in sgml_data["headings"]}  # TI tags
    # Include <N> document identifiers — PDF headings often carry a label prefix
    # e.g. PDF: "CSA Staff Notice 41-307 (Revised)" vs SGML N: "41-307 (Revised)"
    _n_texts = {
        _norm(n) for n in re.findall(r"<N[^>]*>(.*?)</N>", raw_sgml, re.DOTALL)
        if len(n.strip()) >= 4
    }
    _all_sgml_headings = sgml_ti_texts_h | _n_texts

    untagged_headings = []
    for heading in pdf.headings:  # check all headings — no sampling cap
        norm_h = _norm(heading)
        if len(norm_h.split()) < 2:
            continue
        # Check TI match (fuzzy)
        found_ti = any(
            SequenceMatcher(None, norm_h, sh).ratio() >= 0.70
            for sh in sgml_ti_texts_h
        )
        # Check if any SGML TI or N text is a substring of the PDF heading
        # e.g. "csa staff notice 41-307 (revised)" contains "41-307 (revised)"
        # Threshold >=5 (not 8) so short doc numbers like "46-309" (6 chars) match.
        if not found_ti:
            found_ti = any(sh in norm_h for sh in _all_sgml_headings if len(sh) >= 5)
        # Accept if heading text is encoded as bold paragraph (valid alternative)
        found_bold = any(
            norm_h in bp or bp in norm_h or SequenceMatcher(None, norm_h, bp).ratio() >= 0.75
            for bp in bold_para_texts
        )
        # <LINE> TOC coverage: heading that appears in PDF TOC encoded as <LINE> in SGML
        if not found_ti and not found_bold:
            found_ti = any(norm_h in ln or ln in norm_h for ln in _line_texts)
        # Body blob fallback: heading content present in SGML but not tagged as <TI>
        # (e.g. encoded as <P><BOLD> or plain <P> — valid alternative encoding)
        if not found_ti and not found_bold:
            found_ti = norm_h in _sgml_body_blob
        # Reverse substring: any SGML body fragment >=5 chars contained in PDF heading
        if not found_ti and not found_bold:
            found_ti = any(sh in norm_h for sh in _all_sgml_headings if len(sh) >= 5)
        if not found_ti and not found_bold and not _is_omittable(heading):
            untagged_headings.append(heading[:80])

    result.d2_untagged_headings = untagged_headings

    if untagged_headings:
        ratio = len(untagged_headings) / max(1, len(pdf.headings))
        pts = min(1.5, ratio * 2.0)
        score -= pts
        severity = "major" if ratio > 0.4 else "minor"
        _add_issue(result, "tagging_accuracy", severity,
                   f"D2 — {len(untagged_headings)} heading(s) detected in PDF (larger font) "
                   f"not tagged as <TI> in SGML. Examples: {untagged_headings[:3]}",
                   impact=f"-{pts:.1f} pts")

    # D2-d: Image count match (PDF images → GRAPHIC tags)
    # NOTE: Vendor SGML practice: small image counts (≤5) are almost always logos,
    # decorative borders, or signature blocks that are intentionally NOT tagged as
    # <GRAPHIC> in SGML. Only flag when there are MANY images (>5) and SGML has none,
    # which suggests actual content graphics were dropped.  Count mismatch is only
    # flagged when SGML has *some* GRAPHICs (meaning the encoder knows about them).
    sgml_graphic_count = sgml_data["graphic_count"]
    if pdf.image_count > 5 and sgml_graphic_count == 0:
        pts = 0.5
        score -= pts
        _add_issue(result, "tagging_accuracy", "minor",
                   f"D2 — PDF has {pdf.image_count} image(s) but SGML has no <GRAPHIC> tags.",
                   impact=f"-{pts:.1f} pt")
    elif pdf.image_count > 0 and sgml_graphic_count > 0 and abs(sgml_graphic_count - pdf.image_count) > 2:
        pts = 0.5
        score -= pts
        _add_issue(result, "tagging_accuracy", "minor",
                   f"D2 — PDF has {pdf.image_count} image(s) but SGML has "
                   f"{sgml_graphic_count} <GRAPHIC> tag(s). Count mismatch.",
                   impact=f"-{pts:.1f} pt")

    result.tagging_score = max(0.0, score)


# ─────────────────────────────────────────────────────────────────────────────
# DOCX text extraction helper (GAP 1 — two-stage validation)
# ─────────────────────────────────────────────────────────────────────────────
def _extract_docx_text(docx_path: str) -> dict:
    """
    Extract ALL text from an ABBYY-generated DOCX, including table cells.

    python-docx's doc.paragraphs iteration MISSES table cells — they must be
    extracted explicitly via doc.tables.

    Returns a dict:
      ok             – True if extraction succeeded
      error          – error string on failure
      paragraphs     – list of non-empty paragraph strings (≥2 words)
      table_cells    – list of non-empty, de-duplicated table cell strings
      combined_text  – lowercase normalised blob used for n-gram matching
    """
    try:
        from docx import Document as _DocxDocument  # python-docx
        doc = _DocxDocument(docx_path)

        paragraphs: list[str] = []
        bold_runs: list[str] = []
        italic_runs: list[str] = []
        for para in doc.paragraphs:
            t = para.text.strip()
            if t and len(t.split()) >= 2:
                paragraphs.append(t)
            # GAP 4: Extract bold/italic runs directly from DOCX — more reliable
            # than PyMuPDF font-flag inference. python-docx run.bold/run.italic
            # come directly from the OOXML w:rPr/w:b and w:i elements.
            for run in para.runs:
                rt = run.text.strip()
                if not rt or len(rt.split()) < 2:
                    continue
                if run.bold:
                    bold_runs.append(rt)
                if run.italic:
                    italic_runs.append(rt)

        # Explicitly walk tables — doc.paragraphs skips table cells entirely
        table_cells: list[str] = []
        for table in doc.tables:
            seen: set[str] = set()          # deduplicate merged/repeated cells
            for row in table.rows:
                for cell in row.cells:
                    t = cell.text.strip()
                    if t and t not in seen:
                        seen.add(t)
                        table_cells.append(t)

        all_text = paragraphs + table_cells
        combined = " ".join(_norm(t) for t in all_text)
        return {
            "ok": True,
            "error": "",
            "paragraphs": paragraphs,
            "table_cells": table_cells,
            "combined_text": combined,
            "bold_runs": bold_runs,    # GAP 4: runs explicitly marked bold in DOCX
            "italic_runs": italic_runs, # GAP 4: runs explicitly marked italic in DOCX
        }
    except Exception as exc:
        return {
            "ok": False,
            "error": str(exc),
            "paragraphs": [],
            "table_cells": [],
            "combined_text": "",
        }


# ─────────────────────────────────────────────────────────────────────────────
# Multi-stage paragraph coverage check (GAP 2 — replaces nested _para_covered)
# ─────────────────────────────────────────────────────────────────────────────
def _para_covered_v2(
    words: list[str],
    blob: str,
    ngrams: set,
    ngram_size: int = 5,
) -> tuple[bool, float, str]:
    """
    Multi-stage paragraph match. Returns (is_covered, confidence, method).

    Stage 1 – Exact 5-gram: fast path for clean identical text.
    Stage 2 – Fuzzy word coverage: ≥90 % of individual words present
              (handles single-word additions / deletions that break n-grams).
    Stage 3 – Sentence-level SequenceMatcher: catches minor rephrasing in
              short-to-medium paragraphs (≤25 words).
    Stage 4 – Chunked window: handles over-merged PyMuPDF paragraphs that
              span multiple SGML elements.
    """
    if not words:
        return (False, 0.0, "empty")

    # Normalise: strip trailing punctuation fused to words by PDF extraction
    # e.g. 'operation;' → 'operation', 'information;' → 'information'.
    # SGML tag-stripping inserts spaces between words and punctuation, so the
    # SGML blob has them separated.  PDF text does not.
    words = [re.sub(r"[;,.:]+$", "", w) for w in words]
    # Strip surrounding straight quotes (from _norm() curly-quote → " conversion).
    # e.g. '"originally' → 'originally', 'by"' → 'by', '"offered' → 'offered'.
    words = [w.strip('"\'') for w in words]
    # Strip Private Use Area unicode (font-specific bullets fused to words).
    # e.g. \uf0b7in (Wingdings bullet + 'in') → 'in'; standalone \uf0b7 → ''.
    words = [re.sub(r'[\uf000-\uf8ff]+', '', w) for w in words]
    # Strip trailing footnote superscript digits fused to words.
    # e.g. 'registered6' → 'registered', '(ciro)1' → '(ciro)'.
    # Only applies when the word contains letters — prevents mangling hyphenated
    # rule/regulation numbers like '91-102' → '91-10'.
    words = [re.sub(r'\d+$', '', w) if re.search(r'[a-z]', w) else w for w in words]
    # Strip surrounding parentheses/brackets from abbreviation tokens.
    # PDF inline abbreviations like '(UMIR)' → '(umir)' → after strip: 'umir'.
    # SGML may not use the parenthetical form.
    # e.g. '(ni' → 'ni', '41-101)' → '41-101', '(umir)' → 'umir'.
    words = [re.sub(r'^[\[(]+|[\])]+$', '', w) for w in words]
    words = [w for w in words if w]  # drop any that became empty
    if not words:
        return (False, 0.0, "empty")

    # Very short text: word-presence check
    if len(words) < ngram_size:
        found = sum(1 for w in words if w in blob)
        cov = found / len(words)
        return (cov >= 0.85, cov, "short_word_match")

    grams = [tuple(words[i:i + ngram_size]) for i in range(len(words) - ngram_size + 1)]
    if not grams:
        return (False, 0.0, "no_grams")

    # Stage 1: exact n-gram match
    matched = sum(1 for g in grams if g in ngrams)
    ngram_ratio = matched / len(grams)
    if ngram_ratio >= 0.55:
        return (True, ngram_ratio, "exact_ngram")

    # Stage 2: fuzzy word coverage
    words_found = sum(1 for w in words if w in blob)
    word_cov = words_found / len(words)
    if word_cov >= 0.875:  # allows ~3-4 missing words per 33-word paragraph
        return (True, word_cov, "fuzzy_word_coverage")

    # Stage 3: sentence-level similarity (short paragraphs only — performance guard)
    if len(words) <= 25:
        para_text = " ".join(words)
        best_sim = 0.0
        for seg in re.split(r'[.!?]\s+', blob):
            seg_words = seg.split()
            if len(seg_words) >= max(3, len(words) // 2):
                sim = SequenceMatcher(None, para_text, seg).ratio()
                if sim > best_sim:
                    best_sim = sim
                    if best_sim >= 0.82:
                        break  # good enough — stop early
        if best_sim >= 0.82:
            return (True, best_sim, "sentence_fuzzy")

    # Stage 4: chunked window check (handles over-merged PDF paragraphs)
    chunk_size, step = 15, 8
    if len(words) > chunk_size:
        windows_covered = windows_total = 0
        for start in range(0, len(words) - chunk_size + 1, step):
            chunk = words[start:start + chunk_size]
            cgrams = [tuple(chunk[i:i + ngram_size]) for i in range(len(chunk) - ngram_size + 1)]
            if not cgrams:
                continue
            windows_total += 1
            if sum(1 for g in cgrams if g in ngrams) / len(cgrams) >= 0.60:
                windows_covered += 1
        if windows_total > 0 and (windows_covered / windows_total) >= 0.55:
            return (True, windows_covered / windows_total, "chunked_windows")

    best_conf = max(ngram_ratio, word_cov)
    return (False, best_conf, "no_match")


# ─────────────────────────────────────────────────────────────────────────────
# D3-e: Word-level mutation detection helper
# ─────────────────────────────────────────────────────────────────────────────
def _word_diff_result(pdf_para: str, sgml_paras: list[str]) -> "dict | None":
    """
    Detect inline word mutations in a PDF paragraph that passed coverage check.

    Compares the normalized PDF paragraph against all SGML paragraphs (from
    <P>, <ITEM>, <CLAUSE>, <LB> etc.) and finds the closest match. If the
    best match ratio is 0.68–0.91, the paragraph is PRESENT but MUTATED
    (words added, deleted, or changed).

    Returns a dict with mutation details, or None if the paragraph is clean.
    """
    pdf_words = _norm(pdf_para).split()
    if len(pdf_words) < 4:
        return None

    best_ratio = 0.0
    best_sgml = ""

    for sp in sgml_paras:
        sp_w = sp.split()   # already normalised by _extract_sgml_text
        if not sp_w:
            continue
        # Quick length filter: skip if word counts differ by more than 50 %
        len_ratio = min(len(pdf_words), len(sp_w)) / max(len(pdf_words), len(sp_w), 1)
        if len_ratio < 0.40:
            continue
        r = SequenceMatcher(None, pdf_words, sp_w).ratio()
        if r > best_ratio:
            best_ratio = r
            best_sgml = sp
            if best_ratio >= 0.998:
                break  # effectively identical — stop early

    # Mutation zone: paragraph IS present (≥ 0.68) but NOT word-for-word identical (< 0.998)
    # Upper bound raised to 0.998 so single-word deletions/substitutions (ratio ≈ 0.95–0.98)
    # are caught rather than silently passing as "close enough".
    if best_ratio < 0.68 or best_ratio >= 0.998:
        return None

    pdf_w = pdf_words
    sgml_w = best_sgml.split()
    sm = SequenceMatcher(None, pdf_w, sgml_w)
    deleted: list[str] = []
    inserted: list[str] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag in ("delete", "replace"):
            deleted.extend(pdf_w[i1:i2])
        if tag in ("insert", "replace"):
            inserted.extend(sgml_w[j1:j2])

    # Filter noise tokens: length-1 chars, pure punctuation, and PDF artifacts.
    # We require >=2 meaningful alpha-content words (>=3 chars) to reduce false
    # positives from ligature rendering, hyphenation splits, punctuation style
    # differences, and number formatting (all common in PDF→SGML extraction).
    def _sig(tokens: list[str]) -> list[str]:
        return [w for w in tokens
                if len(w) >= 2
                and re.search(r"[a-zA-Z]{2,}", w)  # must contain >=2 alpha chars
                and not re.fullmatch(r"[^a-zA-Z0-9]+", w)]  # not pure punctuation

    meaningful_del = _sig(deleted)
    meaningful_ins = _sig(inserted)
    # Require at least 2 meaningful words on either side to suppress PDF-artifact
    # false positives (ligatures, hyphenation, quote-style differences).
    if len(meaningful_del) < 2 and len(meaningful_ins) < 2:
        return None

    # Suppress structural-split false positives: when the PDF paragraph spans
    # multiple SGML elements (e.g. heading + body merged by PDF extractor),
    # the best-match SGML paragraph will show many 'deleted' words with nothing
    # 'inserted' (or vice versa). This is not a real word mutation.
    # Pattern: one side empty AND the other side has >=3 content words.
    # EXCEPTION: when ratio >= 0.90 the paragraph is a strong match — the
    # deletion is a real targeted edit (e.g. 3 words removed from a 50-word
    # paragraph), not a structural PDF merge artifact (which produces lower ratios).
    if best_ratio < 0.90:
        if (not meaningful_del and len(meaningful_ins) >= 3) or \
           (not meaningful_ins and len(meaningful_del) >= 3):
            return None

    # Suppress word-reorder artifacts: when del and ins contain mostly the
    # same words (just repositioned), it is not a real mutation.
    # Catches PDF two-column merges and paragraph-split recombinations.
    if meaningful_del and meaningful_ins:
        _del_set = set(meaningful_del)
        _ins_set = set(meaningful_ins)
        _overlap = len(_del_set & _ins_set)
        _min_side = min(len(_del_set), len(_ins_set))
        if _min_side > 0 and _overlap / _min_side >= 0.70:
            return None  # same words, different order → structural artifact

    # Suppress hyphenation line-break artifacts: PDF sometimes splits a
    # hyphenated compound across lines as  ['word-', 'part'] while SGML has
    # ['word-part'].  Detect: del is 2 tokens where first ends with '-' and
    # their concatenation equals the single ins token (or vice-versa).
    if len(meaningful_del) <= 2 and len(meaningful_ins) == 1:
        joined = ''.join(meaningful_del)
        if joined == meaningful_ins[0]:
            return None
    if len(meaningful_ins) <= 2 and len(meaningful_del) == 1:
        joined = ''.join(meaningful_ins)
        if joined == meaningful_del[0]:
            return None

    return {
        "pdf_text": pdf_para[:250],
        "sgml_text": " ".join(sgml_w)[:250],
        "ratio": round(best_ratio, 3),
        "deleted_words": meaningful_del[:20],
        "inserted_words": meaningful_ins[:20],
    }


# ─────────────────────────────────────────────────────────────────────────────
# D3: Text accuracy
# ─────────────────────────────────────────────────────────────────────────────
def check_text_accuracy(pdf: _PDFData, sgml_data: dict, result: L4Result,
                        docx_data: "dict | None" = None) -> None:
    """
    D3 — 8 pts: Paragraph-level text diff, number integrity, citation integrity.

    When docx_data is provided (GAP 1): two-stage comparison
      Stage 1 – PDF vs DOCX  → ABBYY extraction errors (informational, not scored)
      Stage 2 – DOCX vs SGML → pipeline conversion errors (scored)
    Fallback (no DOCX): original single-stage PDF vs SGML comparison.
    """
    score = 8.0
    sgml_blob = sgml_data["text"].lower()
    sgml_paragraphs = sgml_data["paragraphs"]
    raw_sgml = sgml_data.get("_raw_sgml", "")
    # Pre-compute once — expensive on large files so share across D3-d/D3-f
    _raw_sgml_norm = _norm(raw_sgml)              # tag-stripped, normalised
    _raw_sgml_norm_words = set(_raw_sgml_norm.split())  # word set for membership checks

    # Amending instruments intentionally reproduce only the changed sections, not
    # the full source document. Low paragraph coverage relative to the complete
    # source PDF is therefore by design. We apply a more lenient minimum score
    # tier for such documents.
    # Detection: <QUOTE> tag (classic form), OR "Amending Instrument" / "Amendment
    # Regulation" phrase in title text (some files omit QUOTE wrapping).
    is_amending_doc = bool(
        re.search(r"<QUOTE[\s>]", raw_sgml) or
        re.search(
            r"\bAmend(?:ing|ment(?:ary)?)\s+(?:Instrument|Regulation|Rule|Order)\b",
            raw_sgml, re.IGNORECASE,
        ) or
        # <N> or <TI> tag contains "Amendment" / "Amending" — covers formats
        # like <N>11-803 (Amendment)</N> and <TI>Amendment Regulations</TI>
        re.search(r"<(?:N|TI)[^>]*>[^<]*\bAmend(?:ing|ment(?:ary)?)\b", raw_sgml, re.IGNORECASE)
    )

    if not pdf.paragraphs:
        result.text_score = 8.0
        result.warnings.append("D3 skipped: no paragraphs extracted from PDF (scanned or encrypted).")
        return

    # Single-character bullet tokens that appear in PDF but not SGML
    _BULLET_TOKENS: frozenset[str] = frozenset({
        "o", "•", "◦", "▪", "▸", "→", "–", "-", ";", ",",
        "○", "●", "■", "□", "\uf0b7",  # additional unicode bullet variants
    })

    # Build SGML n-gram set (shared by both single-stage and two-stage paths)
    ngram_size = 5
    sgml_words = sgml_blob.split()
    sgml_ngrams: set[tuple] = set()
    for i in range(len(sgml_words) - ngram_size + 1):
        sgml_ngrams.add(tuple(sgml_words[i:i + ngram_size]))

    # ── D3-a: Paragraph coverage ──────────────────────────────────────────────
    # docx_data is pre-parsed by validate_source_comparison and passed in
    # (GAP 4 refactor: parse DOCX once, share with D2 and D3).

    if docx_data is not None:
        # ── Two-stage path (GAP 1) ────────────────────────────────────────────
        result.docx_available = True
        docx_blob = docx_data["combined_text"]
        docx_words_list = docx_blob.split()
        docx_ngrams: set[tuple] = set()
        for i in range(len(docx_words_list) - ngram_size + 1):
            docx_ngrams.add(tuple(docx_words_list[i:i + ngram_size]))

        # Stage 1: PDF vs DOCX — what ABBYY missed (informational only, not scored)
        meaningful_pdf = [p for p in pdf.paragraphs if len(p.split()) >= 5 and not _is_omittable(p)]
        abbyy_missing: list[str] = []
        abbyy_missing_details: list[dict] = []
        for para in meaningful_pdf:
            words = [w for w in _norm(para).split() if w not in _BULLET_TOKENS]
            is_cov, _conf, _meth = _para_covered_v2(words, docx_blob, docx_ngrams)
            if not is_cov:
                abbyy_missing.append(para[:100])
                abbyy_missing_details.append({'text': para[:100], 'confidence': _conf, 'method': _meth})
        result.abbyy_missing_paragraphs = abbyy_missing
        result.abbyy_missing_paragraph_details = abbyy_missing_details
        if abbyy_missing:
            result.warnings.append(
                f"D3-ABBYY — {len(abbyy_missing)} paragraph(s) from PDF not captured in "
                f"DOCX (ABBYY extraction gaps — cannot be fixed in SGML editor): "
                f"{[p[:60] for p in abbyy_missing[:2]]}"
            )

        # Stage 2: DOCX vs SGML — what the pipeline missed (scored)
        # Threshold lowered: 5 words min (was 8) so short but meaningful lines are checked.
        meaningful_docx = [p for p in docx_data["paragraphs"] if len(p.split()) >= 5]
        pipeline_missing: list[str] = []
        pipeline_missing_details: list[dict] = []
        pipeline_truncated: list[str] = []   # D3-d: present but leading text deleted
        pipeline_mutations: list[dict] = []  # D3-e: present but with word changes

        # ── LLM-augmented paragraph alignment (DOCX→SGML) ────────────────────
        # When LLM is available: use Opus to align every meaningful DOCX paragraph
        # to its best SGML match, detect truncated starts and inline mutations.
        # This replaces the fragile ngram-best-match + heuristic D3-d/D3-e logic.
        _llm_results_docx: "list[dict] | None" = None
        if _LLM_ENABLED:
            _llm_results_docx = _llm_align_paragraphs(
                meaningful_docx[:_LLM_MAX_PARAS],
                sgml_paragraphs[:int(_LLM_MAX_PARAS * 1.5)],
            )

        if _llm_results_docx is not None:
            # LLM path: trust Opus for alignment, truncation, and mutation detection.
            # LLM false-positive guard: when LLM says "not found", cross-check with
            # the deterministic _para_covered_v2. Vendor SGMLs are 99% accurate so if
            # deterministic confirms coverage, trust it over the LLM.
            for item in _llm_results_docx:
                _pidx = item.get("pdf_idx", 1) - 1   # 0-based
                if _pidx < 0 or _pidx >= len(meaningful_docx):
                    continue
                para = meaningful_docx[_pidx]
                if not item.get("found", True):
                    # Deterministic cross-check before accepting LLM "not found"
                    _xw = [w for w in _norm(para).split() if w]
                    _x_cov, _x_conf, _ = _para_covered_v2(_xw, sgml_blob, sgml_ngrams)
                    if _x_cov:
                        covered += 1  # deterministic says present → trust it, skip
                    else:
                        pipeline_missing.append(para[:100])
                        pipeline_missing_details.append({
                            "text": para[:100], "confidence": item.get("confidence", 0.0),
                            "method": "llm"
                        })
                else:
                    # D3-d: Truncated start OR end detected by LLM
                    # Only count if ≥5 substantive words are missing (filter out list
                    # markers, section numbers, and short labels which are FPs).
                    _del_start = item.get("deleted_start")
                    _del_end   = item.get("deleted_end")
                    _del_start_words = len(str(_del_start).split()) if (_del_start and str(_del_start).strip()) else 0
                    _del_end_words   = len(str(_del_end).split())   if (_del_end   and str(_del_end).strip())   else 0
                    # Word-coverage check: if ≥65% of content words (>3 chars) from
                    # the deleted text appear in the SGML blob, the text IS present
                    # elsewhere (structurally encoded as a heading/title), not truly
                    # deleted. More robust than exact n-gram matching against
                    # pluralization differences ("identifier" vs "identifiers").
                    def _word_cov(phrase, blob):
                        toks = [w.lower().rstrip('.,;:()[]"\'\'').lstrip('(["\'\'') for w in str(phrase).split() if len(w) > 3]
                        if not toks:
                            return True  # no content words → assume present
                        found = sum(1 for w in toks if w in blob)
                        return found / len(toks) >= 0.65
                    _del_start_in_sgml = _del_start_words >= 5 and _word_cov(_del_start, sgml_blob)
                    _del_end_in_sgml   = _del_end_words   >= 5 and _word_cov(_del_end,   sgml_blob)
                    if (_del_start_words >= 5 and not _del_start_in_sgml) or \
                       (_del_end_words >= 5 and not _del_end_in_sgml):
                        pipeline_truncated.append(para[:120])
                        result.llm_confirmed_truncations += 1  # escalation-eligible
                    # D3-e: Inline mutations detected by LLM
                    # Handle two Opus response formats:
                    #   Format A: {deleted: [...], inserted: [...]}
                    #   Format B: {type: "deletion"/"insertion", text: "..."}
                    for _mut in item.get("mutations", []):
                        _del = _mut.get("deleted") or []
                        _ins = _mut.get("inserted") or []
                        # Format B fallback
                        if not _del and not _ins and _mut.get("text"):
                            _mut_type = _mut.get("type", "")
                            _mut_words = str(_mut.get("text", "")).split()[:20]
                            if "del" in _mut_type:
                                _del = _mut_words
                            elif "ins" in _mut_type:
                                _ins = _mut_words
                        _del_list = (list(_del) if isinstance(_del, list) else str(_del).split())[:20]
                        _ins_list = (list(_ins) if isinstance(_ins, list) else str(_ins).split())[:20]
                        # Word-coverage check: if ≥65% of content words (>3 chars)
                        # from the deleted list appear in the SGML, it's a paragraph
                        # misalignment FP, not a real deletion.
                        _del_content_toks = [w.lower().rstrip('.,;:') for w in _del_list[:10] if len(w) > 3]
                        if _del_content_toks:
                            _del_in_blob_ratio = sum(1 for w in _del_content_toks if w in sgml_blob) / len(_del_content_toks)
                            _del_words_in_sgml = _del_in_blob_ratio >= 0.65
                        else:
                            _del_words_in_sgml = True  # no content words → skip
                        if len(_del_list) >= 3 and not _del_words_in_sgml:
                            pipeline_mutations.append({
                                "pdf_text": para[:250],
                                "sgml_text": "",
                                "ratio": item.get("confidence", 0.0),
                                "deleted_words": _del_list,
                                "inserted_words": _ins_list,
                            })
                            result.llm_confirmed_mutations += 1  # escalation-eligible
            # Also check DOCX paras beyond the LLM batch with deterministic fallback
            for para in meaningful_docx[_LLM_MAX_PARAS:]:
                words = [w for w in _norm(para).split() if w not in _BULLET_TOKENS]
                is_cov, _conf, _meth = _para_covered_v2(words, sgml_blob, sgml_ngrams)
                if not is_cov:
                    pipeline_missing.append(para[:100])
                    pipeline_missing_details.append({"text": para[:100], "confidence": _conf, "method": _meth})
        else:
            # Deterministic fallback (LLM unavailable)
            for para in meaningful_docx:
                words = [w for w in _norm(para).split() if w not in _BULLET_TOKENS]
                is_cov, _conf, _meth = _para_covered_v2(words, sgml_blob, sgml_ngrams)
                if not is_cov:
                    pipeline_missing.append(para[:100])
                    pipeline_missing_details.append({'text': para[:100], 'confidence': _conf, 'method': _meth})
                else:
                    _clean = [re.sub(r'[;,.:]+$', '', w).strip('"\' ') for w in words[:6] if w]
                    if len(_clean) >= 4:
                        _first = _clean[0]
                        _is_struct_prefix = (
                            re.match(r'^\d+$', _first)
                            or re.match(r'^\([a-z]{1,4}\)$', _first)
                            or re.match(r'^\(', _first)
                            or (len(_first) > 0 and ord(_first[0]) >= 0xE000)
                            or (len(_first) > 0 and 0x2700 <= ord(_first[0]) <= 0x27BF)
                        )
                        if not _is_struct_prefix:
                            _prefix = ' '.join(_clean[:5])
                            _words_missing = sum(1 for w in _clean[:5] if w not in _raw_sgml_norm_words)
                            if _prefix and _prefix not in sgml_blob and _words_missing >= 2:
                                pipeline_truncated.append(para[:120])
                    if len(words) >= 8:   # deterministic gate: 8 words to reduce FPs
                        _mut = _word_diff_result(para, sgml_paragraphs)
                        if _mut:
                            pipeline_mutations.append(_mut)

        result.pipeline_missing_paragraphs = pipeline_missing
        result.pipeline_missing_paragraph_details = pipeline_missing_details
        result.missing_paragraphs = pipeline_missing   # backward compat
        result.truncated_paragraphs = result.truncated_paragraphs + pipeline_truncated
        result.inline_changed_paragraphs = pipeline_mutations

        if not meaningful_docx:
            result.text_score = 8.0
            return

        # D3-f: Short line check — lines 3-4 words in DOCX not found at all in SGML.
        # Use pre-computed _raw_sgml_norm (tag-stripped, normalised) for tag-safe search.
        _short_docx = [
            p for p in docx_data["paragraphs"]
            if 3 <= len(p.split()) <= 4
            and not _is_omittable(p)
            and not re.match(r'^\d+\s*[-\u2013]\s*\S', p)  # exclude TOC entries
            and not re.search(r'\s\d{1,3}$', p)           # exclude page-number lines
            and not re.search(r'(?<!\s)\d{1,3}$', p)      # exclude footnote-ref lines (e.g. 'text.48')
        ]
        _missing_short: list[str] = []
        for _sl in _short_docx[:80]:
            _nsl = _norm(_sl).rstrip(';:,.')
            if _nsl and _nsl not in sgml_blob and _nsl not in _raw_sgml_norm:
                _missing_short.append(_sl[:80])
        result.missing_short_lines = _missing_short

        coverage = 1.0 - (len(pipeline_missing) / len(meaningful_docx))
        result.text_coverage = coverage

    else:
        # ── Single-stage path (original PDF→SGML) ────────────────────────────
        # Threshold lowered: 5 words min (was 8) so short but meaningful lines are checked.
        meaningful = [p for p in pdf.paragraphs if len(p.split()) >= 5 and not _is_omittable(p)]
        if not meaningful:
            result.text_score = 8.0
            return

        sampled = meaningful
        covered = 0
        missing: list[str] = []
        truncated: list[str] = []   # D3-d: paragraphs present but with leading text deleted
        mutations: list[dict] = []  # D3-e: paragraphs present but with word changes

        # ── LLM-augmented paragraph alignment (PDF→SGML) ─────────────────────
        # When LLM is available: use Opus to align every meaningful PDF paragraph
        # to its best SGML match, detect truncated starts and inline mutations.
        _llm_results_pdf: "list[dict] | None" = None
        if _LLM_ENABLED:
            _llm_results_pdf = _llm_align_paragraphs(
                sampled[:_LLM_MAX_PARAS],
                sgml_paragraphs[:int(_LLM_MAX_PARAS * 1.5)],
            )

        if _llm_results_pdf is not None:
            # LLM path
            # LLM false-positive guard: when LLM says "not found", cross-check with
            # _para_covered_v2. Vendor SGMLs are 99% accurate — if deterministic
            # confirms coverage, trust it and don't flag as missing.
            for item in _llm_results_pdf:
                _pidx = item.get("pdf_idx", 1) - 1
                if _pidx < 0 or _pidx >= len(sampled):
                    continue
                para = sampled[_pidx]
                if not item.get("found", True):
                    # Cross-check with deterministic before accepting LLM verdict
                    _xw = [w for w in _norm(para).split() if w]
                    _x_cov, _x_conf, _ = _para_covered_v2(_xw, sgml_blob, sgml_ngrams)
                    if _x_cov:
                        covered += 1  # deterministic overrides LLM false positive
                    else:
                        # D3-g partial presence: first 8 words in SGML but last 8 not
                        # → paragraph was truncated (not fully deleted). Add to both
                        # missing AND truncated so D3-d penalty fires even in large docs.
                        _pp_w = [w for w in _xw if w not in _BULLET_TOKENS]
                        if len(_pp_w) >= 15:
                            _head_str = ' '.join(_pp_w[:8])
                            _tail_str = ' '.join(_pp_w[-8:])
                            _tail_cw = [w for w in _pp_w[-8:] if len(w) > 3]
                            if len(_tail_cw) >= 3 and _head_str in sgml_blob and _tail_str not in sgml_blob:
                                if para[:120] not in truncated:
                                    truncated.append(para[:120])
                                    result.llm_confirmed_truncations += 1
                        missing.append(para[:100])
                else:
                    covered += 1
                    # D3-d: Truncated start OR end detected by LLM
                    # Only count if ≥5 substantive words are missing (filter out list
                    # markers, section numbers, and short labels which are FPs).
                    _del_start = item.get("deleted_start")
                    _del_end   = item.get("deleted_end")
                    _del_start_words = len(str(_del_start).split()) if (_del_start and str(_del_start).strip()) else 0
                    _del_end_words   = len(str(_del_end).split())   if (_del_end   and str(_del_end).strip())   else 0
                    # Word-coverage check: if ≥65% of content words (>3 chars) from
                    # the deleted text appear in the SGML blob, the text IS present
                    # elsewhere (structurally encoded as a heading/title), not truly
                    # deleted. More robust than exact n-gram matching against
                    # pluralization/form differences ("identifier" vs "identifiers").
                    def _word_cov(phrase, blob):
                        toks = [w.lower().rstrip('.,;:()[]"\'\'').lstrip('(["\'\'') for w in str(phrase).split() if len(w) > 3]
                        if not toks:
                            return True  # no content words → assume present
                        found = sum(1 for w in toks if w in blob)
                        return found / len(toks) >= 0.65
                    _del_start_in_sgml = _del_start_words >= 5 and _word_cov(_del_start, sgml_blob)
                    _del_end_in_sgml   = _del_end_words   >= 5 and _word_cov(_del_end,   sgml_blob)
                    if (_del_start_words >= 5 and not _del_start_in_sgml) or \
                       (_del_end_words >= 5 and not _del_end_in_sgml):
                        truncated.append(para[:120])
                        result.llm_confirmed_truncations += 1  # escalation-eligible
                    # D3-e: Inline mutations — handle both Opus response formats
                    for _mut in item.get("mutations", []):
                        _del = _mut.get("deleted") or []
                        _ins = _mut.get("inserted") or []
                        if not _del and not _ins and _mut.get("text"):
                            _mut_type = _mut.get("type", "")
                            _mut_words = str(_mut.get("text", "")).split()[:20]
                            if "del" in _mut_type:
                                _del = _mut_words
                            elif "ins" in _mut_type:
                                _ins = _mut_words
                        _del_list = (list(_del) if isinstance(_del, list) else str(_del).split())[:20]
                        _ins_list = (list(_ins) if isinstance(_ins, list) else str(_ins).split())[:20]
                        # Word-coverage check: if ≥65% of content words (>3 chars)
                        # from the deleted list appear in the SGML, it's a paragraph
                        # misalignment FP, not a real deletion.
                        _del_content_toks = [w.lower().rstrip('.,;:') for w in _del_list[:10] if len(w) > 3]
                        if _del_content_toks:
                            _del_in_blob_ratio = sum(1 for w in _del_content_toks if w in sgml_blob) / len(_del_content_toks)
                            _del_words_in_sgml = _del_in_blob_ratio >= 0.65
                        else:
                            _del_words_in_sgml = True  # no content words → skip
                        if len(_del_list) >= 3 and not _del_words_in_sgml:
                            mutations.append({
                                "pdf_text": para[:250],
                                "sgml_text": "",
                                "ratio": item.get("confidence", 0.0),
                                "deleted_words": _del_list,
                                "inserted_words": _ins_list,
                            })
                            result.llm_confirmed_mutations += 1  # escalation-eligible
                    # D3-g: Word-count ratio check — catches trailing truncation the LLM misses.
                    # Uses the LLM-identified best_sgml_idx to find the matched SGML paragraph
                    # and compares its word count to the PDF paragraph word count.
                    # If matched SGML paragraph has <55% of PDF paragraph words, it was truncated.
                    # This is immune to repeated phrases (tail 8-word substring check fails on
                    # large legal docs where boilerplate appears in multiple paragraphs).
                    if not item.get("_error", False):
                        _p_w = [w for w in _norm(para).split() if w not in _BULLET_TOKENS]
                        _best_idx = item.get("best_sgml_idx", None)
                        if len(_p_w) >= 20 and _best_idx and isinstance(_best_idx, int):
                            _bi = _best_idx - 1  # 0-based
                            if 0 <= _bi < len(sgml_paragraphs):
                                _matched_wc = len(_norm(sgml_paragraphs[_bi]).split())
                                _pdf_wc = len(_p_w)
                                if _matched_wc > 0 and (_matched_wc / _pdf_wc) < 0.55:
                                    # D3-g word-count ratio: confirm truncation by checking that
                                    # the tail words of the PDF paragraph are absent from the
                                    # full ligature-normalised SGML blob.  In a split-paragraph
                                    # scenario the tail IS in the blob (in a different para).
                                    # Build ligature-normalised blob matching _norm() output.
                                    _nb = sgml_blob
                                    for _lc, _lp in _LIGATURE_MAP.items():
                                        _nb = _nb.replace(_lc, _lp)
                                    # Clean tail words: strip trailing punctuation AND superscript
                                    # digits (footnote refs like "act5." → "act") so they match
                                    # the SGML blob which has no punctuation or footnote markers.
                                    _tail_raw = _p_w[-8:] if len(_p_w) >= 8 else _p_w
                                    _tail_clean = [
                                        re.sub(r'[\d]+$', '', w.rstrip('.,;:!?)"\' '))
                                        for w in _tail_raw
                                    ]
                                    _content_tail = [w for w in _tail_clean if len(w) > 3]
                                    if _content_tail:
                                        _found = sum(1 for w in _content_tail if w in _nb)
                                        # Only flag if majority of content tail absent from blob
                                        _tail_absent = (_found / len(_content_tail)) < 0.50
                                    else:
                                        _tail_absent = False
                                    if _tail_absent:
                                        import os as _os
                                        if _os.environ.get('SV_DEBUG_RATIO'):
                                            print(f"[RATIO] ratio={_matched_wc/_pdf_wc:.3f} pdf_wc={_pdf_wc} sgml_wc={_matched_wc} idx={_best_idx}")
                                            print(f"  PDF : {para[:120]!r}")
                                            print(f"  SGML: {sgml_paragraphs[_bi][:100]!r}")
                                            print(f"  tail_clean={_tail_clean}  found={_found}/{len(_content_tail)}")
                                        if para[:120] not in truncated:
                                            truncated.append(para[:120])
                                            result.llm_confirmed_truncations += 1
                        # Tail 8-word fallback: for paragraphs where best_sgml_idx is absent
                        elif len(_p_w) >= 15:
                            _tail_str = ' '.join(_p_w[-8:])
                            _content_w = [w for w in _p_w[-8:] if len(w) > 3]
                            if len(_content_w) >= 3 and _tail_str not in sgml_blob:
                                if para[:120] not in truncated:
                                    truncated.append(para[:120])
                                    result.llm_confirmed_truncations += 1
            # Deterministic fallback for paras beyond LLM batch limit
            for para in sampled[_LLM_MAX_PARAS:]:
                words = [w for w in _norm(para).split() if w not in _BULLET_TOKENS]
                is_cov, _conf, _meth = _para_covered_v2(words, sgml_blob, sgml_ngrams)
                if is_cov:
                    covered += 1
                else:
                    # D3-g partial presence: head in SGML, tail not → truncation
                    if len(words) >= 15:
                        _hs = ' '.join(words[:8])
                        _ts = ' '.join(words[-8:])
                        _tcw = [w for w in words[-8:] if len(w) > 3]
                        if len(_tcw) >= 3 and _hs in sgml_blob and _ts not in sgml_blob:
                            if para[:120] not in truncated:
                                truncated.append(para[:120])
                    missing.append(para[:100])
        else:
            # Deterministic fallback (LLM unavailable)
            for para in sampled:
                words = [w for w in _norm(para).split() if w not in _BULLET_TOKENS]
                is_cov, _conf, _meth = _para_covered_v2(words, sgml_blob, sgml_ngrams)
                if is_cov:
                    covered += 1
                    # D3-d: Verify the paragraph's opening words are also present.
                    _clean = [re.sub(r'[;,.:]+$', '', w).strip('"\' ') for w in words[:6] if w]
                    if len(_clean) >= 4:
                        _first = _clean[0]
                        _is_struct_prefix = (
                            re.match(r'^\d+$', _first)
                            or re.match(r'^\([a-z]{1,4}\)$', _first)
                            or re.match(r'^\(', _first)
                            or (len(_first) > 0 and ord(_first[0]) >= 0xE000)
                            or (len(_first) > 0 and 0x2700 <= ord(_first[0]) <= 0x27BF)
                        )
                        if not _is_struct_prefix:
                            _prefix = ' '.join(_clean[:5])
                            _words_missing = sum(1 for w in _clean[:5] if w not in _raw_sgml_norm_words)
                            if _prefix and _prefix not in sgml_blob and _words_missing >= 2:
                                truncated.append(para[:120])
                    # D3-g: Tail 8-word check — catches trailing truncation
                    if len(words) >= 15:
                        _tail_str = ' '.join(words[-8:])
                        _content_w = [w for w in words[-8:] if len(w) > 3]
                        if len(_content_w) >= 3 and _tail_str not in sgml_blob:
                            if para[:120] not in truncated:
                                truncated.append(para[:120])
                    if len(words) >= 8:   # deterministic gate: 8 words to reduce FPs
                        _mut = _word_diff_result(para, sgml_paragraphs)
                        if _mut:
                            # Suppress heading-prefix artifacts: PDF page 1 often merges
                            # the document title (LABEL+N+TI) with the intro paragraph as
                            # one text block → appears as deleted title words. If ≥75% of
                            # the 'deleted' words are in the SGML heading vocabulary, it's
                            # a PDF extraction artifact, not a real content mutation.
                            _hdg_vocab: set[str] = set()
                            for _h in sgml_data.get("headings", []):
                                _hdg_vocab.update(_h.split())
                            for _s in sgml_data.get("sections", []):
                                _hdg_vocab.update(_s.split())
                            for _av in sgml_data.get("attrs", {}).values():
                                _hdg_vocab.update(_norm(str(_av)).split())
                            _del_w = [w.lower() for w in _mut.get("deleted_words", [])]
                            if _del_w and sum(1 for w in _del_w if w in _hdg_vocab) / len(_del_w) >= 0.75:
                                _mut = None  # suppress: deleted words = document title/heading
                        if _mut:
                            mutations.append(_mut)
                else:
                    # D3-g partial presence (deterministic no-LLM path): head present,
                    # tail absent → trailing truncation rather than full deletion.
                    if len(words) >= 15:
                        _hs = ' '.join(words[:8])
                        _ts = ' '.join(words[-8:])
                        _tcw = [w for w in words[-8:] if len(w) > 3]
                        if len(_tcw) >= 3 and _hs in sgml_blob and _ts not in sgml_blob:
                            if para[:120] not in truncated:
                                truncated.append(para[:120])
                    missing.append(para[:100])

        # D3-f: Short line check — 3-4 word lines from PDF raw_lines not found in SGML.
        # Use pre-computed _raw_sgml_norm (tag-stripped, normalised) for tag-safe search.
        # Also exclude TOC / page-number style lines.
        _short_raw = [
            ln for ln in pdf.raw_lines
            if 3 <= len(ln.split()) <= 4
            and not _is_omittable(ln)
            and not re.match(r'^\d+\s*[-\u2013]\s*\S', ln)  # exclude "1 - General" TOC
            and not re.search(r'\s\d{1,3}$', ln)             # exclude page-number lines
            and not re.search(r'(?<!\s)\d{1,3}$', ln)        # exclude footnote-ref lines
        ]
        # Deterministic pre-filter: lines that fail basic substring check.
        # Also check a "tightened" form of sgml_blob where spaces that were
        # introduced by SGML tag stripping (e.g. '<BOLD>Orders</BOLD>)' becomes
        # 'Orders )') are collapsed back against the bracket/colon so that
        # '(collectively, the Blanket Orders)' still matches correctly.
        _sgml_blob_tight = sgml_blob.replace(' )', ')').replace(' :', ':').replace(' ,', ',')
        _raw_sgml_tight  = _raw_sgml_norm.replace(' )', ')').replace(' :', ':').replace(' ,', ',')
        _short_candidates = []
        for _sl in _short_raw[:80]:
            _nsl = _norm(_sl).rstrip(';:,.')
            if _nsl and _nsl not in sgml_blob and _nsl not in _raw_sgml_norm \
                     and _nsl not in _sgml_blob_tight and _nsl not in _raw_sgml_tight:
                _short_candidates.append(_sl[:80])

        # LLM verification of short-line candidates (removes FPs where words exist
        # in SGML but in a different element — e.g. heading vs body context).
        # With unlimited budget: verify ALL candidates with Opus when available.
        if _LLM_ENABLED and _short_candidates:
            _llm_sl = _llm_align_paragraphs(
                _short_candidates,
                sgml_paragraphs[:int(_LLM_MAX_PARAS * 1.5)],
            )
            if _llm_sl is not None:
                _missing_short_raw = [
                    _short_candidates[item.get("pdf_idx", 1) - 1]
                    for item in _llm_sl
                    if not item.get("found", True)
                    and 0 <= item.get("pdf_idx", 1) - 1 < len(_short_candidates)
                ]
            else:
                _missing_short_raw = _short_candidates
        else:
            _missing_short_raw = _short_candidates

        result.missing_short_lines = _missing_short_raw

        coverage = covered / len(sampled)
        result.text_coverage = coverage
        result.missing_paragraphs = missing
        result.truncated_paragraphs = truncated
        result.inline_changed_paragraphs = mutations

    # ── Scoring (same tiers for both paths) ──────────────────────────────────
    missing_for_msg = result.missing_paragraphs  # pipeline_missing or missing
    if coverage >= 0.92:
        text_sub_score = 5.0
    elif coverage >= 0.80:
        text_sub_score = 4.0
        _add_issue(result, "text_accuracy", "minor",
                   f"D3 — Paragraph coverage {coverage:.0%}. "
                   f"{len(missing_for_msg)} paragraph(s) from "
                   f"{'DOCX' if docx_data else 'PDF'} not found in SGML.",
                   impact="-1 pt")
    elif coverage >= 0.65:
        text_sub_score = 3.0
        _add_issue(result, "text_accuracy", "major",
                   f"D3 — Paragraph coverage {coverage:.0%}. "
                   f"{len(missing_for_msg)} paragraph(s) missing from SGML. "
                   f"Examples: {missing_for_msg[:2]}",
                   impact="-2 pts")
    else:
        # For amending instruments (<QUOTE> present) low coverage is expected —
        # the SGML only contains the changed sections, not the full source PDF.
        # Floor the sub-score at 3.0 (same as 65-80% tier) instead of 1.0.
        text_sub_score = 3.0 if is_amending_doc else 1.0
        severity = "major" if is_amending_doc else "critical"
        _add_issue(result, "text_accuracy", severity,
                   f"D3 — {'Low' if not is_amending_doc else 'Partial'} paragraph coverage: "
                   f"{coverage:.0%}. "
                   + ("Amending instrument — only changed sections are expected in SGML."
                      if is_amending_doc else
                      "Significant content may be missing from SGML."),
                   impact=f"-{5.0 - text_sub_score:.0f} pts")

    score = score - (5.0 - text_sub_score)

    # D3-d: Truncated paragraph check — paragraph body present but leading text deleted.
    # This catches cases where a few words at the START of a paragraph are removed;
    # the bulk of the paragraph still passes the coverage check so goes undetected
    # without this explicit prefix verification.
    _truncated = result.truncated_paragraphs
    if _truncated:
        _trunc_pts = min(2.0, len(_truncated) * 0.5)
        score -= _trunc_pts
        _add_issue(result, "text_accuracy", "major",
                   f"D3 — {len(_truncated)} paragraph(s) found in SGML but with leading "
                   f"text missing (text was deleted from the start of a paragraph). "
                   f"Examples: {[t[:80] for t in _truncated[:3]]}",
                   impact=f"-{_trunc_pts:.1f} pts")

    # D3-e: Inline word mutation check — paragraphs present but with words changed.
    # Catches mid-sentence additions, deletions, and word replacements that pass
    # the fuzzy coverage check (87.5% threshold) but have altered wording.
    _mutations = result.inline_changed_paragraphs
    if _mutations:
        _mut_pts = min(2.0, len(_mutations) * 0.4)
        score -= _mut_pts
        _example_muts = [
            f"'{m['pdf_text'][:60]}' → deleted: {m['deleted_words'][:5]}"
            for m in _mutations[:2]
        ]
        _add_issue(result, "text_accuracy", "major",
                   f"D3 — {len(_mutations)} paragraph(s) have inline word changes "
                   f"(words added, deleted or replaced vs PDF source). "
                   f"Examples: {_example_muts}",
                   impact=f"-{_mut_pts:.1f} pts")

    # D3-f: Short line check — 3-4 word lines from PDF/DOCX not found anywhere in SGML.
    _short_missing = result.missing_short_lines
    if _short_missing:
        _short_pts = min(1.0, len(_short_missing) * 0.25)
        score -= _short_pts
        _add_issue(result, "text_accuracy", "minor",
                   f"D3 — {len(_short_missing)} short line(s) (3–4 words) from PDF/DOCX "
                   f"not found in SGML. May be contact lines, date lines, or label lines "
                   f"that were deleted. Examples: {_short_missing[:3]}",
                   impact=f"-{_short_pts:.1f} pts")

    # D3-b: Number integrity — check numbers from PDF first page appear in SGML
    pdf_numbers = set(_NUMBER_RE.findall(pdf.first_page_text[:2000]))
    missing_numbers = []
    for num in list(pdf_numbers)[:20]:
        norm_num = num.replace(",", "").replace("$", "").replace("%", "")
        if norm_num not in sgml_blob and num not in sgml_blob:
            missing_numbers.append(num)

    if len(missing_numbers) > 3:
        pts = min(1.5, len(missing_numbers) * 0.1)
        score -= pts
        _add_issue(result, "text_accuracy", "major",
                   f"D3 — {len(missing_numbers)} number(s) from PDF first page not found in SGML: "
                   f"{missing_numbers[:5]}. Numbers may be mis-keyed.",
                   impact=f"-{pts:.1f} pts")

    # D3-c: Legal citation integrity
    pdf_citations = set(_LEGAL_CITATION_RE.findall(pdf.first_page_text[:3000]))
    missing_citations = [c for c in pdf_citations if _norm(c) not in sgml_blob]
    if missing_citations:
        pts = min(1.5, len(missing_citations) * 0.25)
        score -= pts
        _add_issue(result, "text_accuracy", "major",
                   f"D3 — Legal citation(s) from PDF not found in SGML: {missing_citations[:3]}. "
                   f"Citations may be corrupted or missing.",
                   impact=f"-{pts:.1f} pts")

    result.text_score = max(0.0, score)


# ─────────────────────────────────────────────────────────────────────────────
# D4: Completeness
# ─────────────────────────────────────────────────────────────────────────────
def check_completeness(pdf: _PDFData, sgml_data: dict, result: L4Result,
                       docx_data: "dict | None" = None) -> None:
    """
    D4 — 7 pts: Count-based completeness check + cell-level table comparison.

    Tables, images, footnotes, sections, pages.
    GAP 5: When docx_data is provided, adds cell-by-cell table content comparison
    (DOCX table cells vs SGML TBLCELL elements) to catch dropped or corrupted rows.
    """
    score = 7.0

    # D4-a: Table count
    pdf_tables = pdf.table_count
    sgml_tables = sgml_data["table_count"]
    if pdf_tables > 0:
        if sgml_tables == 0:
            # Before flagging CRITICAL, check if the table content actually appears
            # in the SGML text (contact lists are correctly encoded as <P1><LINE>
            # rather than <TABLE> — this is proper SGML practice, not missing tables).
            # If ≥65% of PDF table cell words appear in the SGML text blob, the
            # content IS present and no critical flag should be raised.
            _table_content_in_sgml = False
            if pdf_tables <= 3 and pdf.pdf_table_cells:
                _sgml_lc = sgml_data["text"].lower()
                _tcell_found = 0; _tcell_total = 0
                for _tc in pdf.pdf_table_cells[:40]:
                    _tw = [w for w in _norm(_tc).split() if len(w) > 3]
                    if not _tw:
                        continue
                    _tcell_total += 1
                    if sum(1 for w in _tw if w in _sgml_lc) / len(_tw) >= 0.65:
                        _tcell_found += 1
                if _tcell_total > 0:
                    _table_content_in_sgml = (_tcell_found / _tcell_total) >= 0.65
            if not _table_content_in_sgml:
                pts = min(2.0, pdf_tables * 0.5)
                score -= pts
                _add_issue(result, "completeness", "critical",
                           f"D4 — PDF has ~{pdf_tables} table(s) but SGML has no <TABLE> tags. "
                           f"Tables may have been dropped entirely.",
                           impact=f"-{pts:.1f} pts")
        elif abs(sgml_tables - pdf_tables) > max(1, pdf_tables * 0.25):
            pts = 1.0
            score -= pts
            _add_issue(result, "completeness", "major",
                       f"D4 — Table count mismatch: PDF ~{pdf_tables}, SGML {sgml_tables}. "
                       f"Some tables may be missing or split.",
                       impact=f"-{pts:.1f} pt")

    # D4-b: Image count
    # Vendor SGML practice: logos, signature blocks, and decorative images (≤5) are
    # intentionally omitted from SGML GRAPHIC tags. Only flag significant image drops
    # (>5 images with 0 GRAPHICs) or when SGML acknowledges images but has fewer.
    pdf_images = pdf.image_count
    sgml_graphics = sgml_data["graphic_count"]
    if pdf_images > 5 and sgml_graphics == 0:
        pts = min(0.75, pdf_images * 0.1)
        score -= pts
        _add_issue(result, "completeness", "minor",
                   f"D4 — PDF has {pdf_images} image(s) but SGML has no <GRAPHIC> tags.",
                   impact=f"-{pts:.1f} pts")
    elif pdf_images > 0 and sgml_graphics > 0 and sgml_graphics < pdf_images - 2:
        score -= 0.5
        _add_issue(result, "completeness", "minor",
                   f"D4 — Image count: PDF {pdf_images} vs SGML {sgml_graphics} <GRAPHIC>. "
                   f"Some images may be missing.",
                   impact="-0.5 pts")

    # D4-c: Footnote count
    pdf_fn = pdf.footnote_count
    sgml_fn = sgml_data["fn_count"]
    if pdf_fn > 4 and sgml_fn == 0:
        score -= 1.0
        _add_issue(result, "completeness", "major",
                   f"D4 — PDF has ~{pdf_fn} footnote(s) but SGML has no <FN> tags. "
                   f"Footnotes may have been dropped.",
                   impact="-1.0 pt")
    elif pdf_fn > 0 and sgml_fn > 0 and pdf_fn > sgml_fn + 3:
        score -= 0.5
        _add_issue(result, "completeness", "minor",
                   f"D4 — Footnote count: PDF ~{pdf_fn} vs SGML {sgml_fn}. "
                   f"Some footnotes may be missing.",
                   impact="-0.5 pts")

    # D4-d: Section/heading count ratio
    pdf_sections = len(pdf.headings)
    sgml_sections = len(sgml_data["headings"])
    if pdf_sections > 3 and sgml_sections > 0:
        ratio = sgml_sections / pdf_sections
        # Skip check when SGML has very few sections AND PDF detects many more
        # than expected. This is a false positive for short notice/alert documents
        # where font-size heading detection picks up table column headers, bold
        # data fields, or price entries (e.g. TMX price-list alerts). A document
        # with SGML sections < 4 legitimately has no multi-section structure.
        is_false_heading_detection = sgml_sections < 4 and pdf_sections > sgml_sections * 4
        if not is_false_heading_detection:
            if ratio < 0.5:
                score -= 1.0
                _add_issue(result, "completeness", "major",
                           f"D4 — Section count: PDF {pdf_sections} headings vs SGML "
                           f"{sgml_sections} <TI> tags ({ratio:.0%} coverage). "
                           f"Sections may be missing or merged.",
                           impact="-1.0 pt")
            elif ratio < 0.70:
                score -= 0.5
                _add_issue(result, "completeness", "minor",
                           f"D4 — Section coverage {ratio:.0%}: PDF {pdf_sections} vs "
                           f"SGML {sgml_sections} headings.",
                           impact="-0.5 pts")

    # D4-e: Schedule/appendix detection
    # NOTE: Vendor practice is to integrate Schedule/Appendix content into the main
    # FREEFORM/BLOCK body rather than using <APPENDIX>/<SCHEDDOC> tags. This is valid
    # per the keying spec. Only issue an informational warning — no score deduction.
    appendix_in_pdf = bool(re.search(
        r"\b(Schedule|Appendix|Annex|Exhibit)\s+[A-Z\d]", pdf.first_page_text, re.IGNORECASE
    ))
    appendix_in_sgml = bool(re.search(r"<APPENDIX|<SCHEDDOC|Schedule\s+[A-Z\d]", sgml_data["text"]))
    if appendix_in_pdf and not appendix_in_sgml:
        result.warnings.append(
            "D4 — PDF appears to have Schedule/Appendix content but SGML has no "
            "<APPENDIX> or <SCHEDDOC>. Vendor practice: integrate into FREEFORM body (no penalty)."
        )

    # D4-f: Table cell content coverage (GAP 5 — DOCX cell-by-cell comparison)
    # Only runs when the ABBYY DOCX is available and contains table cells.
    # D4-f: Table cell content coverage (GAP 5 — DOCX cell-by-cell comparison).
    # When DOCX available: LLM-augmented verification of each cell.
    # When no DOCX: fall through to D4-h (PDF cells direct).
    if docx_data and docx_data.get("table_cells"):
        _docx_cells = [c for c in docx_data["table_cells"] if len(c.split()) >= 3]
        _sgml_cells = sgml_data.get("table_cells", [])
        _sgml_text_for_table = sgml_data.get("text", "")
        if _docx_cells:
            _missing_cells: list[dict] = []

            # LLM path — Opus verifies cells against the full SGML text (not just TBLCELL)
            # so it catches tables rendered as body paragraphs with no <TABLE> tags.
            if _LLM_ENABLED:
                _llm_cell_results = _llm_verify_table_cells(
                    _docx_cells[:_LLM_MAX_CELLS],
                    _sgml_text_for_table,
                )
                if _llm_cell_results is not None:
                    for _cr in _llm_cell_results:
                        if not _cr.get("found", True):
                            _cidx = _cr.get("cell_idx", 1) - 1
                            _cell_txt = _docx_cells[_cidx] if 0 <= _cidx < len(_docx_cells) else _cr.get("text", "")[:80]
                            _missing_cells.append({"text": _cell_txt[:80], "confidence": 0.0, "method": "llm"})
                else:
                    # Deterministic fallback
                    if _sgml_cells:
                        _sgml_cell_blob = " ".join(_sgml_cells)
                        _scw = _sgml_cell_blob.split()
                        _sc_ngrams: set[tuple] = set()
                        _ngram_sz = 5
                        for _i in range(len(_scw) - _ngram_sz + 1):
                            _sc_ngrams.add(tuple(_scw[_i:_i + _ngram_sz]))
                        for _cell in _docx_cells:
                            _cwords = _norm(_cell).split()
                            _is_cov, _conf, _meth = _para_covered_v2(_cwords, _sgml_cell_blob, _sc_ngrams)
                            if not _is_cov:
                                _missing_cells.append({"text": _cell[:80], "confidence": _conf, "method": _meth})
            else:
                # Deterministic only
                if _sgml_cells:
                    _sgml_cell_blob = " ".join(_sgml_cells)
                    _scw = _sgml_cell_blob.split()
                    _sc_ngrams2: set[tuple] = set()
                    for _i in range(len(_scw) - 5 + 1):
                        _sc_ngrams2.add(tuple(_scw[_i:_i + 5]))
                    for _cell in _docx_cells:
                        _cwords = _norm(_cell).split()
                        _is_cov, _conf, _meth = _para_covered_v2(_cwords, _sgml_cell_blob, _sc_ngrams2)
                        if not _is_cov:
                            _missing_cells.append({"text": _cell[:80], "confidence": _conf, "method": _meth})

            result.d4_missing_table_cells = _missing_cells

            if _missing_cells:
                _cell_cov = 1.0 - len(_missing_cells) / len(_docx_cells)
                if _cell_cov < 0.75:
                    _pts = 1.5
                    score -= _pts
                    _add_issue(result, "completeness", "major",
                               f"D4 — Table cell coverage {_cell_cov:.0%}: "
                               f"{len(_missing_cells)} of {len(_docx_cells)} DOCX table cell(s) "
                               f"not found in SGML. "
                               f"Examples: {[m['text'][:50] for m in _missing_cells[:2]]}",
                               impact=f"-{_pts:.1f} pts")
                elif _cell_cov < 0.90:
                    _pts = 0.75
                    score -= _pts
                    _add_issue(result, "completeness", "minor",
                               f"D4 — Table cell coverage {_cell_cov:.0%}: "
                               f"{len(_missing_cells)} DOCX table cell(s) not found in SGML.",
                               impact=f"-{_pts:.1f} pts")

    # D4-h: PDF table cells (direct pdfplumber extraction — no DOCX required).
    # Runs when no DOCX is available but pdfplumber found table cells in the PDF.
    # LLM-augmented: Opus checks cells against the FULL SGML text (not just TBLCELL)
    # so it catches tables rendered as body paragraphs.
    #
    # Contact-directory guard: When SGML has no <TABLE> tags at all (sgml_tables==0),
    # contact lists are correctly encoded as <P1><LINE> — not <TABLE>. pdfplumber
    # sees the 2-column contact grid as a table, but the content IS present in SGML
    # in a different structural form. Detect this by checking word-level coverage of
    # the cell text against the SGML body blob before running the expensive LLM check.
    # If ≥60% of cell words appear in the SGML blob → content IS there → skip D4-h.
    _d4h_contact_dir_skip = False
    if (not docx_data) and pdf.pdf_table_cells and sgml_data.get("table_count", 0) == 0:
        _sgml_lc2 = sgml_data.get("text", "").lower()
        _ch_cells = [c for c in pdf.pdf_table_cells if len(c.split()) >= 3]
        if _ch_cells:
            _ch_covered = 0
            for _chc in _ch_cells[:40]:
                _chw = [w for w in _norm(_chc).split() if len(w) > 3]
                if not _chw:
                    _ch_covered += 1
                    continue
                if sum(1 for w in _chw if w in _sgml_lc2) / len(_chw) >= 0.60:
                    _ch_covered += 1
            if _ch_covered / len(_ch_cells[:40]) >= 0.60:
                _d4h_contact_dir_skip = True  # content present as P1/LINE, not TABLE

    if (not docx_data) and pdf.pdf_table_cells and not _d4h_contact_dir_skip:
        _pdf_cells_direct = [c for c in pdf.pdf_table_cells if len(c.split()) >= 3]
        _sgml_text_for_pdf_table = sgml_data.get("text", "")
        if _pdf_cells_direct:
            _pdf_direct_missing: list[dict] = []

            if _LLM_ENABLED:
                _llm_pdf_cell_results = _llm_verify_table_cells(
                    _pdf_cells_direct[:_LLM_MAX_CELLS],
                    _sgml_text_for_pdf_table,
                )
                if _llm_pdf_cell_results is not None:
                    for _cr in _llm_pdf_cell_results:
                        if not _cr.get("found", True):
                            _cidx = _cr.get("cell_idx", 1) - 1
                            _cell_txt = _pdf_cells_direct[_cidx] if 0 <= _cidx < len(_pdf_cells_direct) else _cr.get("text", "")[:80]
                            _pdf_direct_missing.append({"text": _cell_txt[:80], "confidence": 0.0, "method": "llm"})
                else:
                    # Deterministic fallback
                    _sgml_cells_d = sgml_data.get("table_cells", [])
                    if _sgml_cells_d:
                        _sgml_cb_d = " ".join(_sgml_cells_d)
                        _sc_w_d = _sgml_cb_d.split()
                        _sc_ng_d: set[tuple] = set()
                        for _i in range(len(_sc_w_d) - 5 + 1):
                            _sc_ng_d.add(tuple(_sc_w_d[_i:_i + 5]))
                        for _pc in _pdf_cells_direct[:80]:
                            _pcw = _norm(_pc).split()
                            _is_cov2, _conf2, _meth2 = _para_covered_v2(_pcw, _sgml_cb_d, _sc_ng_d)
                            if not _is_cov2:
                                _pdf_direct_missing.append({"text": _pc[:80], "confidence": _conf2, "method": _meth2})
            else:
                _sgml_cells_d = sgml_data.get("table_cells", [])
                if _sgml_cells_d:
                    _sgml_cb_d = " ".join(_sgml_cells_d)
                    _sc_w_d = _sgml_cb_d.split()
                    _sc_ng_d2: set[tuple] = set()
                    for _i in range(len(_sc_w_d) - 5 + 1):
                        _sc_ng_d2.add(tuple(_sc_w_d[_i:_i + 5]))
                    for _pc in _pdf_cells_direct[:80]:
                        _pcw = _norm(_pc).split()
                        _is_cov2, _conf2, _meth2 = _para_covered_v2(_pcw, _sgml_cb_d, _sc_ng_d2)
                        if not _is_cov2:
                            _pdf_direct_missing.append({"text": _pc[:80], "confidence": _conf2, "method": _meth2})

            result.pdf_direct_table_cells_missing = _pdf_direct_missing
            if _pdf_direct_missing:
                _dcov = 1.0 - len(_pdf_direct_missing) / max(len(_pdf_cells_direct), 1)
                if _dcov < 0.80:
                    _dpts = 1.0
                    score -= _dpts
                    _add_issue(result, "completeness", "major",
                               f"D4 — PDF table cell coverage {_dcov:.0%}: "
                               f"{len(_pdf_direct_missing)} of {len(_pdf_cells_direct)} PDF table "
                               f"cell(s) not found in SGML. "
                               f"Examples: {[m['text'][:50] for m in _pdf_direct_missing[:2]]}",
                               impact=f"-{_dpts:.1f} pts")

    result.completeness_score = max(0.0, score)


# ─────────────────────────────────────────────────────────────────────────────
# D5: Ordering / sequence
# ─────────────────────────────────────────────────────────────────────────────
def check_ordering(pdf: _PDFData, sgml_data: dict, result: L4Result,
                   docx_data: "dict | None" = None) -> None:
    """
    D5 — 4 pts: Validate that sections appear in the same order as the PDF.

    D5-a: Section heading order via inversion count (heading-level).
    D5-b: Paragraph sequence check via DOCX (GAP 6 — body-level ordering).
          When docx_data is provided, takes up to 30 meaningful DOCX paragraphs,
          locates their 8-word fingerprints in the SGML text, and counts positional
          inversions to detect gross reordering of body content.
    Multi-column layout caveat: flagged as WARNING (may be false positive).
    """
    score = 4.0

    # Gap 9: suppress D5 for 2-column PDFs — ordering is unreliable
    if pdf.two_column:
        result.ordering_score = 4.0
        result.warnings.append(
            "D5 skipped: 2-column PDF layout detected. Section ordering cannot be "
            "reliably validated (PyMuPDF reads left-to-right across columns). "
            "Full score awarded."
        )
        return

    pdf_headings_norm = [_norm(h) for h in pdf.headings if len(h.split()) >= 2]
    sgml_headings_norm = [h for h in sgml_data["headings"] if len(h.split()) >= 1]

    if len(pdf_headings_norm) < 3 or len(sgml_headings_norm) < 3:
        # Not enough headings to validate order meaningfully
        result.ordering_score = 4.0
        return

    # Match SGML headings to PDF headings and record PDF order positions
    pdf_positions: list[int] = []
    for sh in sgml_headings_norm:
        best_pos = -1
        best_ratio = 0.0
        for i, ph in enumerate(pdf_headings_norm):
            ratio = SequenceMatcher(None, sh, ph).ratio()
            if ratio > best_ratio and ratio >= 0.60:
                best_ratio = ratio
                best_pos = i
        if best_pos >= 0:
            pdf_positions.append(best_pos)

    if len(pdf_positions) < 3:
        result.ordering_score = 4.0
        return

    # Count inversions (O(n²) — acceptable for typical heading counts < 50)
    # Also collect the first few inverted pairs for diff_generator
    inversions = 0
    n = len(pdf_positions)
    _inverted_pairs: list[tuple] = []
    for i in range(n):
        for j in range(i + 1, n):
            if pdf_positions[i] > pdf_positions[j]:
                inversions += 1
                if len(_inverted_pairs) < 5:
                    # i appears before j in SGML but after j in PDF
                    _inverted_pairs.append(
                        (sgml_headings_norm[i], sgml_headings_norm[j])
                    )
    result.d5_inverted_pairs = _inverted_pairs

    max_inversions = n * (n - 1) / 2
    inversion_ratio = inversions / max_inversions if max_inversions > 0 else 0

    if inversion_ratio > 0.3:
        pts = min(2.0, inversion_ratio * 4.0)
        score -= pts
        severity = "major" if pts >= 1.0 else "minor"
        _add_issue(result, "ordering", severity,
                   f"D5 — Section order mismatch: {inversion_ratio:.0%} of heading pairs "
                   f"appear in wrong order vs PDF. "
                   f"({inversions} inversion(s) across {n} matched sections)",
                   impact=f"-{pts:.1f} pts")
        result.sequence_violations.append(
            f"{inversions} section ordering inversion(s) detected"
        )
    elif inversion_ratio > 0.1:
        score -= 0.5
        _add_issue(result, "ordering", "minor",
                   f"D5 — Minor section reordering detected ({inversion_ratio:.0%} inversion rate). "
                   f"May be false positive for multi-column layouts.",
                   impact="-0.5 pts")
        result.warnings.append(
            "D5: Minor reordering detected — may be false positive for 2-column PDF layouts."
        )

    # List item sequence: check (a), (b), (c) order within SGML ITEM tags
    items = re.findall(r"<ITEM[^>]*>.*?</ITEM>", sgml_data.get("_raw_sgml", ""), re.DOTALL)
    # (a), (b), (c) patterns
    list_labels = []
    for item in items:
        m = re.search(r"^\s*\(([a-z])\)", re.sub(r"<[^>]+>", "", item).strip())
        if m:
            list_labels.append(ord(m.group(1)) - ord('a'))

    if len(list_labels) >= 3:
        list_inversions = sum(
            1 for i in range(len(list_labels) - 1)
            if list_labels[i] > list_labels[i + 1]
        )
        if list_inversions > 2:
            score -= 0.5
            _add_issue(result, "ordering", "minor",
                       f"D5 — List item order: {list_inversions} out-of-sequence (a)/(b)/(c) "
                       f"ITEM label(s) detected.",
                       impact="-0.5 pts")

    # D5-b: Paragraph sequence check (GAP 6 — DOCX body-level ordering)
    # When DOCX is available, verify that meaningful paragraphs appear in the
    # same left-to-right order in SGML as they do in the DOCX.
    # Uses an 8-word fingerprint search in the SGML text blob.
    if docx_data and docx_data.get("paragraphs") and score > 0:
        _sgml_text = sgml_data["text"].lower()
        _docx_paras = [
            p for p in docx_data["paragraphs"]
            if len(p.split()) >= 10
        ][:30]  # sample up to 30 meaningful paragraphs

        _para_positions: list[int] = []
        for _para in _docx_paras:
            _words = _norm(_para).split()
            _fp = " ".join(_words[:8])   # 8-word fingerprint
            _pos = _sgml_text.find(_fp)
            if _pos >= 0:
                _para_positions.append(_pos)

        if len(_para_positions) >= 5:
            # Count positional inversions
            _n_p = len(_para_positions)
            _para_inv = sum(
                1 for _i in range(_n_p)
                for _j in range(_i + 1, _n_p)
                if _para_positions[_i] > _para_positions[_j]
            )
            _max_inv = _n_p * (_n_p - 1) / 2
            _para_inv_ratio = _para_inv / _max_inv if _max_inv > 0 else 0.0

            if _para_inv_ratio > 0.30:
                _pts = min(1.5, _para_inv_ratio * 3.0)
                score -= _pts
                _add_issue(result, "ordering", "major",
                           f"D5 — Paragraph sequence mismatch: {_para_inv_ratio:.0%} inversion "
                           f"rate across {_n_p} matched DOCX paragraphs. Body content sections "
                           f"may be reordered in the SGML.",
                           impact=f"-{_pts:.1f} pts")
            elif _para_inv_ratio > 0.15:
                score -= 0.5
                _add_issue(result, "ordering", "minor",
                           f"D5 — Minor paragraph reordering detected ({_para_inv_ratio:.0%} "
                           f"inversion rate across {_n_p} matched paragraphs).",
                           impact="-0.5 pts")

    result.ordering_score = max(0.0, score)


# ─────────────────────────────────────────────────────────────────────────────
# D7: Metadata accuracy
# ─────────────────────────────────────────────────────────────────────────────
def check_metadata(pdf: _PDFData, sgml_data: dict, raw_sgml: str, result: L4Result) -> None:
    """
    D7 — 3 pts: Validate POLIDOC metadata against PDF-extracted values.

    Scoring:
      1.0 pt  Language  (SGML LANG vs PDF language heuristic)
      1.0 pt  Doc number (SGML <N> value found anywhere in PDF text)
      1.0 pt  Date      (PDF date found in SGML; ADDDATE vs pub date is soft warning only)

    GAP 7 additions (soft checks — warnings, no additional point deductions):
      • Required fields completeness (LABEL, LANG, ADDDATE present and non-empty)
      • LANG valid-value check (must be EN, FR, or BI)
      • ADDDATE / MODDATE date-format validation (must be YYYYMMDD)
      • MODDATE consistency (MODDATE must be ≥ ADDDATE if both present)
    """
    score = 3.0
    attrs = sgml_data["attrs"]
    mismatches: list[str] = []

    # Build full PDF text for membership checks
    pdf_text_norm = _norm(pdf.first_page_text)

    # Store for diff_generator
    result.d7_expected_lang = pdf.language_hint
    result.d7_pdf_doc_number = pdf.doc_number

    # D7-a: Language — does PDF language match SGML LANG attribute?
    sgml_lang = attrs.get("LANG", "")
    if sgml_lang and pdf.language_hint and sgml_lang != pdf.language_hint:
        score -= 1.0
        mismatches.append(f"LANG={sgml_lang} but PDF appears to be {pdf.language_hint}")
        _add_issue(result, "metadata", "major",
                   f"D7 — POLIDOC LANG='{sgml_lang}' but PDF text suggests language "
                   f"'{pdf.language_hint}'. Document may be mis-classified.",
                   impact="-1.0 pt")

    # D7-b: Document number — SGML <N> tag value should appear in PDF text.
    # Strategy: extract ALL NN-NNN style numbers from the PDF and check if
    # the <N> tag value is among them. This avoids picking up referenced
    # instruments from body text as the document's own number.
    sgml_n_tags = re.findall(r"<N[^>]*>(.*?)</N>", raw_sgml, re.DOTALL)
    sgml_n_values = [re.sub(r"\s+", " ", v).strip() for v in sgml_n_tags[:3]]

    if sgml_n_values:
        # Extract all doc-number-like tokens from PDF text
        pdf_numbers_found = set(re.findall(r"\d{2}-\d{3,4}", pdf.first_page_text))
        # Also include bare numbers like '13-103'
        n_val = sgml_n_values[0]  # first <N> is primary doc number
        n_stripped = re.sub(r"[\s\-]", "", n_val)  # e.g. '45930'
        n_bare = re.search(r"\d{2}-\d{3,4}", n_val)  # e.g. '45-930'
        n_bare_str = n_bare.group(0) if n_bare else ""

        # TMX/alert notices use YYYY-NNN format (e.g. "2025-008", "2025-060").
        # These don't appear prominently in their PDF text and use a different
        # numbering scheme from regulatory NI XX-XXX documents. Treat as
        # warning-only (no point deduction) to avoid false D7 penalties.
        is_yyyy_nnn = bool(re.match(r"^\d{4}-\d{3}$", n_val))

        found_in_pdf = (
            n_val in pdf.first_page_text or
            (n_bare_str and n_bare_str in pdf_numbers_found) or
            n_stripped in pdf.first_page_text.replace("-", "").replace(" ", "")
        )
        if not found_in_pdf and n_bare_str and not is_yyyy_nnn:
            score -= 1.0
            mismatches.append(f"<N>={n_val!r} not found in PDF first pages")
            _add_issue(result, "metadata", "major",
                       f"D7 — SGML <N> value '{n_val}' not found in PDF text. "
                       f"Document number may be wrong. PDF numbers found: "
                       f"{sorted(pdf_numbers_found)[:5]}",
                       impact="-1.0 pt")
        elif not found_in_pdf and is_yyyy_nnn:
            result.warnings.append(
                f"D7: <N> value '{n_val}' (YYYY-NNN format) not found on PDF first pages — "
                f"TMX/alert notice numbers may not appear in extracted PDF text."
            )

    # D7-c: Date — PDF publication date vs SGML date attributes.
    # ADDDATE is the *keying* date (can differ from pub date) — treated as
    # WARNING only (no point deduction). We check that the PDF date exists
    # somewhere in the SGML text as a loose sanity check.
    if pdf.doc_date:
        sgml_adddate = attrs.get("ADDDATE", "")
        if sgml_adddate:
            try:
                from datetime import datetime
                d_pdf = datetime.strptime(pdf.doc_date, "%Y%m%d")
                d_sgml = datetime.strptime(sgml_adddate, "%Y%m%d")
                delta = abs((d_pdf - d_sgml).days)
                if delta > 1825:  # > 5 years
                    # ADDDATE is the vendor *keying* date, not the publication date.
                    # Historical documents converted in bulk always have a large gap
                    # between ADDDATE (today) and the PDF publication date (years ago).
                    # Deducting points here creates systematic false positives on all
                    # vendor-converted historical SGMLs. Treat as warning only.
                    result.warnings.append(
                        f"D7: ADDDATE='{sgml_adddate}' differs from PDF date '{pdf.doc_date}' "
                        f"by {delta} days — keying date vs historical publication date (expected)."
                    )
                elif delta > 30:
                    # Common case: ADDDATE is the keying date, not the publication date.
                    # Amending instruments re-keyed years after original publication
                    # legitimately have large gaps. Warn only, no point deduction.
                    result.warnings.append(
                        f"D7: ADDDATE ({sgml_adddate}) vs PDF date ({pdf.doc_date}) "
                        f"differ by {delta} days — likely keying date vs publication date."
                    )
            except ValueError:
                pass

    result.metadata_mismatches = mismatches
    result.metadata_score = max(0.0, score)

    # ── GAP 7: Soft metadata completeness checks (warnings only, no score impact) ──

    # D7-d: Required POLIDOC attributes present and non-empty
    _required = ("LABEL", "LANG", "ADDDATE")
    for _attr in _required:
        if not attrs.get(_attr, "").strip():
            result.warnings.append(
                f"D7: Required POLIDOC attribute '{_attr}' is missing or empty. "
                f"This may cause downstream processing failures."
            )
            mismatches.append(f"Missing required attribute: {_attr}")

    # D7-e: LANG must be one of the known valid values (EN, FR, BI)
    _valid_langs = {"EN", "FR", "BI"}
    _sgml_lang_val = attrs.get("LANG", "").strip().upper()
    if _sgml_lang_val and _sgml_lang_val not in _valid_langs:
        result.warnings.append(
            f"D7: POLIDOC LANG='{_sgml_lang_val}' is not a recognised value. "
            f"Expected one of: {sorted(_valid_langs)}."
        )
        mismatches.append(f"Invalid LANG value: {_sgml_lang_val!r}")

    # D7-f: Date fields must be valid YYYYMMDD dates; MODDATE must be ≥ ADDDATE
    from datetime import datetime as _dt
    _date_attrs = {k: attrs.get(k, "").strip() for k in ("ADDDATE", "MODDATE")}
    _parsed_dates: dict[str, _dt] = {}
    for _attr, _val in _date_attrs.items():
        if not _val:
            continue
        try:
            _parsed_dates[_attr] = _dt.strptime(_val, "%Y%m%d")
        except ValueError:
            result.warnings.append(
                f"D7: POLIDOC {_attr}='{_val}' is not a valid YYYYMMDD date."
            )
            mismatches.append(f"Invalid date format: {_attr}={_val!r}")

    if "ADDDATE" in _parsed_dates and "MODDATE" in _parsed_dates:
        if _parsed_dates["MODDATE"] < _parsed_dates["ADDDATE"]:
            result.warnings.append(
                f"D7: MODDATE ({attrs['MODDATE']}) is earlier than ADDDATE "
                f"({attrs['ADDDATE']}). A modification date cannot precede the creation date."
            )
            mismatches.append(
                f"MODDATE ({attrs['MODDATE']}) < ADDDATE ({attrs['ADDDATE']})"
            )

    # D7-g: LABEL should be a non-trivial document type string (not just digits/symbols)
    _label_val = attrs.get("LABEL", "").strip()
    if _label_val and not re.search(r"[A-Za-z]{3}", _label_val):
        result.warnings.append(
            f"D7: POLIDOC LABEL='{_label_val}' does not look like a valid document type label."
        )
        mismatches.append(f"Suspicious LABEL value: {_label_val!r}")


# ─────────────────────────────────────────────────────────────────────────────
# D4-g/h: Contact details & hyperlink verification
# ─────────────────────────────────────────────────────────────────────────────
def _normalize_phone(phone: str) -> str:
    """Strip non-digit chars for phone comparison (e.g. (416) 555-1234 → 4165551234)."""
    digits = re.sub(r"\D", "", phone)
    # Remove leading country code '1' for 11-digit North American numbers
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits if len(digits) == 10 else ""


def _normalize_url(url: str) -> str:
    """Normalise URL: lowercase scheme+host, strip trailing punctuation and slash."""
    url = url.rstrip(".,;:)>\"'/")  # also strip trailing slash
    # Normalize http(s)://www.X → www.X so that PDF hyperlink annotations
    # (which include the scheme) match SGML text URLs (which often omit it).
    url = re.sub(r'^https?://(www\.)', r'\1', url, flags=re.I)
    # Lowercase only scheme and host (path is case-sensitive on most servers)
    m = re.match(r"(https?://|www\.)([^/?#]*)(.*)", url, re.I)
    if m:
        return m.group(1).lower() + m.group(2).lower() + m.group(3)
    return url.lower()


def check_contact_details(
    pdf: "_PDFData",
    sgml_data: dict,
    raw_sgml: str,
    result: "L4Result",
) -> None:
    """
    D4-g/h: Comprehensive contact-detail and hyperlink verification.

    Extracts and compares between PDF and SGML:
      • Email addresses           (D4-g-email)
      • Phone numbers             (D4-g-phone)
      • Fax numbers               (D4-g-fax — same pattern, context-detected)
      • URLs / hyperlinks         (D4-h-url)
      • PDF annotation link URIs  (D4-h-uri — from PyMuPDF page.get_links())
      • Canadian postal codes     (D4-g-postal)

    Bidirectional: flags items in PDF but missing from SGML (missing),
    and items in SGML but NOT in PDF (extra — possible fabrication).

    Deductions applied to result.completeness_score (max -2.0 pts total).
    """
    # ── Build source blobs ────────────────────────────────────────────────────
    # PDF: join all paragraph text + all_text (full unfiltered, all pages)
    # all_text bypasses the word-count filter and the 3000-char page cap,
    # ensuring contact details on any page are found.
    pdf_text_full = " ".join(pdf.paragraphs) + " " + " ".join(pdf.raw_lines)
    pdf_text_full += " " + pdf.all_text

    # SGML: search raw SGML string (covers text content AND href/uri attributes)
    sgml_text_full = raw_sgml

    # ── Email addresses ───────────────────────────────────────────────────────
    _pdf_emails = {e.lower().rstrip(".") for e in _EMAIL_RE_L4.findall(pdf_text_full)}
    _sgml_emails = {e.lower().rstrip(".") for e in _EMAIL_RE_L4.findall(sgml_text_full)}
    result.missing_emails = sorted(_pdf_emails - _sgml_emails)
    result.extra_emails = sorted(_sgml_emails - _pdf_emails)

    # ── Phone & fax numbers ───────────────────────────────────────────────────
    _pdf_phones_raw = _PHONE_RE_L4.findall(pdf_text_full)
    _sgml_phones_raw = _PHONE_RE_L4.findall(sgml_text_full)
    _pdf_phones = {_normalize_phone(p) for p in _pdf_phones_raw if _normalize_phone(p)}
    _sgml_phones = {_normalize_phone(p) for p in _sgml_phones_raw if _normalize_phone(p)}
    result.missing_phones = sorted(_pdf_phones - _sgml_phones)
    result.extra_phones = sorted(_sgml_phones - _pdf_phones)

    # ── URLs / hyperlinks ─────────────────────────────────────────────────────
    def _is_noise_url(u: str) -> bool:
        """Return True for URLs that are extraction artefacts and not real links."""
        ul = u.lower()
        return (
            ul.startswith('mailto:')                               # emails, not URLs
            or 'safelinks.protection.outlook.com' in ul            # MS SafeLinks wrappers
            or 'urldefense.proofpoint.com' in ul                   # Proofpoint wrappers
            # Social media footer links — appear in every PDF footer, never in SGML body
            or re.search(r'(?:facebook\.com|linkedin\.com|twitter\.com|x\.com|'
                         r'instagram\.com|youtube\.com)/(?:company|in/|user/|[a-z])',
                         ul) is not None
            # Email marketing / click-tracking redirectors
            or 'click.email.' in ul
            or re.match(r'https?://(?:www\.)?(?:facebook|linkedin|twitter|x|instagram|youtube)'
                        r'\.com(?:/[^/]{0,40})?$', ul) is not None
        )

    _pdf_urls_text = {_normalize_url(u) for u in _URL_RE_L4.findall(pdf_text_full)
                      if not _is_noise_url(u)}
    _pdf_urls_annot = {_normalize_url(u) for u in pdf.link_uris
                       if not _is_noise_url(u)}
    _pdf_urls = _pdf_urls_text | _pdf_urls_annot

    _sgml_urls_text = {_normalize_url(u) for u in _URL_RE_L4.findall(sgml_text_full)}
    _sgml_urls_attr = {_normalize_url(u) for u in sgml_data.get("sgml_hrefs", [])}
    _sgml_urls = _sgml_urls_text | _sgml_urls_attr

    # Filter noise: skip very short URLs (< 10 chars) that are likely fragments
    _pdf_urls = {u for u in _pdf_urls if len(u) >= 10}
    _sgml_urls = {u for u in _sgml_urls if len(u) >= 10}

    # URL matching: a PDF URL is satisfied if the SGML contains a URL that either
    # (a) exactly matches it, or (b) shares a 30-char prefix with it (handles
    # line-wrap truncation in PDF/SGML where long URLs get cut at different points).
    # This is stricter than domain-level dedup so that distinct resources at the
    # same domain (e.g., oecd.ai/en/ai-principles vs oecd.ai/en/dashboards/...)
    # are treated as different URLs that must both be present.
    _PREFIX_LEN = 30

    def _url_domain(u: str) -> str:
        u = re.sub(r"^(?:https?://|www\.)", "", u, flags=re.I)
        return u.split("/")[0].lower().lstrip("www.")

    def _pdf_url_satisfied(pdf_url: str, sgml_set: set) -> bool:
        """Return True if any SGML URL is a prefix-match for this PDF URL."""
        if pdf_url in sgml_set:
            return True
        # Prefix match: handles line-wrap truncation on either side
        p = pdf_url[:_PREFIX_LEN]
        for su in sgml_set:
            sp = su[:_PREFIX_LEN]
            common_len = min(len(p), len(sp))
            if common_len >= 20 and p[:common_len] == sp[:common_len]:
                return True
        return False

    result.missing_urls = sorted(
        u for u in _pdf_urls
        if not _pdf_url_satisfied(u, _sgml_urls)
    )
    result.extra_urls = sorted(_sgml_urls - _pdf_urls)

    # ── Canadian postal codes ─────────────────────────────────────────────────
    _pdf_postal = {re.sub(r"\s+", " ", p.upper()) for p in _POSTAL_CODE_RE.findall(pdf_text_full)}
    _sgml_postal = {re.sub(r"\s+", " ", p.upper()) for p in _POSTAL_CODE_RE.findall(sgml_text_full)}
    result.missing_postal_codes = sorted(_pdf_postal - _sgml_postal)

    # ── Scoring ───────────────────────────────────────────────────────────────
    _contact_pts = 0.0
    _contact_issues: list[str] = []

    if result.missing_emails:
        _ep = min(1.0, len(result.missing_emails) * 0.5)
        _contact_pts += _ep
        _contact_issues.append(
            f"D4-g — {len(result.missing_emails)} email address(es) from PDF missing "
            f"in SGML: {result.missing_emails[:3]}"
        )
        _add_issue(result, "completeness", "major",
                   f"D4-g — Email address(es) from PDF not found in SGML: "
                   f"{result.missing_emails[:5]}. "
                   f"Email addresses must be reproduced exactly.",
                   impact=f"-{_ep:.1f} pts")

    if result.missing_phones:
        _pp = min(0.5, len(result.missing_phones) * 0.25)
        _contact_pts += _pp
        _add_issue(result, "completeness", "major",
                   f"D4-g — Phone number(s) from PDF not found in SGML: "
                   f"{result.missing_phones[:5]}. "
                   f"Phone numbers must be reproduced exactly.",
                   impact=f"-{_pp:.1f} pts")

    if result.missing_urls:
        _up = min(0.5, len(result.missing_urls) * 0.25)
        _contact_pts += _up
        _add_issue(result, "completeness", "major",
                   f"D4-h — URL/hyperlink(s) from PDF not found in SGML: "
                   f"{[u[:60] for u in result.missing_urls[:3]]}. "
                   f"Hyperlinks must be preserved in SGML.",
                   impact=f"-{_up:.1f} pts")

    if result.missing_postal_codes:
        _add_issue(result, "completeness", "minor",
                   f"D4-g — Canadian postal code(s) from PDF missing in SGML: "
                   f"{result.missing_postal_codes[:5]}.",
                   impact="-0 pts (informational)")

    if result.extra_emails or result.extra_phones or result.extra_urls:
        _extra_items = (
            [f"email:{e}" for e in result.extra_emails[:2]] +
            [f"phone:{p}" for p in result.extra_phones[:2]] +
            [f"url:{u[:40]}" for u in result.extra_urls[:2]]
        )
        result.warnings.append(
            f"D4-g/h — SGML contains contact details NOT present in PDF "
            f"(possible fabrication or copy-paste error): {_extra_items[:5]}"
        )

    # D4-fn: Empty/missing footnote body detection.
    # Catches two tamper patterns:
    #   (a) <FREEFORM> wrapper exists but all text was removed (emptied)
    #   (b) <FOOTNOTE> shell exists but <FREEFORM> was entirely deleted
    # Legitimate SGML never has an empty or content-free FOOTNOTE body.
    _empty_fn_bodies = 0
    for _fn_body in re.findall(
        r"<FOOTNOTE[^>]*>(.*?)</FOOTNOTE>", raw_sgml, re.DOTALL | re.IGNORECASE
    ):
        _ff_bodies = re.findall(
            r"<FREEFORM[^>]*>(.*?)</FREEFORM>", _fn_body, re.DOTALL | re.IGNORECASE
        )
        if not _ff_bodies:
            # (b) FOOTNOTE shell with no FREEFORM content at all — body was deleted
            _empty_fn_bodies += 1
        else:
            for _ff_body in _ff_bodies:
                _ff_text = re.sub(r"<[^>]+>", " ", _ff_body)
                _ff_text = re.sub(r"&[a-zA-Z0-9#]+;", " ", _ff_text).strip()
                if len(_ff_text.split()) < 2:
                    # (a) FREEFORM exists but text was emptied
                    _empty_fn_bodies += 1

    if _empty_fn_bodies > 0:
        result.empty_footnote_bodies = _empty_fn_bodies
        _fn_pts = min(1.0, _empty_fn_bodies * 0.5)
        _contact_pts += _fn_pts
        _add_issue(result, "completeness", "major",
                   f"D4-fn — {_empty_fn_bodies} footnote body/bodies are completely "
                   f"empty (<FREEFORM> wrapper exists but all text was removed). "
                   f"This is a strong indicator that content was deliberately deleted "
                   f"from inside a footnote — human review required.",
                   impact=f"-{_fn_pts:.1f} pts")

    # Apply to completeness score (floor at 0)
    if _contact_pts > 0:
        result.completeness_score = max(0.0, result.completeness_score - min(2.0, _contact_pts))


# ─────────────────────────────────────────────────────────────────────────────
# Main entry point
# ─────────────────────────────────────────────────────────────────────────────
def validate_source_comparison(
    raw_sgml: str,
    pdf_path: Optional[str] = None,
    docx_path: Optional[str] = None,
) -> L4Result:
    """
    Run all Level 4 source-comparison checks.

    Parameters
    ----------
    raw_sgml  : str           — Raw SGML content as read from file.
    pdf_path  : str, optional — Path to source PDF. If None, only D6 (encoding)
                                runs. All other dimensions are skipped with warnings.
    docx_path : str, optional — Path to ABBYY-generated DOCX (intermediate file).
                                When provided, D3 uses two-stage comparison:
                                PDF→DOCX (ABBYY errors) + DOCX→SGML (pipeline errors).
                                Significantly reduces false positives from headers/footers.

    Returns
    -------
    L4Result with score (0-30) and all issues found.
    """
    result = L4Result()

    # D6 always runs — no PDF needed
    check_encoding(raw_sgml, result)

    if not pdf_path:
        result.pdf_available = False
        result.warnings.append(
            "L4: No source PDF provided. D2/D3/D4/D5/D7 skipped. "
            "Only encoding check (D6) ran."
        )
        # Score: only D6 ran out of 30 pts. Normalise D6 score to 3/30
        result.tagging_score = 0.0
        result.text_score = 0.0
        result.completeness_score = 0.0
        result.ordering_score = 0.0
        result.metadata_score = 0.0
        result.score = result.encoding_score  # out of 3
        return result

    result.pdf_available = True

    # Extract PDF data
    pdf = _extract_pdf_data(pdf_path)

    # Store PDF headings for D3 placement heuristic in diff_generator
    result.pdf_headings = pdf.headings  # all headings — no cap

    if not pdf.ok:
        result.pdf_text_extractable = False
        result.warnings.append(f"L4: PDF extraction failed ({pdf.error}). D2-D5/D7 skipped.")
        result.score = result.encoding_score
        return result

    result.pdf_text_extractable = True

    # Extract SGML structured data (pass raw for D5 list check)
    sgml_data = _extract_sgml_text(raw_sgml)
    sgml_data["_raw_sgml"] = raw_sgml  # pass through for D5 list-item check

    # Run all dimensions with individual error isolation
    # GAP 4: parse DOCX once here — share with D2 (check_tagging) and D3 (check_text_accuracy)
    # so we don't parse the DOCX file twice.
    _docx_data: dict | None = None
    if docx_path:
        _d = _extract_docx_text(docx_path)
        if _d["ok"] and (_d["paragraphs"] or _d["bold_runs"] or _d["italic_runs"]):
            _docx_data = _d
        else:
            result.warnings.append(
                f"D2/D3 — DOCX extraction failed ({_d.get('error', 'unknown')}). "
                "Falling back to PDF-only validation."
            )

    try:
        check_tagging(pdf, sgml_data, raw_sgml, result, docx_data=_docx_data)
    except Exception as e:
        result.tagging_score = 5.0  # assume pass on error
        result.warnings.append(f"D2 check error (skipped): {e}")

    try:
        check_text_accuracy(pdf, sgml_data, result, docx_data=_docx_data)
    except Exception as e:
        result.text_score = 8.0
        result.warnings.append(f"D3 check error (skipped): {e}")

    try:
        check_completeness(pdf, sgml_data, result, docx_data=_docx_data)
    except Exception as e:
        result.completeness_score = 7.0
        result.warnings.append(f"D4 check error (skipped): {e}")

    try:
        check_ordering(pdf, sgml_data, result, docx_data=_docx_data)
    except Exception as e:
        result.ordering_score = 4.0
        result.warnings.append(f"D5 check error (skipped): {e}")

    try:
        check_metadata(pdf, sgml_data, raw_sgml, result)
    except Exception as e:
        result.metadata_score = 3.0
        result.warnings.append(f"D7 check error (skipped): {e}")

    # D4-g/h: Contact details & hyperlink check (runs after D4 so it can
    # deduct from completeness_score which D4 already set)
    try:
        check_contact_details(pdf, sgml_data, raw_sgml, result)
    except Exception as e:
        result.warnings.append(f"D4-g/h contact check error (skipped): {e}")

    result.score = (
        result.tagging_score +
        result.text_score +
        result.completeness_score +
        result.ordering_score +
        result.encoding_score +
        result.metadata_score
    )

    # Enrich L4 issues with actionable fix templates
    try:
        from validator.core.fix_templates import enrich_issues
        enrich_issues(result.issues)
    except Exception:
        pass

    return result
