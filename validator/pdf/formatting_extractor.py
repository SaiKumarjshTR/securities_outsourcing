"""
formatting_extractor.py — Bold, italic, heading detection from PyMuPDF spans.

Design decisions:
- Trust font FLAGS first (most reliable for digital PDFs), font NAME second.
- Both can disagree — e.g. some PDFs set the Bold flag but use a regular-weight
  font name variant. We OR both checks so we never miss a detection.
- Heading level is determined by relative font size vs the body baseline.
  We compute the body baseline dynamically (modal font size on that page),
  not a hardcoded threshold, because legal PDFs vary from 8pt to 14pt body.
- "Fake bold" (thicker stroke via PDF graphics state, no flag set) is
  detectable via the 'color' trick in some PDFs but is rare in our corpus;
  we don't over-engineer for it.
"""
import re
import logging
from typing import Dict, List, Optional
from collections import Counter

log = logging.getLogger(__name__)

# ── Font name pattern sets ────────────────────────────────────────────────────
_BOLD_NAME_RE = re.compile(
    r"(?:^|[,-])(?:Bold|Bd|Black|Heavy|Semibold|SemiBold|Demi|ExtraBold)"
    r"|Bold(?:MT)?$",
    re.IGNORECASE,
)
_ITALIC_NAME_RE = re.compile(
    r"(?:^|[,-])(?:Italic|It|Oblique|Slanted|Cursive)"
    r"|(?:Italic|Oblique)(?:MT)?$",
    re.IGNORECASE,
)

# PyMuPDF font flag bit positions
_FLAG_SUPERSCRIPT = 1 << 0   # bit 0 — superscript
_FLAG_ITALIC      = 1 << 1   # bit 1 — italic
_FLAG_BOLD        = 1 << 4   # bit 4 — bold


def is_bold(span: Dict) -> bool:
    """
    Return True if this PyMuPDF span represents bold text.

    Checks:
      1. Font flag bit 4 (most reliable)
      2. Font name contains a bold keyword variant
    """
    flags = span.get("flags", 0)
    if flags & _FLAG_BOLD:
        return True
    font_name = span.get("font", "")
    if font_name and _BOLD_NAME_RE.search(font_name):
        return True
    return False


def is_italic(span: Dict) -> bool:
    """
    Return True if this PyMuPDF span represents italic text.

    Checks:
      1. Font flag bit 1 (most reliable)
      2. Font name contains an italic/oblique keyword variant
    """
    flags = span.get("flags", 0)
    if flags & _FLAG_ITALIC:
        return True
    font_name = span.get("font", "")
    if font_name and _ITALIC_NAME_RE.search(font_name):
        return True
    return False


def is_superscript(span: Dict) -> bool:
    """Return True if span is a superscript (footnote marker candidate)."""
    return bool(span.get("flags", 0) & _FLAG_SUPERSCRIPT)


def compute_body_font_size(spans: List[Dict]) -> float:
    """
    Compute the modal (most common) font size across all spans on a page.
    This is the body text baseline — headings will be larger than this.

    Rounds sizes to nearest 0.5pt to group near-identical sizes together.
    Falls back to 10.0 if no spans provided.
    """
    if not spans:
        return 10.0
    rounded = [round(s.get("size", 10.0) * 2) / 2 for s in spans]
    counter = Counter(rounded)
    modal = counter.most_common(1)[0][0]
    return modal


def detect_heading_level(size: float, body_size: float) -> int:
    """
    Determine heading level (1–4) based on font size relative to body.

    Returns:
        0  — not a heading (size ≤ body + 0.5)
        1  — H1: size > body + 6
        2  — H2: size > body + 3
        3  — H3: size > body + 1.5
        4  — H4: size > body + 0.5
    """
    diff = size - body_size
    if diff > 6:
        return 1
    if diff > 3:
        return 2
    if diff > 1.5:
        return 3
    if diff > 0.5:
        return 4
    return 0


def classify_span(span: Dict, body_size: float) -> Dict:
    """
    Enrich a raw PyMuPDF span dict with computed formatting flags.

    Returns the same dict with added keys:
        _bold     : bool
        _italic   : bool
        _super    : bool
        _heading  : int  (0 = not heading, 1–4 = heading level)
    """
    size = span.get("size", body_size)
    span["_bold"]    = is_bold(span)
    span["_italic"]  = is_italic(span)
    span["_super"]   = is_superscript(span)
    span["_heading"] = detect_heading_level(size, body_size)
    return span


def classify_spans_for_page(spans: List[Dict]) -> List[Dict]:
    """
    Classify an entire page's worth of spans in one call.
    Computes body size dynamically from the page spans.
    """
    body_size = compute_body_font_size(spans)
    return [classify_span(s, body_size) for s in spans]
