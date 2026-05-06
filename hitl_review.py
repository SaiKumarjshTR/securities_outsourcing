"""
hitl_review.py — TR SGML Validator  |  Human-in-the-Loop Review UI
───────────────────────────────────────────────────────────────────
Streamlit app for reviewing validator output, editing SGML inline,
and recording accept/reject decisions.

Key features (v2):
  • Highlighted SGML view  — problem lines coloured red/orange/yellow
  • Actionable Fixes panel — exact line numbers, before/after diffs
  • Auto-fix D6            — apply all encoding fixes with one click
  • PDF page navigation    — jump to evidence page for each fix
  • HITL decision logging  — records to hitl_decisions.jsonl

Run:
    cd "<deployment-folder>"
    streamlit run hitl_review.py

Requirements:
    pip install streamlit pymupdf pdfplumber
"""

from __future__ import annotations

import html
import json
import sys
import tempfile
from pathlib import Path
from datetime import datetime

import streamlit as st
from config import DECISIONS_FILE

# ── path setup ────────────────────────────────────────────────────────────────
sys.path.insert(0, str(Path(__file__).parent))
from validator.validator_main import validate, ValidationReport
from validator.core.diff_generator import (
    generate_fixes,
    get_highlight_map,
    apply_auto_fixes,
    ActionableFix,
)

# ── page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="TR SGML Validator — HITL Review",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── constants ─────────────────────────────────────────────────────────────────
DECISION_COLOURS = {
    "ACCEPT":               "#1a7f37",
    "ACCEPT_WITH_WARNINGS": "#9a6700",
    "REVIEW":               "#0550ae",
    "REJECT":               "#cf222e",
}

SEVERITY_ICONS = {"critical": "🔴", "major": "🟠", "minor": "🟡", "info": "🔵"}
DIM_ICONS = {
    "L2": "🛡️", "D2": "🏷️", "D3": "📝", "D4": "📋",
    "D5": "🔀", "D6": "🔤", "D7": "🗂️",
}
SEVERITY_COLOURS = {"critical": "#ffcccc", "major": "#ffe4b5", "minor": "#fffacd"}

SCORE_BARS = {
    "l1_content_fidelity":  ("L1 Content",    35, "#4c8bf5"),
    "l2_structural":        ("L2 Structural",  40, "#34a853"),
    "l3_corpus_pattern":    ("L3 Corpus",      25, "#fbbc04"),
    "l4_source_comparison": ("L4 Source",      30, "#ea4335"),
}

# DECISIONS_FILE is imported from config.py — set via DECISIONS_FILE env var


# ── helpers ───────────────────────────────────────────────────────────────────
def _colour_badge(decision: str) -> str:
    colour = DECISION_COLOURS.get(decision, "#666")
    return (
        f'<span style="background:{colour};color:white;padding:3px 12px;'
        f'border-radius:4px;font-weight:bold;font-size:1rem">{decision}</span>'
    )


def _score_bar(label: str, score: float, max_score: float, colour: str) -> None:
    pct = min(100.0, score / max_score * 100) if max_score else 0
    st.markdown(
        f"""
        <div style="margin-bottom:6px">
          <div style="display:flex;justify-content:space-between;font-size:.85rem">
            <span>{label}</span><span>{score:.1f}/{max_score} ({pct:.0f}%)</span>
          </div>
          <div style="background:#e0e0e0;border-radius:4px;height:10px">
            <div style="background:{colour};width:{pct}%;height:10px;border-radius:4px"></div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _save_decision(record: dict) -> None:
    with DECISIONS_FILE.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


def _load_past_decisions() -> list[dict]:
    if not DECISIONS_FILE.exists():
        return []
    records = []
    for line in DECISIONS_FILE.read_text(encoding="utf-8").splitlines():
        try:
            records.append(json.loads(line))
        except json.JSONDecodeError:
            pass
    return records


# ── Highlighted SGML renderer ─────────────────────────────────────────────────

def _render_sgml_highlighted(
    sgml_text: str,
    highlight_map: dict[int, str],
    focus_line: int = 0,
) -> str:
    """
    Render SGML as an HTML block with:
      - line numbers
      - coloured background on problem lines (red/orange/yellow per severity)
      - optional focus_line: shows ±12 lines around a clicked fix, target line in bright yellow
      - monospace font, scrollable
    """
    lines = sgml_text.splitlines()
    total = len(lines)

    # If a focus line is active, restrict view to ±12 lines around it
    if focus_line > 0:
        start = max(1, focus_line - 12)
        end   = min(total, focus_line + 12)
        render_lines = [(i, lines[i - 1]) for i in range(start, end + 1)]
        window_note  = (
            f'<div style="font-size:11px;color:#6b7280;padding:3px 6px;'
            f'background:#f0f9ff;border-bottom:1px solid #bae6fd">'
            f'Showing lines {start}–{end} of {total} — '
            f'<b style="color:#1d4ed8">line {focus_line} highlighted</b></div>'
        )
    else:
        render_lines = [(i, lines[i - 1]) for i in range(1, total + 1)]
        window_note  = ""

    html_parts = [
        '<div style="'
        'font-family:\'Courier New\',Courier,monospace;font-size:11.5px;'
        'overflow:auto;max-height:540px;border:1px solid #d0d0d0;'
        'border-radius:4px;background:#fafafa;'
        'white-space:pre;line-height:1.5">',
        window_note,
        '<div style="padding:6px 8px">',
    ]
    for i, line in render_lines:
        escaped_line = html.escape(line)
        ln_span = f'<span style="color:#aaa;user-select:none;margin-right:6px">{i:4d} │</span>'
        if focus_line > 0 and i == focus_line:
            # Bright focus highlight — overrides severity colour
            html_parts.append(
                f'<span id="sgml-focus-line" style="display:block;background:#fef08a;'
                f'border-left:4px solid #eab308;padding-left:2px">'
                f'{ln_span}{escaped_line}</span>'
            )
        elif i in highlight_map:
            colour = highlight_map[i]
            html_parts.append(
                f'<span style="display:block;background:{colour}">'
                f'{ln_span}{escaped_line}</span>'
            )
        else:
            html_parts.append(
                f'<span style="display:block">{ln_span}{escaped_line}</span>'
            )
    html_parts.append("</div></div>")
    return "".join(html_parts)


# ── Fixes panel ───────────────────────────────────────────────────────────────

def _render_fixes_panel(
    fixes: list[ActionableFix],
    sgml_path: Path,
    current_sgml: str,
) -> str:
    """
    Render the actionable fixes panel.
    Returns the (possibly modified) SGML if auto-fixes were applied.
    """
    if not fixes:
        st.success("✅ No actionable fixes found — document looks correct.")
        return current_sgml

    # Summary counts
    auto_fixable = [f for f in fixes if f.auto_fixable]
    manual_fixes = [f for f in fixes if not f.auto_fixable]

    col_a, col_b, col_c = st.columns(3)
    col_a.metric("Total fixes found", len(fixes))
    col_b.metric("🤖 Auto-fixable (1-click)", len(auto_fixable))
    col_c.metric("✏️ Manual review needed", len(manual_fixes))

    # ── Apply all auto-fixes button ────────────────────────────────────────────
    if auto_fixable:
        if st.button(
            f"⚡ Apply all {len(auto_fixable)} auto-fixes (D6 encoding)",
            type="primary",
            help="Applies all safe, deterministic fixes (Unicode → SGML entities). "
                 "Encoding fixes are always safe to apply automatically.",
        ):
            corrected, count = apply_auto_fixes(current_sgml, fixes)
            st.session_state["sgml_content"] = corrected
            sgml_path.write_text(corrected, encoding="utf-8")
            st.success(f"✅ Applied {count} auto-fix(es). File saved. Re-validate to confirm.")
            current_sgml = corrected

    st.markdown("---")

    # ── Individual fixes ───────────────────────────────────────────────────────
    dim_order = ["L2", "D6", "D7", "D2", "D5", "D3", "D4"]
    fixes_by_dim: dict[str, list[ActionableFix]] = {}
    for fix in fixes:
        fixes_by_dim.setdefault(fix.dimension, []).append(fix)

    for dim in dim_order:
        dim_fixes = fixes_by_dim.get(dim, [])
        if not dim_fixes:
            continue

        dim_icon = DIM_ICONS.get(dim, "")
        _dim_desc = {
            "L2": "Structural — tag schema, nesting rules, entities, table structure",
            "D2": "Tag accuracy — bold, italic, heading styles correctly tagged",
            "D3": "Text accuracy — paragraph text matches source PDF",
            "D4": "Completeness — all tables, footnotes & sections present",
            "D5": "Order — sections & paragraphs in correct sequence",
            "D6": "Encoding — special characters converted to SGML entities",
            "D7": "Metadata — title, date, doc-number match source",
        }.get(dim, "")
        st.markdown(
            f"#### {dim_icon} {dim} — {len(dim_fixes)} fix(es) "
            f"<small style='color:#9ca3af;font-weight:400'>{_dim_desc}</small>",
            unsafe_allow_html=True,
        )

        for idx, fix in enumerate(dim_fixes):
            sev_icon = SEVERITY_ICONS.get(fix.severity, "⚪")
            sev_colour = SEVERITY_COLOURS.get(fix.severity, "#fffacd")
            loc = f"Line {fix.line_number}" if fix.line_number > 0 else "Location unknown"
            conf_badge = {
                "high":   "🟢 High confidence",
                "medium": "🟡 Medium confidence",
                "low":    "🔴 Low confidence",
            }.get(fix.confidence, fix.confidence)
            auto_badge = "⚡ Auto-fixable" if fix.auto_fixable else "✏️ Manual"

            label = f"{sev_icon} [{fix.severity.upper()}] {loc} — {fix.description[:70]}"
            with st.expander(label, expanded=fix.severity == "critical"):

                st.markdown(
                    f'<div style="background:{sev_colour};padding:6px 10px;'
                    f'border-radius:4px;margin-bottom:8px;font-size:.9rem">'
                    f'<b>{fix.description}</b></div>',
                    unsafe_allow_html=True,
                )

                info_cols = st.columns(3)
                info_cols[0].caption(f"📍 {loc}")
                info_cols[1].caption(conf_badge)
                info_cols[2].caption(auto_badge)

                if fix.context_before:
                    st.markdown("**Current SGML (problem highlighted ►):**")
                    st.code(fix.context_before, language="xml")

                if fix.suggested_fix:
                    st.markdown("**Suggested fix:**")
                    st.code(fix.suggested_fix, language="xml")

                if fix.pdf_evidence:
                    st.info(
                        f"📄 **PDF evidence"
                        f"{f' (page {fix.pdf_page})' if fix.pdf_page > 0 else ''}:** "
                        f"{fix.pdf_evidence}"
                    )

                # ── Jump-to-line button ─────────────────────────────────────────
                if fix.line_number > 0:
                    if st.button(
                        f"📍 Go to line {fix.line_number} in SGML",
                        key=f"jump_{dim}_{idx}_{fix.line_number}",
                        help="Scrolls the SGML preview to this line",
                    ):
                        st.session_state["jump_to_line"] = fix.line_number

                # Individual apply button for auto-fixable fixes
                if fix.auto_fixable and fix._fix_old and fix._fix_new:
                    if st.button(
                        f"⚡ Apply this fix",
                        key=f"apply_{dim}_{idx}_{fix.line_number}",
                    ):
                        if fix._fix_old in current_sgml:
                            current_sgml = current_sgml.replace(
                                fix._fix_old, fix._fix_new, 1
                            )
                            st.session_state["sgml_content"] = current_sgml
                            sgml_path.write_text(current_sgml, encoding="utf-8")
                            st.success(
                                f"Fix applied on line {fix.line_number}. "
                                f"Re-validate to confirm."
                            )
                        else:
                            st.warning(
                                "Could not apply fix — text may have already been corrected."
                            )

        st.markdown("")  # spacer between dimensions

    return current_sgml


# ── sidebar ───────────────────────────────────────────────────────────────────
def _sidebar() -> tuple[Path | None, Path | None]:
    st.sidebar.header("📂 Current Document")

    _auto_sgml:      str   | None = st.session_state.get("last_sgml_text")
    _auto_pdf:       bytes | None = st.session_state.get("last_pdf_bytes")
    _auto_pdf_name:  str          = st.session_state.get("last_pdf_name")  or "source.pdf"
    _auto_sgml_name: str          = st.session_state.get("last_sgml_name") or "pipeline_output.sgm"

    if _auto_sgml:
        st.sidebar.markdown(
            f'<div style="background:#f0f9ff;border:1px solid #bae6fd;border-radius:6px;'
            f'padding:8px 10px;font-size:0.82em;color:#0c4a6e">'
            f'<b>📌 Loaded from pipeline</b><br>'
            f'<span style="color:#0369a1">{_auto_pdf_name}</span><br>'
            f'<span style="color:#0369a1">{_auto_sgml_name}</span></div>',
            unsafe_allow_html=True,
        )
    else:
        st.sidebar.warning("No pipeline output loaded.\nRun a conversion on the main page first.")

    sgml_path = pdf_path = None

    # ── Resolve from pipeline session state ───────────────────────────────────
    if _auto_sgml:
        _ck = f"_hitl_sgml_{_auto_sgml_name}"
        if _ck not in st.session_state or not Path(str(st.session_state[_ck])).exists():
            _td = tempfile.mkdtemp(prefix="hitl_auto_")
            _tp = Path(_td) / _auto_sgml_name
            _tp.write_text(_auto_sgml, encoding="utf-8")
            st.session_state[_ck] = str(_tp)
        sgml_path = Path(str(st.session_state[_ck]))

    if _auto_pdf:
        _ck2 = f"_hitl_pdf_{_auto_pdf_name}"
        if _ck2 not in st.session_state or not Path(str(st.session_state[_ck2])).exists():
            _td2 = tempfile.mkdtemp(prefix="hitl_pdf_")
            _tp2 = Path(_td2) / _auto_pdf_name
            _tp2.write_bytes(_auto_pdf)
            st.session_state[_ck2] = str(_tp2)
        pdf_path = Path(str(st.session_state[_ck2]))

    return sgml_path, pdf_path


# ── main review panel ─────────────────────────────────────────────────────────
def _render_report(
    report: ValidationReport,
    sgml_path: Path,
    pdf_path: Path,
) -> None:
    d = report.to_dict()
    scores = d["scores"]

    # ── Header ────────────────────────────────────────────────────────────────
    col_title, col_badge = st.columns([4, 1])
    with col_title:
        st.subheader(f"📄 {sgml_path.name}")
    with col_badge:
        st.markdown(_colour_badge(report.decision), unsafe_allow_html=True)

    if report.critical_failures:
        st.error("🚨 Critical failures: " + " · ".join(report.critical_failures))

    st.markdown("---")

    # ── Score breakdown ────────────────────────────────────────────────────────
    with st.expander("📊 Score breakdown", expanded=False):
        for key, (label, mx, colour) in SCORE_BARS.items():
            _score_bar(label, scores[key], mx, colour)

        if d.get("l4_details"):
            l4 = d["l4_details"]
            st.markdown("**L4 sub-scores:**")
            cols = st.columns(6)
            sub = [
                ("D2 Tagging",   l4["tagging_score"],      5),
                ("D3 Text",      l4["text_score"],          8),
                ("D4 Complete",  l4["completeness_score"],  7),
                ("D5 Order",     l4["ordering_score"],      4),
                ("D6 Encoding",  l4["encoding_score"],      3),
                ("D7 Metadata",  l4["metadata_score"],      3),
            ]
            for col, (lbl, val, mx) in zip(cols, sub):
                pct = val / mx * 100 if mx else 0
                col.metric(lbl, f"{val:.1f}/{mx}", f"{pct:.0f}%")
            if l4.get("text_coverage") is not None:
                st.caption(f"Text coverage: {l4['text_coverage']:.0%}")

        if d.get("l1_details") and report.l1:
            total = getattr(report.l1, "total_missing_para_count", 0)
            if total > 0:
                st.warning(
                    f"📌 L1: **{total} paragraph(s)** from PDF not found in SGML — "
                    "check fix cards below for details."
                )

    # ── Generate actionable fixes ─────────────────────────────────────────────
    l4_raw = report.l4   # L4Result object (has d2_untagged_bold etc.)
    if "sgml_content" not in st.session_state:
        st.session_state["sgml_content"] = sgml_path.read_text(
            encoding="utf-8", errors="replace"
        )

    current_sgml: str = st.session_state["sgml_content"]

    fixes: list[ActionableFix] = []
    if l4_raw is not None:
        fixes = generate_fixes(current_sgml, l4_raw, l2_result=report.l2)

    highlight_map = get_highlight_map(fixes)

    # ── Actionable Fixes Panel ────────────────────────────────────────────────
    fix_count = len(fixes)
    auto_count = sum(1 for f in fixes if f.auto_fixable)
    with st.expander(
        f"🔧 Actionable Fixes ({fix_count} found — {auto_count} auto-fixable)",
        expanded=fix_count > 0,
    ):
        current_sgml = _render_fixes_panel(fixes, sgml_path, current_sgml)

    st.markdown("---")

    # ── Side-by-side: PDF viewer (left) | SGML editor (right) ──────────────────
    if "pdf_page_n" not in st.session_state:
        st.session_state["pdf_page_n"] = 1

    # Remove Streamlit's default top-padding on nested columns so both headers
    # sit at exactly the same vertical position.
    st.markdown(
        """
        <style>
        /* Align both panel headers to the same top baseline */
        div[data-testid="column"] > div:first-child {
            padding-top: 0 !important;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] {
            padding-top: 0 !important;
        }
        /* Remove extra gap that number_input adds above itself */
        div[data-testid="stNumberInput"] {
            margin-top: 0.35rem !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    col_pdf, col_sgml = st.columns(2)

    with col_pdf:
        try:
            import fitz
            _doc = fitz.open(str(pdf_path))
            _total_pages = len(_doc)

            # ── Header row: title + page control at same baseline ─────────────
            _pdf_title_col, _pdf_ctrl_col = st.columns([3, 2])
            with _pdf_title_col:
                st.markdown(
                    "<div style='margin-bottom:0;padding-bottom:0'>"
                    "<h3 style='margin:0;padding:0;line-height:2.2rem'>📄 Source PDF</h3>"
                    "</div>",
                    unsafe_allow_html=True,
                )
            with _pdf_ctrl_col:
                _page_n = st.number_input(
                    "Page",
                    min_value=1,
                    max_value=_total_pages,
                    value=st.session_state["pdf_page_n"],
                    step=1,
                    key="pdf_page_input",
                    label_visibility="collapsed",
                )
            st.session_state["pdf_page_n"] = _page_n

            _page = _doc[_page_n - 1]
            _pix = _page.get_pixmap(dpi=130)
            st.image(_pix.tobytes("png"), use_container_width=True)
            st.caption(f"Page {_page_n} of {_total_pages}")
            _doc.close()

        except ImportError:
            st.warning("PyMuPDF not installed — PDF preview unavailable.")
        except Exception as e:
            st.error(f"PDF render error: {e}")

    with col_sgml:
        # ── Highlighted view ──────────────────────────────────────────────────
        problem_count = len(highlight_map)
        colour_legend = (
            " &nbsp;🔴 red=critical &nbsp;🟠 orange=major &nbsp;🟡 yellow=minor"
            if highlight_map else ""
        )

        # Focus line state
        focus_line = st.session_state.get("jump_to_line", 0)

        _sgml_hdr_col, _sgml_clear_col = st.columns([5, 1])
        with _sgml_hdr_col:
            st.markdown(
                "<div style='margin-bottom:0;padding-bottom:0'>"
                f"<h3 style='margin:0;padding:0;line-height:2.2rem'>✏️ SGML"
                f"&nbsp;<small style='font-size:.8rem;font-weight:normal;color:#666'>"
                f"{problem_count} problem line(s) highlighted{colour_legend}</small></h3>"
                "</div>",
                unsafe_allow_html=True,
            )
        with _sgml_clear_col:
            if focus_line > 0:
                if st.button("✖ Clear focus", key="clear_jump", help="Show full SGML"):
                    st.session_state["jump_to_line"] = 0
                    focus_line = 0

        # Highlighted read-only view
        st.markdown(
            _render_sgml_highlighted(current_sgml, highlight_map, focus_line=focus_line),
            unsafe_allow_html=True,
        )

        st.markdown(
            "**Edit SGML** <small style='color:#888;font-size:0.8em'>"
            "(line numbers shown above for reference)</small>",
            unsafe_allow_html=True,
        )

        # ── Line-numbered editor ──────────────────────────────────────────────
        _sgml_lines = current_sgml.splitlines()
        _n_lines = len(_sgml_lines)
        _ln_html = "".join(
            f'<div style="height:21px;line-height:21px;color:#94a3b8;'
            f'font-size:12px;text-align:right;padding-right:3px">{i}</div>'
            for i in range(1, _n_lines + 1)
        )
        _ln_col, _ta_col = st.columns([1, 16])
        with _ln_col:
            st.markdown(
                f'<div style="font-family:\'Courier New\',monospace;'
                f'height:420px;overflow:hidden;background:#f8fafc;'
                f'border:1px solid #d1d5db;border-right:none;'
                f'border-radius:4px 0 0 4px;padding:6px 3px 0 3px;'
                f'user-select:none">{_ln_html}</div>',
                unsafe_allow_html=True,
            )
        with _ta_col:
            edited = st.text_area(
                label="Edit SGML",
                value=current_sgml,
                height=420,
                key="sgml_editor",
                label_visibility="collapsed",
            )
        if edited != current_sgml:
            st.session_state["sgml_content"] = edited
            current_sgml = edited

        dl_col, save_col = st.columns(2)
        with dl_col:
            st.download_button(
                "⬇️ Download SGML",
                data=current_sgml.encode("utf-8"),
                file_name=sgml_path.name,
                mime="text/plain",
            )
        with save_col:
            if st.button("💾 Save to disk", key="save_sgml"):
                sgml_path.write_text(current_sgml, encoding="utf-8")
                st.success(f"Saved → {sgml_path}")
                # Re-run validation on the saved SGML so scores + fixes reflect edits.
                try:
                    from validator.validator_main import validate
                    refreshed = validate(
                        sgml_path,
                        pdf_path if pdf_path and pdf_path.exists() else None,
                    )
                    st.session_state["validation_report"] = refreshed
                    st.info("🔄 Validation refreshed after save.")
                    st.rerun()
                except Exception as _rv_exc:
                    st.warning(f"Re-validation failed: {_rv_exc}")

    # ── Validator issues (raw) ────────────────────────────────────────────────
    with st.expander(
        f"⚠️ All validator issues ({len(report.all_issues)})",
        expanded=False,
    ):
        if not report.all_issues:
            st.success("No issues found.")
        else:
            by_level: dict[str, list] = {}
            for iss in report.all_issues:
                by_level.setdefault(iss.get("level", "?"), []).append(iss)
            for level in ["L1", "L2", "L3", "L4"]:
                issues = by_level.get(level, [])
                if not issues:
                    continue
                st.markdown(f"**{level}** — {len(issues)} issue(s)")
                for iss in sorted(
                    issues,
                    key=lambda x: {"critical": 0, "major": 1, "minor": 2}.get(
                        x.get("severity", ""), 3
                    ),
                ):
                    icon = SEVERITY_ICONS.get(iss.get("severity", ""), "⚪")
                    st.markdown(
                        f"{icon} **[{iss.get('severity','').upper()}]** "
                        f"`{iss.get('category','')}` — {iss.get('description','')}",
                    )

    if report.warnings:
        with st.expander(f"ℹ️ Validator warnings ({len(report.warnings)})", expanded=False):
            for w in report.warnings:
                st.caption(w)

    st.markdown("---")

    # ── HITL Decision ─────────────────────────────────────────────────────────
    st.markdown("### 🧑‍⚖️ Human Decision")

    # Warn if any unresolved critical-severity fixes remain
    open_criticals = [f for f in fixes if f.severity == "critical"]
    if open_criticals:
        st.error(
            f"⚠️ **{len(open_criticals)} unresolved critical issue(s)** — "
            "resolving these before accepting is strongly recommended. "
            "Selecting ACCEPT with open critical issues requires an explicit reviewer note."
        )

    dcol1, dcol2 = st.columns([3, 1])
    with dcol1:
        reviewer_note = st.text_area(
            "Reviewer note (required for REJECT / REVIEW overrides)",
            height=80,
            key="reviewer_note",
            placeholder=(
                "e.g. 'D3 low coverage is acceptable — document is an alert with minimal text'"
            ),
        )
    with dcol2:
        human_decision = st.radio(
            "Override decision",
            options=[
                "(keep validator decision)",
                "ACCEPT",
                "ACCEPT_WITH_WARNINGS",
                "REVIEW",
                "REJECT",
            ],
            key="human_decision",
        )

    if st.button("✅ Record Decision", type="primary"):
        final = (
            report.decision
            if human_decision == "(keep validator decision)"
            else human_decision
        )
        # Block ACCEPT when critical issues are open unless reviewer has noted justification
        if final in ("ACCEPT", "ACCEPT_WITH_WARNINGS") and open_criticals and not reviewer_note.strip():
            st.error(
                f"⛔ Cannot record ACCEPT with {len(open_criticals)} open critical issue(s) "
                "and no reviewer note. Please resolve the issues or add a justification note."
            )
            return
        record = {
            "timestamp": datetime.utcnow().isoformat(),
            "sgml_file": str(sgml_path),
            "pdf_file": str(pdf_path),
            "validator_decision": report.decision,
            "human_decision": final,
            "normalised_score": round(scores["normalised"], 1),
            "reviewer_note": reviewer_note,
            "fixes_found": len(fixes),
            "auto_fixes_available": auto_count,
            "scores": scores,
        }
        _save_decision(record)
        colour = DECISION_COLOURS.get(final, "#666")
        st.markdown(
            f'<div style="background:{colour};color:white;padding:10px;'
            f'border-radius:6px;font-size:1.1rem">'
            f'Decision recorded: <b>{final}</b></div>',
            unsafe_allow_html=True,
        )

    # Full JSON debug
    with st.expander("🔍 Full validation JSON", expanded=False):
        st.json(d)


# ── history tab ───────────────────────────────────────────────────────────────
def _render_history() -> None:
    st.header("📋 Decision History")
    records = _load_past_decisions()
    if not records:
        st.info("No decisions recorded yet.")
        return

    st.caption(f"{len(records)} decision(s) in `{DECISIONS_FILE.name}`")

    from collections import Counter
    counts = Counter(r["human_decision"] for r in records)
    cols = st.columns(4)
    for col, key in zip(
        cols, ["ACCEPT", "ACCEPT_WITH_WARNINGS", "REVIEW", "REJECT"]
    ):
        col.metric(key, counts.get(key, 0))

    st.markdown("---")
    for rec in reversed(records):
        colour = DECISION_COLOURS.get(rec["human_decision"], "#666")
        ts = rec.get("timestamp", "")[:19].replace("T", " ")
        fname = Path(rec.get("sgml_file", "?")).name
        score = rec.get("normalised_score", "?")
        note = rec.get("reviewer_note", "")
        fixes = rec.get("fixes_found", "")
        st.markdown(
            f'<span style="background:{colour};color:white;padding:2px 8px;'
            f'border-radius:3px">{rec["human_decision"]}</span> &nbsp; '
            f'`{fname}` &nbsp; score={score}'
            + (f" &nbsp; fixes={fixes}" if fixes else "")
            + f" &nbsp; {ts}"
            + (f"  \n> _{note}_" if note else ""),
            unsafe_allow_html=True,
        )


# ── main ──────────────────────────────────────────────────────────────────────
def main() -> None:
    st.title("⚖️ TR SGML Validator — HITL Review")

    tab_review, tab_history = st.tabs(["🔍 Review", "📋 History"])

    with tab_review:
        sgml_path, pdf_path = _sidebar()

        if sgml_path is None:
            st.info(
                "Upload an SGML file and its source PDF using the sidebar, "
                "or enter disk paths directly."
            )
            st.markdown(
                """
                **What this tool does:**
                - Runs all 4 validation levels (L1–L4) on the vendor SGML
                - Shows **highlighted SGML** — problem lines coloured red/orange/yellow
                - Shows **Actionable Fixes** — exact line numbers, before/after diffs
                - Lets you **auto-apply D6 encoding fixes** with one click
                - Lets you edit the SGML inline and save to disk
                - Records your human accept/reject decision to `hitl_decisions.jsonl`

                **Colour key for highlighted SGML:**
                - 🔴 Red background = critical issue on this line
                - 🟠 Orange background = major issue on this line
                - 🟡 Yellow background = minor issue on this line
                """
            )
            return

        if pdf_path is None:
            st.warning(
                "📌 SGML auto-loaded from last pipeline run. "
                "Upload the source PDF in the sidebar to begin full HITL validation."
            )
            return

        # Reset SGML content in session when a new file is loaded
        run_key = f"{sgml_path}|{pdf_path}"
        if st.session_state.get("_last_run") != run_key:
            st.session_state.pop("sgml_content", None)
            st.session_state["pdf_page_n"] = 1

        if st.session_state.get("_last_run") != run_key or st.button("🔄 Re-validate"):
            with st.spinner("Running validator (L1 → L2 → L3 → L4)…"):
                report = validate(str(sgml_path), str(pdf_path))
            st.session_state["_report"] = report
            st.session_state["_last_run"] = run_key

        report = st.session_state.get("_report")
        if report:
            _render_report(report, sgml_path, pdf_path)

    with tab_history:
        _render_history()


if __name__ == "__main__":
    main()
