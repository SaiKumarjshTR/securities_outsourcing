#!/usr/bin/env python3
"""
Golden-test conversion helper — runs THIS REPO'S OWN pipeline/batch_runner_deploy.py
on a single DOCX and writes the resulting SGML to disk.

Usage: python3 _convert.py <input.docx> <output.sgm>

Each invocation is meant to run in its own process (invoked via subprocess by
_generate.py / test_golden.py) for full isolation between test cases — the
pipeline module has module-level singletons (_DETERMINISTIC_FIXER, etc.) that
are only proven safe for one conversion per process.
"""
import os
import sys

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(os.path.dirname(_THIS_DIR))
_PIPELINE_DIR = os.path.join(_REPO_ROOT, "pipeline")

# Same WSL2 runtime data locations used throughout dev/testing this session.
# Override via real env vars if running in a different environment (CI, etc.).
os.environ.setdefault("VENDOR_SGMS_DIR", "/mnt/c/Temp/juri_extract/juri")
os.environ.setdefault("KEYING_RULES_PATH", "/opt/sgml-pipeline/data/COMPLETE_KEYING_RULES_UPDATED.txt")
os.environ.setdefault("RAG_PERSIST_DIR", "/opt/sgml-pipeline/data/chroma_db")

sys.path.insert(0, _PIPELINE_DIR)


def convert(docx_path: str) -> str:
    """Run extract -> tag -> (LLM if ambiguous) -> generate -> deterministic-fix."""
    from batch_runner_deploy import (
        CompleteDOCXExtractor, PatternBasedTagger, SGMLGenerator,
        SequentialSGMLLayer, _DETERMINISTIC_FIXER,
        RAGManager, Anthropic, PATHS, RAG_CONFIG, KEYING_SPECIFICATIONS,
    )

    stem = os.path.splitext(os.path.basename(docx_path))[0]

    extractor = CompleteDOCXExtractor(docx_path)
    doc_data = extractor.extract_complete_document()
    metadata = doc_data["metadata"]
    paragraphs = doc_data["paragraphs"]
    content = doc_data["content"]

    tagger = PatternBasedTagger()
    confirmed, ambiguous = tagger.tag_paragraphs(paragraphs)

    for para in confirmed + ambiguous:
        para.inline_formatting = tagger.extract_inline_formatting(para)
        para.docx_formatting = para.inline_formatting.copy()

    if ambiguous:
        client = Anthropic()
        rag = None
        if RAG_CONFIG.get("enabled", False):
            try:
                rag = RAGManager(
                    keying_specs_path=PATHS["keying_rules"],
                    persist_dir=RAG_CONFIG["persist_dir"],
                    vendor_sgms=RAG_CONFIG["vendor_sgms"],
                    n_rules=RAG_CONFIG["n_rules"],
                    n_examples=RAG_CONFIG["n_examples"],
                )
                rag.initialize()
            except Exception:
                rag = None
        llm_layer = SequentialSGMLLayer(client, keying_specs=KEYING_SPECIFICATIONS, rag_manager=rag)
        ambiguous = llm_layer.process_ambiguous_paragraphs(
            ambiguous, full_paragraphs=paragraphs, exclude_vendor_source=f"{stem}.sgm"
        )

    all_paragraphs = sorted(confirmed + ambiguous, key=lambda p: p.index)
    para_map = {p.index: p for p in all_paragraphs}
    for item in content:
        if item["type"] == "paragraph":
            orig = item["data"]
            if orig.index in para_map:
                item["data"] = para_map[orig.index]

    gen = SGMLGenerator()
    gen.use_container_blocks = True
    sgml_content = gen.generate_sgml(metadata, content)
    sgml_content, _changes = _DETERMINISTIC_FIXER.fix(sgml_content)
    return sgml_content


if __name__ == "__main__":
    docx_arg = sys.argv[1]
    out_arg = sys.argv[2]
    result = convert(docx_arg)
    with open(out_arg, "w", encoding="utf-8") as fh:
        fh.write(result)
    print(f"Wrote {len(result):,} chars to {out_arg}")
