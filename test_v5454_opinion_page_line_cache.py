from pathlib import Path

import fitz

import gorus_audit
from gorus_audit import (
    annotate_quote_locations,
    build_page_line_index,
    clear_page_line_index_cache,
    validate_quote_locations_against_spec,
)
from rules import APP_VERSION, RULESET_VERSION

ROOT = Path(__file__).resolve().parent


def _sample_rendered_pdf() -> bytes:
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    # Printed every-fifth line anchors in the left margin.
    page.insert_text((30, 100), "5", fontsize=10)
    page.insert_text((30, 185), "10", fontsize=10)
    # Body text to be cited.
    page.insert_text((125, 117), "Alpha beta gamma", fontsize=10)
    page.insert_text((125, 134), "Delta epsilon", fontsize=10)
    data = doc.tobytes()
    doc.close()
    return data


def test_v5454_version_and_ruleset():
    assert APP_VERSION == "v5.4.78"
    assert RULESET_VERSION == "2026-09-25.v70"
    assert (ROOT / "README.md").read_text(encoding="utf-8").startswith("# Patent Atölyesi v5.4.78")


def test_page_line_index_cache_uses_exact_authoritative_pdf_hash(monkeypatch):
    clear_page_line_index_cache()
    rendered = _sample_rendered_pdf()
    calls = {"count": 0}
    original_page_lines = gorus_audit._page_lines

    def counted_page_lines(page):
        calls["count"] += 1
        return original_page_lines(page)

    monkeypatch.setattr(gorus_audit, "_page_lines", counted_page_lines)

    first = build_page_line_index("final.pdf", rendered)
    count_after_first = calls["count"]
    second = build_page_line_index("final.pdf", rendered)
    assert first == second
    assert calls["count"] == count_after_first, "Aynı otorite PDF ikinci kez fiziksel indekslenmemeli."

    changed = rendered + b"\n"
    build_page_line_index("final.pdf", changed)
    assert calls["count"] > count_after_first, "Otorite PDF baytları değişirse SHA-256 anahtarı değişmeli."


def test_page_line_index_rejects_word_as_physical_line_authority():
    import pytest
    with pytest.raises(ValueError, match="doğrulama kaynağı PDF"):
        build_page_line_index("final.docx", b"word-bytes")


def test_annotation_and_hard_gate_share_prebuilt_index_without_skipping_validation(monkeypatch):
    index = [
        {"page": 1, "line": 5, "text": "Alpha beta gamma", "y": 100.0},
        {"page": 1, "line": 6, "text": "Delta epsilon", "y": 117.0},
    ]
    opinion = {
        "sections": [
            {"blocks": [{"type": "quote", "text": "Alpha beta gamma"}]}
        ]
    }

    def should_not_rebuild(*args, **kwargs):
        raise AssertionError("Hazır fiziksel indeks varken yeniden render/index oluşturulmamalı.")

    monkeypatch.setattr(gorus_audit, "build_page_line_index", should_not_rebuild)
    annotate_quote_locations(
        opinion, "final.docx", b"spec", "Türkçe", page_line_index=index
    )
    validate_quote_locations_against_spec(
        opinion, "final.docx", b"spec", "Türkçe", page_line_index=index
    )

    quote = opinion["sections"][0]["blocks"][0]
    assert (quote["page"], quote["line_start"], quote["page_end"], quote["line_end"]) == (1, 5, 1, 5)
    assert "sayfa 1" in quote["lead"] and "satır 5" in quote["lead"]


def test_active_opinion_gate_prebuilds_one_index_and_passes_it_to_both_gates():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    start = src.index("def build_and_gate_gorus_opinion(")
    block = src[start:start + 2600]
    assert "page_line_index = build_page_line_index(authority_name, authority_bytes)" in block
    assert "annotate_quote_locations(" in block
    assert "validate_quote_locations_against_spec(" in block
    assert block.count("page_line_index=page_line_index") >= 2
