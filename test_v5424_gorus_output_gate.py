from __future__ import annotations

from pathlib import Path

import pytest

from gorus_audit import validate_opinion_payload
from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES

ROOT = Path(__file__).resolve().parent


def _base_opinion():
    return {
        "application_no": "2024/000001",
        "applicant": "ÖRNEK ANONİM ŞİRKETİ",
        "reference": "699999",
        "cited_documents": [],
        "sections": [],
        "combined_assessment": {
            "heading": "Dokümanların birlikte değerlendirilmesi",
            "paragraphs": [
                "En yakın dokümana göre teknik fark, aynı teknik unsurun iki işlevi belirli bir ilişki içinde gerçekleştirmesidir. "
                "Bu teknik farkın teknik etkisi, ayrı işlevlerin ortak bir yapı üzerinden gerçekleştirilmesidir. "
                "Buna göre objektif teknik problem, bu işlevlerin ortak yapıda nasıl gerçekleştirileceğidir. "
                "İkinci dokümanda bu yönde bir motivasyon veya yönlendirme bulunmamaktadır. "
                "Uzman kişinin sonuca ulaşması için birden fazla yapısal ve işlevsel değişiklik yapması gerekir. "
                "Bu kombinasyon için önceki teknikte açıklanmayan ek yapısal ve işlevsel değişiklikler gerekir ve D1 ile D2 bu değişikliklere yönelik teknik bir yönlendirme sağlamaz. "
            ] * 4,
        },
    }


def test_version_and_ruleset_bumped_for_opinion_gate():
    assert APP_VERSION == "v5.4.50"
    assert RULESET_VERSION == "2026-09-04.v40"


def test_gorus_rules_include_language_physical_line_original_figures_and_full_output_gate():
    low = GORUS_RULES.casefold()
    for phrase in [
        "görüş dili",
        "tarifname sayfa",
        "satır",
        "birebir",
        "özgün patent",
        "birlikte değerlendirildiğinde",
        "teknik fark",
        "teknik etki",
        "objektif teknik problem",
        "motivasyon",
                "şablon",
        "render",
    ]:
        assert phrase in low


def test_active_app_wires_opinion_language_applicant_exact_line_and_figure_gates():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert 'st.selectbox("Görüş dili"' in src
    assert 'st.text_input("Başvuru sahibi (raporda yoksa girin)"' in src
    assert "annotate_quote_locations(" in src
    assert "extract_cited_original_figure_pages(" in src
    assert "validate_gorus_template_fidelity(" in src
    assert "render_gorus_docx_smoke_test(current_data)" in src
    assert 'opinion["applicant"] = source_state["applicant_override"]' in src


def test_model_is_forbidden_to_invent_page_line_numbers_and_must_write_strong_combined_section():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert "Tarifname alıntılarında sayfa/satır numarası YAZMA" in src
    assert "teknik fark → teknik etki → objektif teknik problem" in src
    assert "hindsight" in src and "kullanma" in src


def test_payload_rejects_missing_metadata():
    op = _base_opinion()
    op["applicant"] = ""
    with pytest.raises(ValueError, match="Başvuru Sahibi"):
        validate_opinion_payload(op, "Buluş basamağı. D1 ve D2 birlikte değerlendirilmiştir.", "örnek tarifname")


def test_payload_rejects_weak_inventive_step_combined_assessment():
    op = _base_opinion()
    op["cited_documents"] = [{"label":"D1","number":"","category":"X"},{"label":"D2","number":"","category":"Y"}]
    op["sections"] = [
        {"label":"D1","blocks":[{"type":"paragraph","text":"Teknik fark ve teknik etki açıklanır."}],"inventive_step_paragraphs":["Objektif teknik problem ve motivasyon değerlendirilir."]},
        {"label":"D2","blocks":[{"type":"paragraph","text":"İkinci dokümanın teknik öğretisi açıklanır."}],"inventive_step_paragraphs":["İlave teknik değişiklikler değerlendirilir."]},
    ]
    op["combined_assessment"] = {"heading": "D1 ve D2 Dokümanları Birlikte Değerlendirildiğinde", "paragraphs": ["D1 ve D2 farklıdır."]}
    with pytest.raises(ValueError, match="buluş basamağı"):
        validate_opinion_payload(op, "Buluş basamağı", "örnek tarifname")


def test_source_page_line_is_deterministic_not_llm_field():
    src = (ROOT / "gorus_audit.py").read_text(encoding="utf-8")
    assert "build_page_line_index" in src
    assert "locate_quote_page_lines" in src
    assert 'q["line_start"] = l1' in src
    assert 'q["line_end"] = l2' in src
    assert "printed line numbers" in src or "basılı satır" in src


def test_template_gate_checks_opening_blanks_and_two_blanks_before_each_figure():
    src = (ROOT / "gorus_audit.py").read_text(encoding="utf-8")
    assert "girişteki fiziksel boş paragraf" in src
    assert "iki fiziksel boş paragraf" in src
    assert "1,5 satır aralığından sapma" in src
    assert "özgün görsel yok" in src
