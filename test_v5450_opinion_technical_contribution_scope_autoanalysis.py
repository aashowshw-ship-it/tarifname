from pathlib import Path

from app_core import validate_gorus_analysis
from gorus_audit import validate_opinion_payload
from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES

ROOT = Path(__file__).resolve().parent


def test_version_v5450_and_rules():
    assert APP_VERSION == "v5.4.51"
    assert RULESET_VERSION == "2026-09-04.v41"
    low = GORUS_RULES.casefold()
    assert "teknik katkı" in low
    assert "yalnız x kategorisi" in low
    assert "ikinci teknik analiz" in low and "otomatik" in low


def test_analysis_requires_source_based_technical_contribution_fields():
    spec = "İstem 1 teknik zincir. Tarifname dayanağı burada birebir geçmektedir."
    analysis = {
        "amendment_required": False,
        "amendments": [],
        "technical_contributions": [{
            "claim_number": "1",
            "feature": "teknik zincir",
            "technical_effect": "ölçüm kararlılığı",
            "basis_quote": "Tarifname dayanağı burada birebir geçmektedir.",
            "defence_priority": "high",
        }],
        "description_prior_art_updates": [],
    }
    validate_gorus_analysis(analysis, spec)


def test_x_only_combined_section_is_rejected():
    op = {
        "application_no":"1", "applicant":"A", "reference":"R",
        "cited_documents":[
            {"label":"D1","number":"","category":"X"},
            {"label":"D2","number":"","category":"X"},
        ],
        "sections":[
            {"label":"D1","blocks":[{"type":"paragraph","text":"Teknik fark, teknik etki, teknik problem ve motivasyon değerlendirmesi yeterince ayrıntılıdır. "*20}],"novelty_paragraphs":["Yenilik açıklaması."],"inventive_step_paragraphs":["Buluş basamağı açıklaması."*40]},
            {"label":"D2","blocks":[{"type":"paragraph","text":"Teknik fark, teknik etki, teknik problem ve yönlendirme değerlendirmesi yeterince ayrıntılıdır. "*20}],"novelty_paragraphs":["Yenilik açıklaması."],"inventive_step_paragraphs":["Buluş basamağı açıklaması."*40]},
        ],
        "combined_assessment":{"heading":"","paragraphs":[]},
    }
    validate_opinion_payload(op, "D1 X. D2 X. Buluş basamağı itirazı ayrı ayrı.", "teknik tarifname")


def test_second_analysis_button_removed_and_auto_trigger_present():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert 'st.button("2. Raporu, istemleri ve savunma dokümanlarını teknik olarak analiz et"' not in src
    assert "analysis_upload_signature" in src
    assert "Savunma dokümanları yüklenir yüklenmez ikinci teknik analiz otomatik çalışır" in src
