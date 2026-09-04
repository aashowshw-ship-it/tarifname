from pathlib import Path

import pytest

from gorus_audit import (
    detect_ep_xy_documents,
    validate_examiner_persuasion_assessment,
    validate_opinion_narrative_rules,
    validate_opinion_payload,
)
from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES

ROOT = Path(__file__).resolve().parent


def _xy_report():
    return """
SUPPLEMENTARY EUROPEAN SEARCH REPORT
DOCUMENTS CONSIDERED TO BE RELEVANT
X XP033966287
-----
Y US 2020/145229 A1
-----
A US 2021/273931 A1
CATEGORY OF CITED DOCUMENTS
D1: XP033966287
D2: US 2020/145229 A1
D3: US 2021/273931 A1
Inventive step
D1 and D2 are relevant to the assessment.
"""


def _long_combined():
    base = (
        "The distinguishing technical difference is the claimed functional relationship between the processing stages. "
        "The technical effect is that the claimed data representation is used by the following processing stage in the defined sequence. "
        "The objective technical problem is therefore how to implement that processing relationship without changing the claimed data flow. "
        "D1 teaches a first structure and D2 teaches a different operation, but neither document provides a motivation, teaching or suggestion to alter its disclosed structure in the additional manner required by the claim. "
        "Combining the documents would require additional structural and functional modification of the data path, the processing order and the interaction of the claimed components. "
        "Those additional modifications are not disclosed as a coordinated solution in either document and are not presented as an adaptation for solving the stated technical problem. "
    )
    return base * 4


def _opinion():
    return {
        "application_no": "24223505.9",
        "applicant": "EXAMPLE INC.",
        "reference": "698199",
        "intro": "The search opinion raises an inventive-step objection. Applicant's observations are submitted below.",
        "cited_documents": [
            {"label": "D1", "number": "", "category": "X", "summary": "D1 discloses a first technical arrangement."},
            {"label": "D2", "number": "", "category": "Y", "summary": "D2 discloses a second technical arrangement."},
        ],
        "sections": [
            {
                "label": "D1",
                "blocks": [{"type": "paragraph", "text": "D1 is briefly explained and the distinguishing technical difference is identified together with its technical contribution."}],
                "novelty_heading": "",
                "novelty_paragraphs": ["D1 does not directly and unambiguously disclose the complete claimed technical combination."],
                "inventive_step_heading": "",
                "inventive_step_paragraphs": ["The technical effect and objective technical problem are addressed, and D1 provides no motivation or teaching for the additional technical modification required by the claim."],
            },
            {
                "label": "D2",
                "blocks": [{"type": "paragraph", "text": "D2 is briefly explained and its technical teaching is distinguished from the claimed arrangement."}],
                "novelty_heading": "",
                "novelty_paragraphs": [],
                "inventive_step_heading": "",
                "inventive_step_paragraphs": ["The technical effect and objective technical problem are addressed, and D2 provides no motivation or suggestion for the additional structural modification required by the claim."],
            },
        ],
        "combined_assessment": {
            "heading": "D1 and D2 Documents Considered Together",
            "paragraphs": [_long_combined()],
        },
        "conclusion": ["The claimed subject-matter therefore involves an inventive step."],
    }


def test_version_v546_and_binding_rules_present():
    assert APP_VERSION == "v5.4.50"
    assert RULESET_VERSION == "2026-09-04.v40"
    low = GORUS_RULES.casefold()
    for phrase in [
        "x kategorisindeki", "y kategorisindeki", "considered together",
        "en kapsamlı ve en güçlü", "uzman perspektifi", "geri çektirme olasılığı",
        "şekilden sonraki", "giriş kısmı normal", "birebir tarifname alıntısı",
    ]:
        assert phrase in low


def test_xy_detection_preserves_x_y_categories_and_excludes_a():
    docs = detect_ep_xy_documents(_xy_report())
    assert [(d["label"], d["category"]) for d in docs] == [("D1", "X"), ("D2", "Y")]


def test_xy_structure_x_has_novelty_y_must_not_have_novelty():
    op = _opinion()
    validate_opinion_narrative_rules(op, "Inventive step objection", "technical specification")
    bad = _opinion()
    bad["sections"][1]["novelty_paragraphs"] = ["Not disclosed."]
    with pytest.raises(ValueError, match="Y dokümanı"):
        validate_opinion_narrative_rules(bad, "Inventive step objection", "technical specification")


def test_y_combination_requires_combined_main_defence():
    op = _opinion()
    validate_opinion_payload(op, "Inventive step objection", "technical specification")
    bad = _opinion()
    bad["combined_assessment"] = {"heading": "D1 assessment", "paragraphs": ["technical effect and technical problem"]}
    with pytest.raises(ValueError, match="çoklu-doküman|birlikte"):
        validate_opinion_payload(bad, "Inventive step objection", "technical specification")



def test_x_only_documents_must_not_create_combined_section():
    op = _opinion()
    op["cited_documents"][1]["category"] = "X"
    op["combined_assessment"] = {"heading": "", "paragraphs": []}
    validate_opinion_payload(op, "Inventive step objection for D1. Separate inventive step objection for D2.", "technical specification")
    bad = _opinion()
    bad["cited_documents"][1]["category"] = "X"
    with pytest.raises(ValueError, match="X-doküman|Birlikte"):
        validate_opinion_payload(bad, "Inventive step objection for D1. Separate inventive step objection for D2.", "technical specification")

def test_final_opinion_bans_hindsight_internal_forms_and_semicolon():
    for text, pattern in [
        ("The conclusion can only be reached with hindsight.", "hindsight"),
        ("Müşteri görüş formu bu teknik farkı açıklamaktadır.", "iç-kaynak"),
        ("The technical difference is clear; the effect follows.", "noktalı virgül"),
    ]:
        op = _opinion()
        op["sections"][0]["blocks"][0]["text"] += " " + text
        with pytest.raises(ValueError, match=pattern):
            validate_opinion_narrative_rules(op, "Inventive step objection", "technical specification")


def test_examiner_persuasion_is_not_quality_score_and_is_validated():
    good = {
        "persuasion_probability": 68,
        "likely_examiner_response": "The objection may be maintained unless the technical link is accepted.",
        "strongest_points": ["Directly supported technical difference"],
        "remaining_risks": ["Technical effect may still be considered insufficient"],
        "technical_difference_focus": ["Explain the functional link more concretely"],
        "can_strengthen_without_new_matter": True,
    }
    validate_examiner_persuasion_assessment(good)
    bad = dict(good)
    bad["persuasion_probability"] = 101
    with pytest.raises(ValueError, match="0-100"):
        validate_examiner_persuasion_assessment(bad)


def test_app_runs_examiner_simulation_after_word_gates_and_one_strengthening_cycle():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert "data = build_and_gate_gorus_opinion(opinion, final_spec_name, final_spec_bytes, source_state)" in src
    assert "gorus_examiner_persuasion_prompt(" in src
    assert 'int(examiner_assessment.get("persuasion_probability", 0)) < 75' in src
    assert "gorus_examiner_strengthen_prompt(" in src
    assert "data = build_and_gate_gorus_opinion(opinion, final_spec_name, final_spec_bytes, source_state)" in src
    assert "Mevcut uzman itirazını geri çektirme olasılığı (tahmini)" in src
    assert "genel kalite puanı" in src


def test_figure_caption_bold_and_post_figure_blank_are_hard_gated():
    app_src = (ROOT / "app.py").read_text(encoding="utf-8")
    audit_src = (ROOT / "gorus_audit.py").read_text(encoding="utf-8")
    assert 'r.font.name = "Arial"; r.font.size = Pt(11); r.bold = True' in app_src
    assert "tablosundan sonra savunmadan önce fiziksel boş paragraf yok" in audit_src
    assert "şekil başlığı kalın değil" in audit_src
    assert "sayfa-satır giriş kısmı normal yazı olmalıdır" in audit_src
