from __future__ import annotations

import copy
import pytest

from gorus_audit import validate_opinion_long_form_depth, validate_customer_defence_point_coverage
from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES


def words(n: int, token: str = "teknik") -> str:
    return " ".join([token] * n) + "."


def long_opinion() -> dict:
    xpars = [
        words(180, "fark"),
        words(180, "etki"),
        words(180, "problem"),
        words(180, "yönlendirme"),
    ]
    ypars = [
        words(230, "kombinasyon"),
        words(230, "işlevsel"),
        words(230, "motivasyon"),
        words(230, "değişiklik"),
    ]
    conclusion = [words(620, "sonuç")]
    return {
        "cited_documents": [
            {"label": "D1", "category": "X"},
            {"label": "D2", "category": "Y"},
            {"label": "D3", "category": "Y"},
        ],
        "sections": [
            {
                "label": "D1",
                "blocks": [{"type": "paragraph", "text": p} for p in xpars],
                "novelty_paragraphs": [],
                "inventive_step_paragraphs": [],
            },
            {"label": "D2", "blocks": [{"type": "paragraph", "text": "Kısa objektif D2 tanıtımı."}]},
            {"label": "D3", "blocks": [{"type": "paragraph", "text": "Kısa objektif D3 tanıtımı."}]},
        ],
        "combined_assessment": {
            "groups": [
                {
                    "group": "Y",
                    "labels": ["D2", "D3"],
                    "heading": "D2 ve D3 Dokümanları Birlikte Değerlendirildiğinde",
                    "paragraphs": ypars,
                }
            ]
        },
        "conclusion": conclusion,
        "customer_point_ids_used": ["C1"],
    }


def test_release_version_and_binding_rules():
    assert APP_VERSION == "v5.4.72"
    assert RULESET_VERSION == "2026-09-18.v64"
    low = GORUS_RULES.casefold()
    assert "en az 2200 kelime" in low
    assert "en az 700 kelimelik" in low
    assert "en az 900 kelimelik" in low
    assert "customer_defence_points" in GORUS_RULES
    assert "customer_point_ids_used" in GORUS_RULES


def test_long_form_gate_accepts_substantive_long_opinion():
    validate_opinion_long_form_depth(long_opinion(), "Yenilik ve buluş basamağı")


def test_long_form_gate_rejects_short_x_defence():
    op = long_opinion()
    op["sections"][0]["blocks"] = [{"type": "paragraph", "text": words(100)}] * 4
    with pytest.raises(ValueError, match="X savunması"):
        validate_opinion_long_form_depth(op, "Buluş basamağı")


def test_long_form_gate_rejects_short_combination():
    op = long_opinion()
    op["combined_assessment"]["groups"][0]["paragraphs"] = [words(100)] * 4
    with pytest.raises(ValueError, match="Y kombinasyon"):
        validate_opinion_long_form_depth(op, "Buluş basamağı")


def customer_preanalysis() -> dict:
    return {
        "customer_defence_points": [
            {
                "id": "C1",
                "source_quote": "Müşteri, tandem hücre ile lityum-hava bataryasının birlikte çalışmasının önemli olduğunu belirtmektedir.",
                "point": "tandem hücre ve lityum-hava bataryasının özel işlevsel birlikteliği",
                "technical_effect": "enerji hasadı ile yüksek yoğunluklu depolamanın birlikte yönetilmesi",
                "basis_quote": "tarifname dayanağı",
                "coverage_terms": ["tandem hücre", "lityum-hava bataryası", "işlevsel birliktelik"],
                "use_required": True,
                "omission_reason": "",
            }
        ]
    }


def test_customer_coverage_gate_requires_used_id_and_visible_content():
    op = long_opinion()
    op["sections"][0]["blocks"][0]["text"] += " tandem hücre ile lityum-hava bataryası arasındaki işlevsel birliktelik açıklanır."
    customer = "Müşteri, tandem hücre ile lityum-hava bataryasının birlikte çalışmasının önemli olduğunu belirtmektedir."
    validate_customer_defence_point_coverage(op, customer_preanalysis(), customer)

    missing = copy.deepcopy(op)
    missing["customer_point_ids_used"] = []
    with pytest.raises(ValueError, match="zorunlu müşteri"):
        validate_customer_defence_point_coverage(missing, customer_preanalysis(), customer)


def test_customer_coverage_gate_rejects_id_only_without_technical_content():
    op = long_opinion()
    customer = "Müşteri, tandem hücre ile lityum-hava bataryasının birlikte çalışmasının önemli olduğunu belirtmektedir."
    with pytest.raises(ValueError, match="görüş metninde görünür"):
        validate_customer_defence_point_coverage(op, customer_preanalysis(), customer)
