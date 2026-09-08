from pathlib import Path

import pytest

from gorus_audit import (
    detect_ep_xy_documents,
    validate_xy_scope_ai_payload,
    validate_y_combination_group_coverage,
    xy_scope_ai_fallback_prompt,
)
from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES

ROOT = Path(__file__).resolve().parent


def _tr_flat_report():
    return """
TÜRK PATENT VE MARKA KURUMU
ARAŞTIRMA RAPORU
C. İLGİLİ DOKÜMANLAR
Kategori Dokümanlar İlgili Olduğu İstem
Y1,Y2
Y1
Y2
CN103644554B (YUAN XIMING)
CN117460385A (LG INNOTEK CO LTD)
CN201360012Y (JINXIANG JIANG)
1-6
1-6
1-6
Kategorilerin Açıklaması:
“X” Buluşun yeni olmadığını tek başına gösteren doküman
“Y” Buluşun buluş basamağı içermediğini başka bir dokümanla bir araya getirildiğinde gösteren doküman
"""


def test_rev6_ruleset_and_binding_y_group_rules_present():
    assert APP_VERSION == "v5.4.44"
    assert RULESET_VERSION == "2026-09-08.v46"
    low = GORUS_RULES.casefold()
    for phrase in ["y1,y2", "gerçek kombinasyon grupları", "deterministik parser", "ai fallback", "fail-closed"]:
        assert phrase in low


def test_turkpatent_flat_y1_y2_table_is_normalized_without_losing_groups():
    docs = detect_ep_xy_documents(_tr_flat_report())
    assert [(d["label"], d["number"], d["category"]) for d in docs] == [
        ("D1", "CN103644554B", "Y"),
        ("D2", "CN117460385A", "Y"),
        ("D3", "CN201360012Y", "Y"),
    ]
    assert docs[0]["category_marker"] == "Y1,Y2"
    assert docs[0]["combination_groups"] == ["Y1", "Y2"]
    assert docs[1]["combination_groups"] == ["Y1"]
    assert docs[2]["combination_groups"] == ["Y2"]


def test_ai_fallback_is_source_validated_and_fail_closed():
    good = {
        "documents": [
            {"number": "CN103644554B", "category_marker": "Y1,Y2"},
            {"number": "CN117460385A", "category_marker": "Y1"},
            {"number": "CN201360012Y", "category_marker": "Y2"},
        ]
    }
    docs = validate_xy_scope_ai_payload(_tr_flat_report(), good)
    assert [d["category"] for d in docs] == ["Y", "Y", "Y"]

    bad_number = {"documents": [{"number": "CN999999999A", "category_marker": "Y1"}]}
    with pytest.raises(ValueError, match="raporda doğrulanamayan"):
        validate_xy_scope_ai_payload(_tr_flat_report(), bad_number)

    singleton_group = {"documents": [{"number": "CN103644554B", "category_marker": "Y1"}]}
    with pytest.raises(ValueError, match="tek dokümana bağlı"):
        validate_xy_scope_ai_payload(_tr_flat_report(), singleton_group)


def test_y1_y2_groups_require_separate_visible_pair_specific_headings():
    opinion = {
        "combined_assessment": {
            "heading": "",
            "paragraphs": [],
            "groups": [
                {
                    "group": "Y1",
                    "labels": ["D1", "D2"],
                    "heading": "D1 ve D2 Dokümanları Birlikte Değerlendirildiğinde",
                    "paragraphs": ["D1 ve D2 birlikte değerlendirildiğinde teknik fark ve teknik etki korunmaktadır."],
                },
                {
                    "group": "Y2",
                    "labels": ["D1", "D3"],
                    "heading": "D1 ve D3 Dokümanları Birlikte Değerlendirildiğinde",
                    "paragraphs": ["D1 ve D3 birlikte değerlendirildiğinde teknik fark ve teknik etki korunmaktadır."],
                },
            ],
        }
    }
    validate_y_combination_group_coverage(opinion, _tr_flat_report())

    bad = {
        "combined_assessment": {
            "heading": "D1, D2 ve D3 Dokümanları Birlikte Değerlendirildiğinde",
            "paragraphs": ["D1, D2 ve D3 birlikte tek kombinasyon olarak değerlendirildiğinde teknik fark korunmaktadır."],
            "groups": [],
        }
    }
    with pytest.raises(ValueError, match="ayrı görünür|ayrı başlık"):
        validate_y_combination_group_coverage(bad, _tr_flat_report())


def test_app_wires_validated_cached_ai_fallback_only_after_parser_failure():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert "if xy_scope and not required_docs:" in src
    assert "xy_scope_ai_fallback_prompt(report_text_scope)" in src
    assert "validate_xy_scope_ai_payload(report_text_scope, fallback_payload)" in src
    assert '"validated_scope", {"required_docs": required_docs}' in src
    assert "required_documents=source_state.get(\"required_docs\") or []" in src

    prompt = xy_scope_ai_fallback_prompt(_tr_flat_report()).casefold()
    assert "hukuki veya teknik yorum yapma" in prompt
    assert "emin değilsen documents=[]" in prompt
