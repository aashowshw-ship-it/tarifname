from pathlib import Path

from rules import APP_VERSION, RULESET_VERSION, TARIFNAME_RULES, GORUS_RULES
from validators import _method_claim_how_findings

ROOT = Path(__file__).resolve().parent


def test_rev5_ruleset_and_method_how_rule_present():
    assert APP_VERSION == "v5.4.44"
    assert RULESET_VERSION == "2026-09-08.v46"
    assert "34E-1." in TARIFNAME_RULES
    assert "her zorunlu işlem adımı" in TARIFNAME_RULES.lower()
    assert "kimden/nereden" in TARIFNAME_RULES
    assert "52A. CÜMLE İÇİ BÜYÜK HARF KAPISI" in TARIFNAME_RULES
    assert "60A. TARİFNAME OLUŞTURMA HIZ/DEVAM KURALI" in TARIFNAME_RULES
    assert "GÖRÜŞ HIZ/DEVAM KURALI" in GORUS_RULES


def test_vague_method_step_fails_how_finding():
    draft = {
        "method_steps": [
            {"number": "1001", "text": "Kullanıcı ve mod seçiminin alınması"},
        ]
    }
    findings = _method_claim_how_findings(draft)
    assert findings
    assert any("nereden" in x["message"] or "nasıl" in x["message"] for x in findings)


def test_source_bound_method_step_passes_how_finding():
    draft = {
        "method_steps": [
            {
                "number": "1001",
                "text": "Kullanıcı tarafından arayüz üzerinden sağlanan kaynak dijital içeriğin ve hedef format seçiminin elektronik işlem birimi üzerinde çalışan yazılıma dönüşüm girdisi olarak alınması",
            },
        ]
    }
    assert _method_claim_how_findings(draft) == []


def test_app_has_fail_closed_method_how_and_sentence_case_flags():
    app = (ROOT / "app.py").read_text(encoding="utf-8")
    assert '"method_how_steps_passed":true' in app
    assert '"sentence_case_clean":true' in app
    assert '"method_how_steps_passed", "sentence_case_clean"' in app
    assert "_validate_method_claim_how_test(draft, language)" in app
    assert '"how_test"]' in app


def test_tarifname_checkpoint_stages_and_content_signature_present():
    app = (ROOT / "app.py").read_text(encoding="utf-8")
    assert "hashlib.sha256(bytes(data or b\"\"))" in app
    assert "APP_VERSION, RULESET_VERSION, MODEL" in app
    for stage in ["source_package", "extracted", "literature", "final_draft_audit", "docx"]:
        assert f'"{stage}"' in app
    assert '_workflow_checkpoint_get("tarifname_create"' in app
    assert '_workflow_checkpoint_set("tarifname_create"' in app


def test_gorus_checkpoint_stages_present_and_quality_not_bypassed():
    app = (ROOT / "app.py").read_text(encoding="utf-8")
    assert '_workflow_signature(\n            "gorus_analysis"' in app
    assert '"analysis_source"' in app
    assert '"audited_opinion"' in app
    assert '"final_bundle"' in app
    # audited_opinion is only checkpointed after the AI quality audit path.
    audit_pos = app.index("validate_ai_quality_audit(quality_audit)")
    checkpoint_pos = app.index('_workflow_checkpoint_set("gorus_opinion", opinion_signature, "audited_opinion"')
    assert checkpoint_pos > audit_pos
