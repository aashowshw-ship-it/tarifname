import pytest

from rules import APP_VERSION, RULESET_VERSION, TARIFNAME_RULES, EXTRA_CONTROL_GATE_KEYS
from source_guards import validate_detailed_description_fact_coverage, validate_detailed_description_source_transfer


def _state():
    registry = [
        {"passage_id":"B0001","source":"BBF","text":"Sistem, yalnızca AM1.5G benzeri geniş spektrumlu bir ışık profiline yaklaştırılabilmekte ve 365–1000 nm aralığında çalışmaktadır."},
        {"passage_id":"B0002","source":"BBF","text":"Birinci eleman (1), farklı dalga boylarındaki LED gruplarını bir araya getirir."},
    ]
    extracted = {
        "technical_facts":[
            {"id":"T001","category":"çözüm","statement":"AM1.5G benzeri profil ve 365–1000 nm çalışma aralığı", "mandatory":True},
            {"id":"T002","category":"unsur","statement":"Solar spektrum kafası farklı LED gruplarını bir araya getirir", "mandatory":True},
        ],
        "source_passage_audit":[
            {"passage_id":"B0001","classification":"technical","fact_ids":["T001"],"reason":"spektral çalışma koşulu ve teknik değer içerir"},
            {"passage_id":"B0002","classification":"technical","fact_ids":["T002"],"reason":"birinci teknik unsurun işlevini açıklar"},
        ],
    }
    coverage = [
        {"fact_id":"T001","covered":True,"sections":["BULUŞUN DETAYLI AÇIKLAMASI"],"evidence":"AM1.5G benzeri geniş spektrumlu bir ışık profiline yaklaştırılabilmekte ve 365–1000 nm aralığında çalışmaktadır."},
        {"fact_id":"T002","covered":True,"sections":["BULUŞUN DETAYLI AÇIKLAMASI"],"evidence":"Solar spektrum kafası (1), farklı dalga boylarındaki LED gruplarını bir araya getirir."},
    ]
    detail = "AM1.5G benzeri geniş spektrumlu bir ışık profiline yaklaştırılabilmekte ve 365–1000 nm aralığında çalışmaktadır. Solar spektrum kafası (1), farklı dalga boylarındaki LED gruplarını bir araya getirir."
    elements = [{"number":"1","name":"Solar spektrum kafası"}]
    return registry, extracted, coverage, detail, elements


def test_version_and_binding_rule():
    assert APP_VERSION == "v5.4.52"
    assert RULESET_VERSION == "2026-09-01.v34"
    assert "BULUŞUN DETAYLI AÇIKLAMASI bölümünde de eksiksiz" in TARIFNAME_RULES
    assert "AM1.5G" in TARIFNAME_RULES
    assert "detail_source_transfer" in EXTRA_CONTROL_GATE_KEYS


def test_detail_fact_gate_passes_with_detail_evidence_and_literals():
    registry, extracted, coverage, detail, elements = _state()
    stats = validate_detailed_description_fact_coverage(extracted, coverage, detail, elements)
    assert stats["detail_required_facts"] == 2
    assert stats["detail_covered_facts"] == 2
    full = validate_detailed_description_source_transfer(extracted, registry, coverage, detail, elements)
    assert full["detail_protected_literals"] >= 2


def test_detail_gate_rejects_missing_am15g_even_if_other_sections_cover_fact():
    registry, extracted, coverage, detail, elements = _state()
    broken_detail = detail.replace("AM1.5G benzeri ", "")
    coverage[0]["evidence"] = "365–1000 nm aralığında çalışmaktadır. Solar spektrum kafası"
    with pytest.raises(ValueError, match="literal"):
        validate_detailed_description_source_transfer(extracted, registry, coverage, broken_detail, elements)


def test_detail_gate_rejects_generic_element_placeholder_when_real_name_known():
    registry, extracted, coverage, detail, elements = _state()
    broken = detail + " Birinci eleman (1), ışık üretir."
    with pytest.raises(ValueError, match="geçici/genel"):
        validate_detailed_description_fact_coverage(extracted, coverage, broken, elements)


def test_independent_audit_rejects_classification_mismatch_and_false_detail_flag():
    from source_guards import validate_final_raw_source_audit, _registry_fingerprint, _draft_fingerprint
    registry=[{"passage_id":"B0001","source":"BBF","text":"Sistem AM1.5G profilinde 850 nm ışınım üretir."}]
    extracted={
        "technical_facts":[{"id":"F1","category":"solution","statement":"Sistem AM1.5G profilinde 850 nm ışınım üretir."}],
        "source_passage_audit":[{"passage_id":"B0001","classification":"technical","fact_ids":["F1"],"reason":"teknik çözüm"}],
    }
    final="Sistem AM1.5G profilinde 850 nm ışınım üretir."
    base={"audit_meta":{"audit_mode":"independent_raw_source_second_read_v2","audit_nonce":"n","source_fingerprint":_registry_fingerprint(registry),"draft_fingerprint":_draft_fingerprint(final),"independent_second_read":True,"prior_classification_used":False,"source_coverage_map_used":False},"all_pass":True}
    bad=dict(base); bad["passage_checks"]=[{"passage_id":"B0001","classification":"nontechnical","classification_reason":"yanlış sınıflandırma testi","source_quote":"Sistem AM1.5G profilinde 850 nm ışınım üretir.","covered":True,"evidence":[],"detail_transfer_required":False,"detail_evidence":""}]
    import pytest
    with pytest.raises(ValueError, match="classification_mismatch"):
        validate_final_raw_source_audit(bad,extracted,registry,final,detail_text=final,expected_audit_nonce="n")
    bad2=dict(base); bad2["passage_checks"]=[{"passage_id":"B0001","classification":"technical","classification_reason":"buluşun teknik çözümünü açıklar","source_quote":"Sistem AM1.5G profilinde 850 nm ışınım üretir.","covered":True,"evidence":[final],"detail_transfer_required":False,"detail_evidence":""}]
    with pytest.raises(ValueError, match="detail_transfer_mismatch"):
        validate_final_raw_source_audit(bad2,extracted,registry,final,detail_text=final,expected_audit_nonce="n")

def test_all_explicit_elements_must_be_in_system_claim_set():
    import app_core, pytest
    draft={"elements":[{"number":"1","name":"Ana birim"},{"number":"14","name":"Bağlantı girişleri"}],"system_claim":{"preamble":"Teknik sistem olup, özelliği;","elements":["işlem yapan ana birim (1),"],"closing":"içermesidir."},"dependent_system_claims":[]}
    with pytest.raises(ValueError, match=r"Bağlantı girişleri \(14\)"):
        app_core._validate_all_elements_covered_in_claims(draft)
    draft["dependent_system_claims"]=["İstem 1’e uygun sistem olup, özelliği; veri iletimi sağlayan bağlantı girişleri (14) içermesidir."]
    app_core._validate_all_elements_covered_in_claims(draft)
