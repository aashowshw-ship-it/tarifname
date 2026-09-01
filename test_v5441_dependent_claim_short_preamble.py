from __future__ import annotations
import copy
import pytest
from rules import APP_VERSION, RULESET_VERSION, TARIFNAME_RULES
from validators import validate_draft
import app_core


def minimal_draft():
    return {
        "title":"Test sistemi",
        "field":"Test alanı",
        "prior_art":["Mevcut sistemlerde teknik bir problem bulunmaktadır."],
        "objectives":["Teknik problemi çözmektir."],
        "figures":[],
        "elements":[{"number":"1","name":"Kontrol birimi"}],
        "method_steps":[],
        "detailed_paragraphs":["Kontrol birimi (1), veriyi işler."],
        "system_claim":{"preamble":"Test ortamında veriyi işleyerek teknik bir çıktı üreten ve bu çıktıyı bağlı birime aktaran elektronik kontrol sistemi","elements":["veriyi işleyerek çıktı üreten kontrol birimi (1),"],"closing":"içermesidir."},
        "dependent_system_claims":["İstem 1’e uygun sistem olup, özelliği; kontrol biriminin (1) bir bellek içermesidir."],
        "method_claim":None,
        "dependent_method_claims":[],
        "abstract":"Test sistemi teknik bir çıktı üretmektedir.",
        "coverage_audit":{},
    }


def test_versions_and_rule_text():
    assert APP_VERSION == "v5.4.43"
    assert RULESET_VERSION == "2026-09-01.v33"
    assert "İstem X’e uygun sistem olup, özelliği;" in TARIFNAME_RULES
    assert "İstem X’e uygun yöntem olup, özelliği;" in TARIFNAME_RULES


def test_generic_validator_rejects_long_dependent_system_title():
    d=minimal_draft()
    d["dependent_system_claims"]=["İstem 1’e uygun ayarlanabilir spektral dağılıma sahip LED tabanlı solar simülatör sistemi olup, özelliği; bağlantı soketleri içermesidir."]
    msgs=[x["message"] for x in validate_draft(d) if x.get("level")=="Hata"]
    assert any("kısa giriş" in m for m in msgs)


def test_generic_validator_accepts_short_dependent_system_and_method_starts():
    d=minimal_draft()
    d["dependent_method_claims"]=["İstem 1'e uygun yöntem olup, özelliği; bir doğrulama işlem adımını içermesidir."]
    msgs=[x["message"] for x in validate_draft(d) if x.get("level")=="Hata"]
    assert not any("kısa giriş kalıbıyla" in m for m in msgs)


def test_word_output_gate_rejects_long_dependent_preamble():
    good=minimal_draft()
    app_core._validate_dependent_claim_short_starts_texts(
        ["Test ortamında uzun bağımsız sistem olup, özelliği;", good["dependent_system_claims"][0]], good, "Türkçe"
    )
    bad=copy.deepcopy(good)
    bad["dependent_system_claims"]=["İstem 1’e uygun ayarlanabilir solar simülatör sistemi olup, özelliği; kontrol biriminin (1) bir bellek içermesidir."]
    with pytest.raises(ValueError, match="bağımlı istem girişinde"):
        app_core._validate_dependent_claim_short_starts_texts(
            ["Test ortamında uzun bağımsız sistem olup, özelliği;", bad["dependent_system_claims"][0]], bad, "Türkçe"
        )
