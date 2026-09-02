import pytest
import app_core
from rules import APP_VERSION, RULESET_VERSION, TARIFNAME_RULES
from source_guards import validate_detailed_description_source_transfer


def test_version_and_lossless_patent_layer_rules():
    assert APP_VERSION == "v5.4.54"
    assert RULESET_VERSION == "2026-09-01.v34"
    assert "KAYIPSIZ PATENT YENİDEN YAZIM KURALI" in TARIFNAME_RULES
    assert "TEKNİK LİTERAL KAYIP KAPISI" in TARIFNAME_RULES


def _draft(with_embodiment=True):
    p2=("Buluşun bir yapılanmasında, " if with_embodiment else "Buluşta, ") + "solar spektrum kafası (1) içindeki çok bantlı LED dizisi (2), kontrol ünitesi (7) üzerinden birlikte sürülmektedir."
    return {
        "elements":[
            {"number":"1","name":"Solar spektrum kafası"},
            {"number":"2","name":"Çok bantlı LED dizisi"},
            {"number":"7","name":"Kontrol ünitesi"},
        ],
        "detailed_paragraphs":[
            "Solar spektrum kafası (1), ışık üretir. Çok bantlı LED dizisi (2), farklı bantları içerir. Kontrol ünitesi (7), LED gruplarını sürer.",
            p2,
            "Sistem, AM1.5G benzeri geniş spektrumlu profile ve 365–1000 nm çalışma aralığına göre ayarlanabilmektedir. Yakın kızılötesi LED grubu 850 nm çevresinde baskın sürülebilmektedir. PWM tabanlı kontrol kullanılmaktadır.",
        ],
        "working_principle":"Solar spektrum kafası (1) içerisinde bulunan çok bantlı LED dizisi (2), kontrol ünitesi (7) üzerinden sürülerek seçilen spektral dağılımın oluşturulmasını sağlamaktadır.",
        "alternatives":[],
    }


def _extracted():
    return {"technical_facts":[
        {"id":"T1","category":"unsur","statement":"Solar spektrum kafası ışık üretir"},
        {"id":"T2","category":"örnek","statement":"Örneğin sistem 850 nm çevresinde seçilebilir modda çalışabilir"},
    ]}


def test_patent_layer_requires_embodiment_wording_for_source_example():
    app_core._validate_detailed_patent_drafting_layer(_draft(True), _extracted(), "Türkçe")
    with pytest.raises(ValueError, match="Buluşun bir yapılanmasında"):
        app_core._validate_detailed_patent_drafting_layer(_draft(False), _extracted(), "Türkçe")


def test_patent_layer_requires_real_working_principle():
    d=_draft(True); d["working_principle"]="Kısa."
    with pytest.raises(ValueError, match="working_principle"):
        app_core._validate_detailed_patent_drafting_layer(d, _extracted(), "Türkçe")


def test_raw_literal_gate_catches_am15g_even_when_fact_statement_omits_it():
    registry=[{"passage_id":"B1","source":"BBF","text":"Sistem AM1.5G benzeri profilde 365–1000 nm aralığında ve 850 nm seçici modunda PWM ile çalışır."}]
    extracted={
        "technical_facts":[{"id":"T1","category":"çözüm","statement":"Sistem ayarlanabilir spektral modda çalışır"}],
        "source_passage_audit":[{"passage_id":"B1","classification":"technical","fact_ids":["T1"],"reason":"buluşun teknik çalışma koşullarını açıklar"}],
    }
    coverage=[{"fact_id":"T1","covered":True,"sections":["BULUŞUN DETAYLI AÇIKLAMASI"],"evidence":"Sistem ayarlanabilir spektral modda çalışır ve 365–1000 nm aralığında çalışır."}]
    detail="Sistem ayarlanabilir spektral modda çalışır ve 365–1000 nm aralığında çalışır. 850 nm seçici modunda PWM ile çalışır."
    with pytest.raises(ValueError, match="AM1.5G"):
        validate_detailed_description_source_transfer(extracted,registry,coverage,detail,[])


def test_raw_literal_gate_passes_when_am15g_is_present():
    registry=[{"passage_id":"B1","source":"BBF","text":"Sistem AM1.5G benzeri profilde 365–1000 nm aralığında ve 850 nm seçici modunda PWM ile çalışır."}]
    extracted={
        "technical_facts":[{"id":"T1","category":"çözüm","statement":"Sistem ayarlanabilir spektral modda çalışır"}],
        "source_passage_audit":[{"passage_id":"B1","classification":"technical","fact_ids":["T1"],"reason":"buluşun teknik çalışma koşullarını açıklar"}],
    }
    detail="Sistem ayarlanabilir spektral modda çalışır ve AM1.5G benzeri profilde 365–1000 nm aralığında çalışır. 850 nm seçici modunda PWM ile çalışır."
    coverage=[{"fact_id":"T1","covered":True,"sections":["BULUŞUN DETAYLI AÇIKLAMASI"],"evidence":"Sistem ayarlanabilir spektral modda çalışır ve AM1.5G benzeri profilde 365–1000 nm aralığında çalışır."}]
    stats=validate_detailed_description_source_transfer(extracted,registry,coverage,detail,[])
    assert stats["detail_protected_literals"] >= 4
