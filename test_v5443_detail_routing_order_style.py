import pytest
from rules import APP_VERSION, RULESET_VERSION, TARIFNAME_RULES
from source_guards import _detail_fact_required
import app_core


def base_draft():
    return {
        "title":"Ayarlanabilir Spektral Dağılıma Sahip LED Tabanlı Solar Simülatör Sistemi",
        "elements":[
            {"number":"1","name":"Solar spektrum kafası"},
            {"number":"2","name":"Çok bantlı LED dizisi"},
        ],
        "detailed_paragraphs":[
            "Solar spektrum kafası (1), farklı dalga boylarında çalışan çoklu LED gruplarını bir araya getirerek ışık üretimini sağlayan ana optik başlıktır. Çok bantlı LED dizisi (2), UV, görünür ve yakın kızılötesi bölgelerde çalışan LED’lerden oluşur.",
            "Buluş, fotovoltaik cihazların kontrollü ışık koşulları altında test edilmesinde kullanılabilmektedir.",
            "Sistem, AM1.5G benzeri geniş spektrumlu bir ışık profiline yaklaştırılabilmektedir.",
        ],
        "source_coverage_map":[],
    }


def test_version_and_rules():
    assert APP_VERSION == "v5.4.46"
    assert RULESET_VERSION == "2026-09-01.v34"
    assert "DETAYLI AÇIKLAMA SIRA KURALI" in TARIFNAME_RULES
    assert "Buluş;" in TARIFNAME_RULES
    assert "uygundur" in TARIFNAME_RULES


def test_problem_fact_not_required_in_detail():
    assert _detail_fact_required({"category":"problem"}) is False
    assert _detail_fact_required({"category":"önceki_teknik"}) is False
    assert _detail_fact_required({"category":"çözüm"}) is True


def test_first_detail_paragraph_must_contain_elements_in_order():
    d=base_draft()
    app_core._validate_detailed_section_order_routing_and_style(d, None, "Türkçe")
    d["detailed_paragraphs"][0]="Çok bantlı LED dizisi (2) açıklanır. Solar spektrum kafası (1) açıklanır."
    with pytest.raises(ValueError, match="sırasını"):
        app_core._validate_detailed_section_order_routing_and_style(d, None, "Türkçe")


def test_prior_art_problem_paragraph_rejected_from_detail():
    d=base_draft(); d["detailed_paragraphs"].append("Bu uygulamalar sonucunda ortaya çıkan temel teknik problemlerden biri homojenliğin sağlanamamasıdır.")
    with pytest.raises(ValueError, match="önceki-teknik/problem"):
        app_core._validate_detailed_section_order_routing_and_style(d, None, "Türkçe")


def test_semicolon_uygundur_and_solution_subject_rejected():
    for bad, msg in [
        ("Buluş; kontrollü ışık üretir.", "noktalı virgül"),
        ("Buluş, deneysel çalışmalar için uygundur.", "uygundur"),
        ("Bu çözüm, LED gruplarını kontrol eder.", "Sunulan çözüm/Bu çözüm/Çözüm"),
    ]:
        d=base_draft(); d["detailed_paragraphs"].append(bad)
        with pytest.raises(ValueError, match=msg):
            app_core._validate_detailed_section_order_routing_and_style(d, None, "Türkçe")


def test_problem_fact_cannot_point_to_detail_section():
    d=base_draft(); d["source_coverage_map"]=[{"fact_id":"P1","covered":True,"sections":["BULUŞUN DETAYLI AÇIKLAMASI"],"evidence":"uzun teknik problem evidence cümlesi burada yer almaktadır"}]
    extracted={"technical_facts":[{"id":"P1","category":"problem","statement":"mevcut sistemlerde sorun"}]}
    with pytest.raises(ValueError, match="evidence olarak bağlanamaz"):
        app_core._validate_detailed_section_order_routing_and_style(d, extracted, "Türkçe")
