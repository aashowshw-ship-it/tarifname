from copy import deepcopy

import pytest

from app_core import (
    ARASTIRMA_TEMPLATE,
    RESEARCH_WARNING_INTRO,
    _fill_warning_cell,
    safe_output_name,
    validate_research_quality_audit,
    validate_research_report_language,
    validate_research_selection,
)
from rules import APP_VERSION, RULESET_VERSION


def _base_report():
    return {
        "reference": "182046",
        "title": "Wastewater polishing decision support system",
        "report_date": "07.09.2026",
        "purpose": "Belirlenen konuda araştırmanın gerçekleştirilmesi",
        "scope": "Global (İlan edilmiş olan patent başvuruları)",
        "keywords": ["wastewater polishing", "AHP decision support", "global warming potential", "multi criteria process selection", "wastewater treatment selection", "carbon footprint wastewater", "COD polishing", "membrane process selection", "advanced oxidation selection", "adsorption process selection"],
        "ipc_cpc": [{"code": "C02F1/00", "description": "Treatment of water, waste water, or sewage"}],
        "evaluation_intro": "Araştırma sonucunda teknik yakınlığı en yüksek dokümanlar belirlenmiştir.",
        "documents": [
            {
                "label": "D1",
                "number": "CN1",
                "description": [
                    "D1 dokümanı su arıtma alternatiflerinin çok kriterli değerlendirilmesine ilişkindir. AHP tabanlı ağırlıklandırma ve sıralama açıklanmaktadır."
                ],
                "abstract": "The invention relates to a wastewater treatment method; the method includes a first step and a second step.",
                "comparison_rows": [
                    {"feature": "AHP ile kriter ağırlıklandırması", "status_evidence": "+ Özet, İstem 1, ağırlıklandırma açıklanmaktadır."}
                ],
                "novelty_assessment": [
                    "D1 teknik olarak yakın bir karar yapısı açıklamaktadır. Ancak GWP entegrasyonu ve bağlamsal proses seçimi birlikte verilmemektedir. D1 dokümanında GWP entegrasyonu ve bağlamsal proses seçimi ile ilgili bir emareye rastlanmamıştır. Bu kapsamda araştırma konusu buluşun D1 dokümanı varlığında yeni olduğu düşünülmektedir."
                ],
            }
        ],
        "inventive_step_paragraphs": [
            "D1, su arıtma alternatiflerinin çok kriterli değerlendirilmesinde AHP tabanlı ağırlıklandırmayı başlangıç noktası olarak sunmaktadır. Araştırma konusu D1'den farklı olarak GWP hesabını proses seçim skoruyla bütünleştirmekte ve bağlamsal aday proses havuzu oluşturmaktadır. Bu ayrım, teknik performans ile çevresel etkinin aynı seçim mekanizmasında değerlendirilmesi problemine yönelmektedir.",
            "D2, çevresel performans veya karbon etkisinin su arıtma kararlarında dikkate alınmasına ilişkin tamamlayıcı bir öğretim sunduğu varsayılan dokümandır. İlgili alandaki uzman kişinin D1'deki AHP değerlendirmesini geliştirirken D2'deki çevresel metriği kullanmaya yönelmesi için somut teknik motivasyon bulunup bulunmadığı ayrıca incelenmelidir. Böyle bir motivasyon varsa GWP değerinin normalize edilip kompozit skora alınması öngörülebilir bir uygulama olarak değerlendirilebilir.",
            "Kalan bağlamsal filtreleme ve veri sunum özellikleri bilinen veri işleme tercihlerinin ötesinde ortak bir teknik etki oluşturup oluşturmadığı yönünden değerlendirilmelidir. Unsurlar yalnız bağımsız işlevlerini sürdürüyorsa birleşimden beklenmeyen veya sinerjik bir sonuç çıktığı söylenemez. Bu nedenle mevcut varsayımsal D1 ve D2 öğretisi altında araştırma konusunun buluş basamağı kriterini sağlamadığı düşünülmektedir."
        ],
        "conclusion_paragraphs": ["Araştırma konusunun yenilik kriterini sağladığı, buluş basamağı kriterini sağlamadığı düşünülmektedir."],
        "warnings": [
            RESEARCH_WARNING_INTRO,
            "Tarifname yazımı için sistem modülleri arasındaki veri akışının ve çalışma prensibinin netleştirilmesini rica ederiz.",
        ],
    }


def test_version_and_filename_are_v5453_and_underscore_safe():
    assert APP_VERSION == "v5.4.59"
    assert RULESET_VERSION == "2026-09-11.v52"
    assert safe_output_name("Ön%20Araştırma%20Raporu_182046_rev.docx", "x.docx") == "Ön_Araştırma_Raporu_182046_rev.docx"
    assert safe_output_name("Ön Araştırma Raporu_182046_rev.docx", "x.docx") == "Ön_Araştırma_Raporu_182046_rev.docx"


def test_original_abstract_semicolon_is_allowed_but_model_semicolon_is_blocked():
    report = _base_report()
    validate_research_report_language(report)

    bad = deepcopy(report)
    bad["documents"][0]["novelty_assessment"] = [
        "D1 teknik olarak yakındır; ancak ayırt edici özellikler farklıdır. D1 dokümanında GWP entegrasyonu ile ilgili bir emareye rastlanmamıştır. Bu kapsamda araştırma konusu buluşun D1 dokümanı varlığında yeni olduğu düşünülmektedir."
    ]
    with pytest.raises(ValueError, match="noktalı virgül"):
        validate_research_report_language(bad)


def test_post_table_claim_repetition_and_wrong_novelty_tail_are_blocked():
    report = _base_report()
    bad = deepcopy(report)
    bad["documents"][0]["novelty_assessment"] = [
        "İstem 1 AHP yapısını açıklamaktadır. D1 dokümanında GWP entegrasyonu ile ilgili bir emareye rastlanmamıştır. Bu kapsamda araştırma konusu buluşun D1 dokümanı varlığında yeni olduğu düşünülmektedir."
    ]
    with pytest.raises(ValueError, match="istem/şekil"):
        validate_research_report_language(bad)

    bad2 = deepcopy(report)
    bad2["documents"][0]["novelty_assessment"] = ["D1 bazı özellikleri açıklamamaktadır. Bu nedenle yeni olduğu düşünülmektedir."]
    with pytest.raises(ValueError, match="bitiş kalıbı"):
        validate_research_report_language(bad2)


def test_warnings_are_limited_to_two_dynamic_paragraphs():
    report = _base_report()
    bad = deepcopy(report)
    bad["warnings"] = [RESEARCH_WARNING_INTRO, "Bir.", "İki.", "Üç."]
    with pytest.raises(ValueError, match="en fazla iki"):
        validate_research_report_language(bad)


def test_selection_table_semicolon_is_blocked():
    selection = {
        "d1": {"number": "CN1", "abstract_en": "An English abstract."},
        "d2": None,
        "comparison_rows_d1": [{"feature": "feature", "status_evidence": "+ Özet; İstem 1"}],
        "comparison_rows_d2": [],
    }
    with pytest.raises(ValueError, match="noktalı virgül"):
        validate_research_selection(selection)


def test_second_reader_is_fail_closed():
    good = {
        "semicolon_free": {"pass": True, "note": ""},
        "d1d2_concise_template": {"pass": True, "note": ""},
        "novelty_tail_template": {"pass": True, "note": ""},
        "inventive_step_three_substantive_paragraphs": {"pass": True, "note": ""},
        "warnings_source_specific_minimal": {"pass": True, "note": ""},
        "no_redundant_warning_requests": {"pass": True, "note": ""},
        "result_consistency": {"pass": True, "note": ""},
        "overall_pass": True,
    }
    validate_research_quality_audit(good)
    bad = deepcopy(good)
    bad["warnings_source_specific_minimal"] = {"pass": False, "note": "standard uyarı"}
    bad["overall_pass"] = False
    with pytest.raises(ValueError, match="Word kullanıcıya sunulmadı"):
        validate_research_quality_audit(bad)


def test_empty_warning_slot_has_no_visible_bullet():
    from docx import Document

    doc = Document(str(ARASTIRMA_TEMPLATE))
    cell = doc.tables[4].rows[0].cells[2]
    _fill_warning_cell(cell, [RESEARCH_WARNING_INTRO, "Birinci dinamik uyarı.", "İkinci dinamik uyarı."])
    assert len(cell.paragraphs) == 4
    last = cell.paragraphs[3]
    assert last.text == ""
    assert last._p.pPr is None or last._p.pPr.numPr is None
