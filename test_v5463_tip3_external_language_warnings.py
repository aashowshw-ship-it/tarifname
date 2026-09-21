from copy import deepcopy

import pytest

from app_core import (
    RESEARCH_WARNING_INTRO,
    RESEARCH_METHOD_WARNING,
    RESEARCH_SYSTEM_WARNING,
    validate_research_report_language,
)
from rules import APP_VERSION, RULESET_VERSION, ARASTIRMA_RULES


def _report(warning):
    return {
        "reference":"182292",
        "title":"Dağıtık akustik konum ve rota kestirimi",
        "report_date":"14.09.2026",
        "purpose":"Belirlenen konuda araştırmanın gerçekleştirilmesi",
        "scope":"Global (İlan edilmiş olan patent başvuruları)",
        "keywords":["distributed acoustic sensing","TDOA localization","GCC PHAT","drone localization","acoustic tracking","multilateration","track before detect","multi target tracking","propagation delay","trajectory estimation"],
        "ipc_cpc":[{"code":"G01S5/20","description":"Position-fixing by acoustic waves"}],
        "evaluation_intro":"Araştırma sonucunda, araştırma konusu ile teknik yakınlığı en yüksek dokümanlar US1 (D1) ve US2 (D2) olarak değerlendirilmiştir.",
        "documents":[
            {"label":"D1","number":"US1","description":["D1 dağıtık akustik algılama ile hava hedefi takibini açıklamaktadır. Yön ve mesafe verileri üzerinden rota tahmini yapılmaktadır."],"abstract":"An acoustic detection system for aerial vehicles.","comparison_rows":[{"feature":"TDOA tabanlı konum kestirimi","status_evidence":"- Özet, doğrudan TDOA açıklanmamaktadır."}],"novelty_assessment":["D1 teknik olarak yakın bir akustik takip yaklaşımı açıklamaktadır. Ancak GCC-PHAT tabanlı M-TDOA ve gecikme belirsizliğinin kovaryans modeline taşınması birlikte verilmemektedir. D1 dokümanında GCC-PHAT tabanlı M-TDOA ve gecikme uyarlaması ile ilgili bir emareye rastlanmamıştır. Bu kapsamda araştırma konusu buluşun D1 dokümanı varlığında yeni olduğu düşünülmektedir."]},
            {"label":"D2","number":"US2","description":["D2 mikrofon dizisiyle üç boyutlu konum belirlemeyi açıklamaktadır. TDOA tabanlı geometrik konumlandırma kullanılmaktadır."],"abstract":"A microphone array determines a three-dimensional position.","comparison_rows":[{"feature":"TDOA tabanlı konum kestirimi","status_evidence":"+ Şekil 11, TDOA tabanlı konum açıklanmaktadır."}],"novelty_assessment":["D2 TDOA tabanlı konumlandırma bakımından yakındır. Ancak süreç gürültüsü seğirmesi ve HMM tabanlı tespit-öncesi-takip zinciri açıklanmamaktadır. D2 dokümanında süreç gürültüsü seğirmesi ve HMM tabanlı tespit-öncesi-takip ile ilgili bir emareye rastlanmamıştır. Bu kapsamda araştırma konusu buluşun D2 dokümanı varlığında yeni olduğu düşünülmektedir."]},
        ],
        "inventive_step_paragraphs":[
            "D1, dağıtık akustik hedef takibini başlangıç noktası olarak sunmaktadır. Araştırma konusu D1'den farklı olarak tek mikrofonlu düğümlerden elde edilen zaman farklarını GCC-PHAT ve M-TDOA ile üç boyutlu konuma dönüştürmektedir. Ayrıca yayılım gecikmesinin oluşturduğu belirsizlik süreç kovaryansına taşınmaktadır. Bu ayrımlar dinamik hava hedeflerinde konum ve rota kestiriminin kararlılığına yönelmektedir.",
            "D2, TDOA tabanlı üç boyutlu konumlandırma öğretisini tamamlayıcı nitelikte sunmaktadır. Bununla birlikte D1'den D2'ye yönelmek, süreç gürültüsü seğirmesi ile çoklu hedef ve düşük SNR takibini kendiliğinden sağlamamaktadır. İlgili alandaki uzman kişinin bu unsurları aynı işlem zincirinde bir araya getirmesi için ayrıca teknik seçimler yapması gerekir. Bu nedenle kombinasyonun araştırma konusundaki bütün işlevsel ilişkilere doğrudan ulaştığı söylenemez.",
            "Kalan çoklu hedef takibi ve tespit-öncesi-takip özellikleri yalnız bağımsız algoritma seçimleri olarak değil, zayıf akustik tepelerin zamansal tutarlılıkla korunması ve rota kestirimiyle birlikte değerlendirilmelidir. Bu yapı düşük SNR ve hayalet tepe koşullarında kararlı hedef varlığı üretmeye yöneliktir. D1 ve D2'nin birlikte öğretisi bu teknik etkiyi aynı zincirde açıkça kurmamaktadır. Bu nedenle araştırma konusunun buluş basamağı kriterini sağladığı düşünülmektedir."
        ],
        "conclusion_paragraphs":["Araştırma konusunun yenilik ve buluş basamağı kriterlerini sağladığı düşünülmektedir."],
        "warnings":[RESEARCH_WARNING_INTRO, warning],
    }


def test_version_and_rules():
    assert APP_VERSION == "v5.4.73"
    assert RULESET_VERSION == "2026-09-21.v65"
    assert "yalnız sabit girişle boş bırakılamaz" in ARASTIRMA_RULES
    assert "dış okuyucuya sunulan bağımsız bir uzman raporu" in ARASTIRMA_RULES


def test_method_or_system_baseline_warning_is_mandatory():
    validate_research_report_language(_report(RESEARCH_METHOD_WARNING))
    validate_research_report_language(_report(RESEARCH_SYSTEM_WARNING))
    bad=_report(RESEARCH_METHOD_WARNING)
    bad["warnings"]=[RESEARCH_WARNING_INTRO]
    with pytest.raises(ValueError, match="zorunlu temel uyarı"):
        validate_research_report_language(bad)


def test_arbitrary_warning_is_rejected_as_first_dynamic_warning():
    bad=_report("Buluş daha ayrıntılı açıklanmalıdır.")
    with pytest.raises(ValueError, match="bağlayıcı temel uyarı"):
        validate_research_report_language(bad)


def test_internal_process_wording_is_rejected():
    bad=_report(RESEARCH_METHOD_WARNING)
    bad["inventive_step_paragraphs"][2] = "Raporda doğrulanan D1 ve D2 öğretisi bu unsurları birlikte açıklamamaktadır. Kalan özellikler teknik etki bakımından farklılaşmaktadır. Bu farklılık düşük SNR koşullarındaki kararlılığa katkı sağlamaktadır. Bu nedenle araştırma konusunun buluş basamağı kriterini sağladığı düşünülmektedir."
    with pytest.raises(ValueError, match="dış-okuyucu dili"):
        validate_research_report_language(bad)
