import copy
import pytest

from gorus_audit import (
    detect_ep_xy_documents,
    validate_opinion_narrative_rules,
    validate_opinion_payload,
    validate_y_combination_group_coverage,
)
from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES


REPORT = """
TÜRK PATENT VE MARKA KURUMU
ARAŞTIRMA RAPORU
C. İLGİLİ DOKÜMANLAR
Kategori Dokümanlar İlgili Olduğu İstem
X
Y
Y
Y
WO2022144052A1 (A)
WO2018195255A1 (B)
WO2015005644A1 (C)
US10748450B1 (D)
1
1
1
1
Kategorilerin Açıklaması:
“X” Buluşun yeni olmadığını veya buluş basamağı içermediğini tek başına gösteren doküman
“Y” Buluşun buluş basamağı içermediğini başka bir dokümanla bir araya getirildiğinde gösteren doküman
"""


def _deep_group():
    p1 = (
        "Araştırma raporundaki Y itirazı D2, D3 ve D4 dokümanlarının tamamlayıcı teknik öğretilerinin birlikte ele alınmasına dayanır. "
        "Teknik başlangıç noktası dikkate alındığında ayırt edici teknik fark, işlem adımlarının belirli bir işlevsel ilişki içinde ardışık çalışmasıdır. "
        "Bu işlevsel ilişki teknik farkın nasıl ortaya çıktığını gösterir ve teknik etki, kullanıcının fiziksel etkileşim ile geri bildirim arasında kontrollü bir eğitim döngüsü kurabilmesidir. "
        "Buna göre objektif teknik problem, bilinen ayrı eğitim araçlarından hareketle bu özel fiziksel ve işlevsel eğitim döngüsünün nasıl kurulacağıdır. "
    )
    p2 = (
        "D2, D3 ve D4 birlikte okunsa dahi bu objektif teknik problemi çözmek için açıklanan yapıları söz konusu özel ilişkiye dönüştürmeye yönelik somut bir motivasyon veya yönlendirme bulunmamaktadır. "
        "İsteme ulaşmak için parçaların görevlerinin değiştirilmesi, fiziksel yerleştirme ilişkilerinin yeniden kurulması ve geri bildirimin senaryo akışına bağlanması gibi ilave yapısal ve işlevsel değişiklikler gerekir. "
        "Bu ilave değişiklikler kaynakların teknik öğretisinde doğal veya kaçınılmaz bir uyarlama olarak sunulmamaktadır. "
        "Dolayısıyla dokümanların birlikte öğretisi özel unsur-işlev ilişkisini doğrudan ve açık biçimde ortaya koymamaktadır. "
    )
    return [p1 + p1, p2 + p2]


def _opinion():
    return {
        "application_no": "2026/003962",
        "applicant": "KOCAELİ ÜNİVERSİTESİ",
        "reference": "699132",
        "intro": "30.06.2026 tarihli araştırma raporunda, 1 numaralı istemin D1, D2, D3 ve D4 dokümanları varlığında yenilik ve buluş basamağı kriterlerini sağlamadığı belirtilmiştir. Başvuru sahibinin görüşleri aşağıda dikkatinize sunulmaktadır. Araştırma raporunda 1 numaralı istem bakımından gösterilen benzer dokümanlar aşağıdadır:",
        "cited_documents": [
            {"label": "D1", "number": "WO2022144052A1", "category": "X", "summary": "D1 teknik öğretisi."},
            {"label": "D2", "number": "WO2018195255A1", "category": "Y", "summary": "D2 teknik öğretisi."},
            {"label": "D3", "number": "WO2015005644A1", "category": "Y", "summary": "D3 teknik öğretisi."},
            {"label": "D4", "number": "US10748450B1", "category": "Y", "summary": "D4 teknik öğretisi."},
        ],
        "sections": [
            {"label": "D1", "blocks": [{"type": "paragraph", "text": "D1 dokümanı objektif olarak açıklanmıştır."}], "novelty_heading": "", "novelty_paragraphs": ["D1 istemdeki tam teknik bütünü doğrudan ve açık biçimde açıklamamaktadır."], "inventive_step_heading": "", "inventive_step_paragraphs": ["D1 tek başına teknik etki ve objektif teknik problem bakımından istemdeki çözüme yönelik motivasyon veya yönlendirme sağlamamaktadır."]},
            {"label": "D2", "blocks": [{"type": "paragraph", "text": "D2 dokümanı taşınabilir dijital diyabet eğitim cihazını ve sanal hasta yönetimini açıklamaktadır."}], "novelty_heading": "", "novelty_paragraphs": [], "inventive_step_heading": "", "inventive_step_paragraphs": []},
            {"label": "D3", "blocks": [{"type": "paragraph", "text": "D3 dokümanı görme engellilere yönelik puzzle elemanları ve Braille işaretlemeyi açıklamaktadır."}], "novelty_heading": "", "novelty_paragraphs": [], "inventive_step_heading": "", "inventive_step_paragraphs": []},
            {"label": "D4", "blocks": [{"type": "paragraph", "text": "D4 dokümanı artırılmış gerçeklikle tıbbi prosedür eğitimi sağlayan peluş oyuncak ve mobil cihaz düzenini açıklamaktadır."}], "novelty_heading": "", "novelty_paragraphs": [], "inventive_step_heading": "", "inventive_step_paragraphs": []},
        ],
        "combined_assessment": {
            "heading": "",
            "paragraphs": [],
            "groups": [{
                "group": "Y",
                "labels": ["D2", "D3", "D4"],
                "heading": "D2, D3 ve D4 Dokümanları Birlikte Değerlendirildiğinde",
                "paragraphs": _deep_group(),
            }],
        },
        "conclusion": ["Bu nedenlerle istem 1 bakımından araştırma raporundaki itirazların yeniden değerlendirilmesi gerektiği kanaatindeyiz."],
    }


def test_v5458_version_and_rules_present():
    assert APP_VERSION == "v5.4.59"
    assert RULESET_VERSION == "2026-09-11.v52"
    low = GORUS_RULES.casefold()
    for phrase in ["bireysel y", "numarasız y", "d2, d3 ve d4", "x dokümanı", "objektif teknik"]:
        assert phrase in low


def test_flat_y_group_excludes_x_and_groups_only_y_documents():
    docs = detect_ep_xy_documents(REPORT)
    assert [(d["label"], d["category"]) for d in docs] == [("D1", "X"), ("D2", "Y"), ("D3", "Y"), ("D4", "Y")]
    op = _opinion()
    validate_y_combination_group_coverage(op, REPORT)
    validate_opinion_payload(op, REPORT, "teknik tarifname")

    bad = copy.deepcopy(op)
    bad["combined_assessment"]["groups"][0]["labels"] = ["D1", "D2", "D3", "D4"]
    bad["combined_assessment"]["groups"][0]["heading"] = "D1, D2, D3 ve D4 Dokümanları Birlikte Değerlendirildiğinde"
    with pytest.raises(ValueError, match="Y-kombinasyon|yalnız"):
        validate_y_combination_group_coverage(bad, REPORT)


def test_individual_y_defence_is_fail_closed_but_x_defence_remains():
    op = _opinion()
    validate_opinion_narrative_rules(op, REPORT, "teknik tarifname")
    bad = copy.deepcopy(op)
    bad["sections"][1]["inventive_step_paragraphs"] = ["D2 karşısında istem buluş basamağı içermektedir."]
    with pytest.raises(ValueError, match="Y dokümanı"):
        validate_opinion_narrative_rules(bad, REPORT, "teknik tarifname")

    bad = copy.deepcopy(op)
    bad["sections"][2]["blocks"][0]["text"] += " Başvuru konusu istem bakımından ayırt edici teknik fark burada bulunmamaktadır."
    with pytest.raises(ValueError, match="Y-bireysel"):
        validate_opinion_narrative_rules(bad, REPORT, "teknik tarifname")
