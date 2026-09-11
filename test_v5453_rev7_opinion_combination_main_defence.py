from pathlib import Path

import pytest
from docx import Document

from app_core import build_gorus_docx
from gorus_audit import validate_opinion_narrative_rules, validate_opinion_payload
from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES

ROOT = Path(__file__).resolve().parent


def _tr_y_report():
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
“Y” Buluşun buluş basamağı içermediğini başka bir dokümanla bir araya getirildiğinde gösteren doküman
"""


def _group_text(a: str, b: str) -> list[str]:
    p1 = (
        f"Araştırma raporundaki uzman itirazı {a} ve {b} dokümanlarını tamamlayıcı teknik öğretiler olarak birlikte ele almaktadır. "
        f"{a} teknik başlangıç noktası kabul edildiğinde istemde kalan ayırt edici teknik fark, başvurudaki unsurların belirli bir işlevsel ilişki içinde birlikte çalışmasıdır. "
        "Bu işlevsel ilişki teknik farkın nasıl ortaya çıktığını da açıklar, çünkü bir unsurdan elde edilen fiziksel veya işlevsel sonuç sonraki unsurun çalışma koşulunu doğrudan belirlemektedir. "
        "Bu teknik farkın teknik etkisi, tek tek bilinen parçaların yan yana bulunmasından farklı olarak sistem davranışının kararlı ve kontrollü biçimde elde edilmesidir. "
        "Buna göre objektif teknik problem, mevcut başlangıç yapısını bu özel işlevsel ilişkiyi sağlayacak biçimde nasıl düzenlemek gerektiğidir. "
    )
    p2 = (
        f"{b} dokümanının öğretisi incelendiğinde bu objektif teknik problemi çözmek amacıyla {a} yapısının söz konusu işlevsel ilişkiye dönüştürülmesine yönelik somut bir motivasyon veya yönlendirme bulunmamaktadır. "
        "Dokümanların birlikte okunabilmesi, istemdeki çözüme doğrudan ulaşılacağı anlamına gelmez. "
        "İsteme ulaşmak için mevcut elemanların görevlerinin değiştirilmesi, aralarındaki etkileşim sırasının yeniden kurulması ve teknik çıktının sonraki unsurla bağlanması gibi ilave yapısal ve işlevsel değişiklikler gerekir. "
        "Bu ilave değişiklikler iki dokümanın açıklamalarında doğal veya kaçınılmaz bir uyarlama olarak öğretilmemektedir. "
        "Dolayısıyla uzman kişinin yalnız bu iki kaynağın gerçek teknik öğretisinden hareketle başvurudaki özel unsur-işlev ilişkisine ulaşması için gerekli teknik köprü mevcut değildir. "
        "Bu nedenle birlikte değerlendirme de istemdeki teknik katkıyı aşikâr hale getirmemektedir. "
    )
    # Depth gate intentionally demands a genuinely substantive main defence.
    return [p1 + p1, p2 + p2]


def _opinion():
    return {
        "application_no": "2025/008936",
        "applicant": "ASELSAN ELEKTRONİK SANAYİ VE TİCARET ANONİM ŞİRKETİ",
        "reference": "699870",
        "intro": "09.06.2026 tarihli araştırma raporunda, 1-6 numaralı istemlerin D1, D2 ve D3 dokümanları varlığında buluş basamağı kriterini sağlamadığı belirtilmiştir. Başvuru sahibinin görüşleri aşağıda dikkatinize sunulmaktadır. Araştırma raporunda 1-6 numaralı istemler bakımından gösterilen benzer dokümanlar aşağıdadır:",
        "cited_documents": [
            {"label": "D1", "number": "CN103644554B", "category": "Y", "category_marker": "Y1,Y2", "combination_groups": ["Y1", "Y2"], "summary": "D1 teknik öğretisi."},
            {"label": "D2", "number": "CN117460385A", "category": "Y", "category_marker": "Y1", "combination_groups": ["Y1"], "summary": "D2 teknik öğretisi."},
            {"label": "D3", "number": "CN201360012Y", "category": "Y", "category_marker": "Y2", "combination_groups": ["Y2"], "summary": "D3 teknik öğretisi."},
        ],
        "sections": [
            {"label": "D1", "blocks": [{"type": "paragraph", "text": "D1 dokümanının gerçek teknik öğretisi kısa ve objektif olarak açıklanmıştır."}], "novelty_heading": "", "novelty_paragraphs": [], "inventive_step_heading": "", "inventive_step_paragraphs": []},
            {"label": "D2", "blocks": [{"type": "paragraph", "text": "D2 dokümanının gerçek teknik öğretisi kısa ve objektif olarak açıklanmıştır."}], "novelty_heading": "", "novelty_paragraphs": [], "inventive_step_heading": "", "inventive_step_paragraphs": []},
            {"label": "D3", "blocks": [{"type": "paragraph", "text": "D3 dokümanının gerçek teknik öğretisi kısa ve objektif olarak açıklanmıştır."}], "novelty_heading": "", "novelty_paragraphs": [], "inventive_step_heading": "", "inventive_step_paragraphs": []},
        ],
        "combined_assessment": {
            "heading": "",
            "paragraphs": [],
            "groups": [
                {"group": "Y1", "labels": ["D1", "D2"], "heading": "D1 ve D2 Dokümanları Birlikte Değerlendirildiğinde", "paragraphs": _group_text("D1", "D2")},
                {"group": "Y2", "labels": ["D1", "D3"], "heading": "D1 ve D3 Dokümanları Birlikte Değerlendirildiğinde", "paragraphs": _group_text("D1", "D3")},
            ],
        },
        "conclusion": ["Bu nedenlerle istemlerin buluş basamağı içerdiği değerlendirilmektedir."],
        "signoff": "Saygılarımızla,\nDESTEK PATENT A.Ş.",
    }


def test_rev7_rules_make_combination_the_main_defence():
    assert APP_VERSION == "v5.4.59"
    assert RULESET_VERSION == "2026-09-11.v52"
    low = GORUS_RULES.casefold()
    for phrase in [
        "asıl ve en ikna edici",
        "ayrı görünür başlık",
        "en az iki dolu savunma paragrafı",
        "uzmanın kombinasyon mantığı",
        "fail-closed",
    ]:
        assert phrase in low


def test_numbered_y_groups_pass_only_as_two_deep_separate_sections():
    op = _opinion()
    validate_opinion_payload(op, _tr_y_report(), "teknik tarifname")

    bad = _opinion()
    bad["combined_assessment"]["groups"][0]["paragraphs"] = ["D1 ve D2 birlikte farklıdır."]
    with pytest.raises(ValueError, match="derinlik|iki dolu"):
        validate_opinion_payload(bad, _tr_y_report(), "teknik tarifname")


def test_mega_heading_for_y1_y2_is_fail_closed():
    bad = _opinion()
    bad["combined_assessment"] = {
        "heading": "D1, D2 ve D3 Dokümanları Birlikte Değerlendirildiğinde",
        "paragraphs": _group_text("D1", "D2"),
        "groups": [],
    }
    with pytest.raises(ValueError, match="ayrı görünür|groups"):
        validate_opinion_payload(bad, _tr_y_report(), "teknik tarifname")


def test_docx_renders_each_combination_as_its_own_heading():
    op = _opinion()
    data = build_gorus_docx(op)
    doc = Document(__import__('io').BytesIO(data))
    texts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    assert texts.count("D1 ve D2 Dokümanları Birlikte Değerlendirildiğinde") == 1
    assert texts.count("D1 ve D3 Dokümanları Birlikte Değerlendirildiğinde") == 1
    assert not any("D1, D2 ve D3 Dokümanları Birlikte" in x for x in texts)


def test_generated_opinion_rejects_plus_sign_combination_notation():
    op = _opinion()
    op["intro"] = "Applicant observations are submitted below."
    op["combined_assessment"]["groups"][0]["paragraphs"][0] += " D1+D2 birlikte ele alınmıştır."
    with pytest.raises(ValueError, match="artı işareti"):
        validate_opinion_narrative_rules(op, "Inventive step objection", "teknik tarifname")
