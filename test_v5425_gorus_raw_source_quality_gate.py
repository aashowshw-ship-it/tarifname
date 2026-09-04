from __future__ import annotations
from pathlib import Path
import pytest
from gorus_audit import detect_examiner_reasoned_documents, validate_opinion_against_raw_sources
from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES
ROOT = Path(__file__).resolve().parent
REPORT = """
İlgili Dokümanlar
D1: JP2005137881A
D2: CN108832840A
D3: US2013020842A1
Patentlenebilirlik Şartları
Yenilik
Tekniğin bilinen durumuna en yakın doküman olarak seçilen D1 nolu dokümanda açıklamalar yapılmıştır.
Buluş basamağı
D1 dokümanının [0006], [0025], [0026], [0009] ve [0010] nolu paragrafları incelendiğinde değerlendirme yapılmıştır.
İstem 1 ve ona bağlı 2 nolu istem buluş basamağı içermemektedir.
"""
SPEC = "İstem 1 kapsamında varlık sensörü (1), ısıtıcı (2), sıcaklık ölçüm birimi (3), ince film (4) bulunmaktadır."
def base_opinion():
    return {
        "application_no":"2024/007330","applicant":"ASSAN","reference":"173541",
        "intro":"20.05.2026 tarihli 1. İnceleme Bildiriminde 1 ve 2 numaralı istemlerin yenilik ve sanayiye uygulanabilirlik kriterlerini sağladığı, ancak buluş basamağı kriterini sağlamadığı değerlendirilmiştir. Başvuru sahibinin görüşleri aşağıda dikkatinize sunulmaktadır.",
        "cited_documents":[{"label":"D1","number":"JP2005137881A","summary":"Araç koltuğu ısıtmasına ilişkin bir düzen açıklanmaktadır."}],
        "sections":[{"label":"D1","heading":"D1 (JP2005137881A) dokümanı:","blocks":[
            {"type":"paragraph","text":"İnceleme bildiriminde D1’in [0006], [0025], [0026], [0009] ve [0010] paragrafları esas alınmıştır. İstem 1 bakımından ayırt edici teknik fark ve teknik katkı, aynı piezoelektrik varlık sensörünün algılama ve enerji üretimini birlikte gerçekleştirmesidir. Teknik etki, bu işlevlerin ortak esnek yapıda bütünleşmesidir. Buna göre objektif teknik problem bu iki işlevin ortak yapıda nasıl gerçekleştirileceğidir. D1’de bu dönüşüme yönelik motivasyon veya yönlendirme bulunmamaktadır. D1 bu ilave teknik değişikliğe yönelik bir öğretim veya yönlendirme sağlamamaktadır."},
            {"type":"quote","text":SPEC,"attach_to_previous":True}],
            "inventive_step_heading":"D1 karşısında buluş basamağı",
            "inventive_step_paragraphs":["İstem 2, istem 1’in teknik katkısını devralmaktadır. Teknik etki ve objektif teknik problem bakımından aynı çekirdek devam eder. Uzman kişinin çözüme ulaşması için ilave yapısal değişiklikler gerekir ve D1’de yönlendirme yoktur."]}],
        "combined_assessment":{"heading":"D1 dokümanının inceleme gerekçesiyle birlikte değerlendirilmesi","paragraphs":[("Teknik fark ve teknik etki birlikte değerlendirildiğinde objektif teknik problem açık hale gelmektedir. D1’in somut öğretisinde başvurudaki teknik katkıya yönelik motivasyon veya yönlendirme yoktur. Uzman kişinin birden fazla yapısal ve işlevsel değişiklik yapması gerekir. Bu ilave değişiklikler D1’in somut teknik öğretisinden doğrudan çıkmaz ve D1 bu yönde teknik bir yönlendirme sağlamaz. "*5)]},
        "conclusion":["İstem 1 ve İstem 2 bakımından buluş basamağı kriterinin sağlandığı değerlendirilmektedir."]}
def test_version_bumped():
    assert APP_VERSION == "v5.4.45" and RULESET_VERSION == "2026-09-04.v35"
def test_rules_capture_new_controls():
    low=GORUS_RULES.casefold()
    for phrase in ["noktalı virgül","aynı paragrafta","ham-kaynak","okuma kapisi","teknik katkı","savunmada gerekli dokümanlar","referans numaraları"]: assert phrase in low
def test_reasoned_docs_only_d1():
    docs=detect_examiner_reasoned_documents(REPORT)
    assert [d["label"] for d in docs]==["D1"] and docs[0]["number"]=="JP2005137881A"
def test_rejects_irrelevant_d2():
    op=base_opinion(); op["cited_documents"].append({"label":"D2","number":"CN108832840A","summary":"x"})
    with pytest.raises(ValueError,match="kullanılmayan|fiilen"): validate_opinion_against_raw_sources(op,REPORT,SPEC)
def test_rejects_intro_procedure_and_semicolon():
    op=base_opinion(); op["intro"] += " D1 üzerinden savunma yapılacaktır."
    with pytest.raises(ValueError,match="giriş"): validate_opinion_against_raw_sources(op,REPORT,SPEC)
    op=base_opinion(); op["sections"][0]["blocks"][0]["text"] += " Bu teknik fark önemlidir; etki sağlar."
    with pytest.raises(ValueError,match="noktalı virgül"): validate_opinion_against_raw_sources(op,REPORT,SPEC)
def test_quote_attach_and_prior_art_ref_number():
    op=base_opinion(); op["sections"][0]["blocks"][1]["attach_to_previous"]=False
    with pytest.raises(ValueError,match="aynı paragraf"): validate_opinion_against_raw_sources(op,REPORT,SPEC)
    op=base_opinion(); op["sections"][0]["blocks"][0]["text"] += " Oturma tespit anahtarı 150 kullanılır."
    with pytest.raises(ValueError,match="referans"): validate_opinion_against_raw_sources(op,REPORT,SPEC)
def test_app_wires_staged_scope_second_read_and_quality_report():
    src=(ROOT/"app.py").read_text(encoding="utf-8")
    for phrase in ["Raporu analiz et ve savunmada gerekli dokümanları belirle","detect_examiner_reasoned_documents","gorus_quality_audit_prompt","gorus_repair_prompt","validate_opinion_against_raw_sources","validate_gorus_docx_content_flow","Çıktı kalite kontrolü","_append_quote_with_location"]: assert phrase in src
