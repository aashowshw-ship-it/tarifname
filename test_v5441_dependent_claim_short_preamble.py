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
    assert APP_VERSION == "v5.4.78"
    assert RULESET_VERSION == "2026-09-25.v70"
    assert "İstem X’e uygun sistem olup, özelliği;" in TARIFNAME_RULES
    assert "İstem X’e uygun yöntem olup, özelliği;" in TARIFNAME_RULES
    assert "28A." in TARIFNAME_RULES
    assert "Günümüzde" in TARIFNAME_RULES


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


def test_v5478_turkish_claim_dative_is_number_aware():
    assert app_core._expected_claim_dative_suffix(1) == "e"
    assert app_core._expected_claim_dative_suffix(2) == "ye"
    assert app_core._expected_claim_dative_suffix(6) == "ya"
    assert app_core._expected_claim_dative_suffix(9) == "a"
    assert app_core._expected_claim_dative_suffix(10) == "a"
    assert app_core._expected_claim_dative_suffix(11) == "e"
    assert app_core._normalize_dependent_claim_dative("İstem 9’e uygun yöntem olup, özelliği; x işlem adımını içermesidir.").startswith("İstem 9’a uygun yöntem")


def test_v5478_method_step_repetition_is_rejected_without_percentage_threshold():
    method={"preamble":"Bir elektronik cihaz üzerinde yürütülen yöntem", "steps":["Verinin sunucudan alınması (1001)","Verinin sınıflandırılması (1002)"]}
    deps=["İstem 1’e uygun yöntem olup, özelliği; Verinin sunucudan alınması (1001) işlem adımını içermesidir."]
    with pytest.raises(ValueError, match="yeni bir teknik sınırlama"):
        app_core._validate_dependent_method_claim_semantic_repetition(method,deps,1)


def test_v5478_paraphrased_system_repetition_is_rejected_feature_by_feature():
    sc={"preamble":"Ajan keşif sistemi", "elements":["gelen ajan keşif talebinde belirtilen yetenek, protokol, kimliklendirme yöntemi, iletişim hızı, gecikme süresi, algılama işlevi ve hesaplama kapasitesi gereksinimlerini değerlendirerek eşleştirme yapan ajan keşif birimi (502)"]}
    deps=["İstem 1’e uygun sistem olup, özelliği; ajan keşif biriminin (502), ajan keşif talebinde yetenek, protokol, kimliklendirme yöntemi, iletişim hızı, gecikme süresi, algılama işlevi ve hesaplama kapasitesi gereksinimlerini alan bir birim olmasıdır."]
    with pytest.raises(ValueError, match="yeni bir teknik sınırlama"):
        app_core._validate_dependent_claim_semantic_repetition(sc,deps)


def test_v5478_multiple_new_references_require_real_bullets():
    sc={"preamble":"Bir sistem", "elements":["ana birim (1)"]}
    deps=["İstem 1’e uygun sistem olup, özelliği; birinci yardımcı birim (2) ve ikinci yardımcı birim (3) içermesidir."]
    with pytest.raises(ValueError, match="madde işaretleri"):
        app_core._validate_dependent_claim_semantic_repetition(sc,deps)
    good="İstem 1’e uygun sistem olup, özelliği;\n• birinci yardımcı birim (2),\n• ikinci yardımcı birim (3)\niçermesidir."
    app_core._validate_dependent_claim_semantic_repetition(sc,[good])
    parts=app_core._split_dependent_claim_for_render(good)
    assert parts and len(parts[1])==2


def test_v5478_prior_art_must_start_with_general_gunumuzde_paragraph():
    bad={"prior_art_general_paragraphs":["3GPP dokümanında belirli bir çözüm açıklanmaktadır.","Yukarıda belirtilen eksiklikler, teknik probleme yol açmaktadır."]}
    with pytest.raises(ValueError, match="Günümüzde"):
        app_core._validate_prior_art_general_intro(bad,"Türkçe")
    good={"prior_art_general_paragraphs":["Günümüzde mobil haberleşme ağlarında farklı uç, çekirdek ve bulut kaynakları arasında veri ve hizmet erişiminin dinamik biçimde yönetilmesi yaygınlaşmakta, ağ koşulları ile hesaplama kaynakları uygulama başarımını doğrudan etkileyebilmektedir. Bu nedenle ağ üzerinde çalışan yazılımsal bileşenlerin kullanılabilir kaynakları ve bağlantı koşullarını dikkate alması teknik açıdan önem taşımaktadır.","Yukarıda belirtilen eksiklikler, teknik probleme yol açmaktadır."]}
    app_core._validate_prior_art_general_intro(good,"Türkçe")


def test_v5478_house_style_normalizes_ajanlar_arasi_and_wrong_dative():
    d=minimal_draft()
    d["title"]="Ajanlar-Arası Test Sistemi"
    d["dependent_system_claims"]=["İstem 2’e uygun sistem olup, özelliği; yardımcı birim (2) içermesidir."]
    styled=app_core.apply_tarifname_house_style(d,"Yalnızca sistem",[],"Türkçe")
    assert "Ajanlar Arası" in styled["title"]
    assert styled["dependent_system_claims"][0].startswith("İstem 2’ye uygun")
