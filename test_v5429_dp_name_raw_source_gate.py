from pathlib import Path

import pytest

from source_guards import (
    derive_tarifname_output_names,
    validate_final_raw_source_audit,
    validate_final_source_coverage_chain,
)


def test_dp_reference_derives_both_output_names_without_extra_filename_input():
    assert derive_tarifname_output_names("181267") == (
        "Tarifname_181267.docx",
        "Şekiller_181267.docx",
    )


def test_dp_reference_is_filename_sanitized_but_semantically_preserved():
    assert derive_tarifname_output_names(" 181267 / A ") == (
        "Tarifname_181267___A.docx",
        "Şekiller_181267___A.docx",
    )


def test_empty_dp_reference_is_rejected():
    with pytest.raises(ValueError, match="DP referans numarası"):
        derive_tarifname_output_names("  ")


def _sample_source_state():
    registry = [
        {"passage_id": "B0001", "source": "BBF", "text": "Teknik modül belleği dışarıdan okur."},
        {"passage_id": "B0002", "source": "BBF", "text": "İmza ve iletişim bilgileri."},
    ]
    extracted = {
        "technical_facts": [
            {"id": "T001", "statement": "Teknik modül belleği dışarıdan okur.", "mandatory": True},
        ],
        "source_passage_audit": [
            {"passage_id": "B0001", "classification": "technical", "fact_ids": ["T001"], "reason": ""},
            {"passage_id": "B0002", "classification": "nontechnical", "fact_ids": [], "reason": "idari form alanı"},
        ],
    }
    evidence = "Teknik modül, sanal makine belleğini hipervizör üzerinden dışarıdan okur."
    coverage = [{"fact_id": "T001", "covered": True, "sections": ["BULUŞUN DETAYLI AÇIKLAMASI"], "evidence": evidence}]
    final_text = "Başlık\n" + evidence + "\nSonuç"
    raw_audit = {
        "passage_checks": [{"passage_id": "B0001", "covered": True, "evidence": [evidence], "missing_detail": ""}],
        "fact_checks": [{"fact_id": "T001", "covered": True, "evidence": [evidence], "missing_detail": ""}],
        "all_pass": True,
    }
    return registry, extracted, coverage, final_text, raw_audit, evidence


def test_final_raw_source_chain_passes_only_with_real_final_evidence():
    registry, extracted, coverage, final_text, raw_audit, evidence = _sample_source_state()
    stats = validate_final_source_coverage_chain(extracted, registry, coverage, final_text)
    assert stats == {
        "raw_passages_total": 2,
        "technical_passages": 1,
        "nontechnical_passages": 1,
        "technical_facts": 1,
        "covered_facts": 1,
    }
    audit_stats = validate_final_raw_source_audit(raw_audit, extracted, registry, final_text)
    assert audit_stats["audited_technical_passages"] == 1
    assert audit_stats["audited_technical_facts"] == 1


def test_final_source_chain_rejects_fact_without_word_evidence():
    registry, extracted, coverage, final_text, raw_audit, evidence = _sample_source_state()
    coverage[0]["evidence"] = "Bu ifade nihai Word dosyasında yer almamaktadır."
    with pytest.raises(ValueError, match="NİHAİ HAM VERİ ZİNCİRİ"):
        validate_final_source_coverage_chain(extracted, registry, coverage, final_text)


def test_final_raw_second_read_rejects_missing_technical_passage():
    registry, extracted, coverage, final_text, raw_audit, evidence = _sample_source_state()
    raw_audit["passage_checks"] = []
    with pytest.raises(ValueError, match="SON HAM KAYNAK İKİNCİ OKUMA"):
        validate_final_raw_source_audit(raw_audit, extracted, registry, final_text)


def test_tarifname_ui_no_longer_asks_for_output_filenames():
    app_text = (Path(__file__).parent / "app.py").read_text(encoding="utf-8")
    tarifname_section = app_text.split('if work_type == "Tarifname oluşturma":', 1)[1].split('# GÖRÜŞ', 1)[0]
    assert 'st.text_input("Çıktı dosyasının adı"' not in tarifname_section
    assert 'st.text_input("Şekiller dosyasının adı"' not in tarifname_section
    assert 'derive_tarifname_output_names(reference)' in tarifname_section
    assert 'Ham veri kontrolü yapıldı:' in tarifname_section
