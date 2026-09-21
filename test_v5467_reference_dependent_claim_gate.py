from __future__ import annotations

import pytest

from rules import APP_VERSION, RULESET_VERSION, TARIFNAME_RULES
from app_core import _validate_no_reference_marks_before_detail, _validate_dependent_system_claim_product_language


def test_version_and_written_rules():
    assert APP_VERSION == "v5.4.73"
    assert RULESET_VERSION == "2026-09-21.v65"
    low = TARIFNAME_RULES.casefold()
    assert "referans numaralari bölümünden önce" in low
    assert "sağlaması" in low and "sağlayan" in low
    assert "çalıştırmaması" in low and "çalıştırmayan" in low


def test_pre_detail_known_reference_is_fail_closed():
    draft = {
        "elements": [{"number": "90", "name": "Genel tahmin modeli"}, {"number": "91", "name": "Kişiye özel tahmin modeli"}],
        "method_steps": [{"number": "1001", "text": "Oturumun başlatılması"}],
        "technical_field": "Buluş, bir sistem ile ilgilidir.\n\nBuluş, özellikle tahmin ile ilgilidir.",
        "prior_art_general_paragraphs": [],
        "literature_paragraphs": [],
        "short_description_intro": "Buluş, kişiye özel tahmin modelini (91) genel tahmin modelinden (90) ayırmaktadır.",
        "objectives": [],
        "unumbered_invention_definition": "numarasız tanım olup, özelliği;",
        "unumbered_invention_features": ["numarasız özellik"],
        "figure_descriptions": [],
    }
    with pytest.raises(ValueError, match="REFERANS NUMARALARI bölümünden önce"):
        _validate_no_reference_marks_before_detail(draft, "Türkçe")


def test_pre_detail_unnumbered_equivalent_passes():
    draft = {
        "elements": [{"number": "90", "name": "Genel tahmin modeli"}, {"number": "91", "name": "Kişiye özel tahmin modeli"}],
        "method_steps": [{"number": "1001", "text": "Oturumun başlatılması"}],
        "technical_field": "Buluş, bir sistem ile ilgilidir.\n\nBuluş, özellikle tahmin ile ilgilidir.",
        "prior_art_general_paragraphs": [],
        "literature_paragraphs": [],
        "short_description_intro": "Buluş, kişiye özel tahmin modelini genel tahmin modelinden ayırmaktadır.",
        "objectives": [],
        "unumbered_invention_definition": "numarasız tanım olup, özelliği;",
        "unumbered_invention_features": ["numarasız özellik"],
        "figure_descriptions": [],
    }
    _validate_no_reference_marks_before_detail(draft, "Türkçe")


def test_dependent_system_action_noun_fails_even_with_allowed_closure():
    bad = [
        "İstem 1’e uygun sistem olup, özelliği; görüntü alma biriminin (11) görsel işaretleri görüntü tahmin modeline (32) girdi olarak sağlaması ve kaynak kullanılabilirlik biriminin (41) ilgili tahmin modelini çalıştırmaması olmasıdır."
    ]
    with pytest.raises(ValueError, match="eylem isimleştirmesi"):
        _validate_dependent_system_claim_product_language(bad, "Türkçe")


def test_dependent_system_active_element_language_passes():
    good = [
        "İstem 1’e uygun sistem olup, özelliği; görsel işaretleri görüntü tahmin modeline (32) girdi olarak sağlayan görüntü alma birimi (11), fizyolojik sinyalleri fizyolojik tahmin modeline (33) girdi olarak sağlayan biyometrik ve fizyolojik sensör birimi (12) ve söz konusu kaynaklardan biri kapatıldığında ilgili tahmin modelini çalıştırmayan kaynak kullanılabilirlik birimi (41) içermesidir."
    ]
    _validate_dependent_system_claim_product_language(good, "Türkçe")


def test_dependent_system_bare_suitability_fails():
    bad = ["İstem 1’e uygun sistem olup, özelliği; sistemin 6G ortamında çalışmaya uygun bir sistem olmasıdır."]
    with pytest.raises(ValueError, match="uygun"):
        _validate_dependent_system_claim_product_language(bad, "Türkçe")
