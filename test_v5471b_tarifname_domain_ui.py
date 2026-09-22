from pathlib import Path

import pytest

from rules import APP_VERSION, RULESET_VERSION, TARIFNAME_RULES
from source_guards import INVENTION_DOMAINS, validate_invention_domain

ROOT = Path(__file__).resolve().parent


def test_release_is_v5472_and_domain_rule_is_binding():
    assert APP_VERSION == "v5.4.75"
    assert RULESET_VERSION == "2026-09-22.v67"
    assert "6V. BULUŞ ALANI OTOMATİK SINIFLANDIRMA KURALI" in TARIFNAME_RULES
    assert INVENTION_DOMAINS == (
        "Elektrik-Elektronik / Yazılım",
        "Kimya / Biyoloji",
        "Mekanik",
    )


def test_domain_validator_canonicalizes_and_fails_closed():
    payload = {"invention_domain": "kimya/biyoloji"}
    assert validate_invention_domain(payload) == "Kimya / Biyoloji"
    assert payload["invention_domain"] == "Kimya / Biyoloji"
    with pytest.raises(ValueError):
        validate_invention_domain({"invention_domain": ""})
    with pytest.raises(ValueError):
        validate_invention_domain({"invention_domain": "Belirsiz"})


def test_tarifname_ui_uses_reference_number_and_auto_domain_notice():
    app = (ROOT / "app.py").read_text(encoding="utf-8")
    tariff_start = app.index('if work_type == "Tarifname oluşturma":')
    tariff_end = app.index('# TARİFNAME DÜZENLEME', tariff_start)
    tariff_ui = app[tariff_start:tariff_end]
    assert 'st.text_input("Referans Numarası"' in tariff_ui
    assert 'st.text_input("DP referans numarası"' not in tariff_ui
    assert 'st.info(f"**Buluş alanı belirlendi:** {invention_domain}")' in tariff_ui
    assert 'validate_invention_domain(extracted)' in tariff_ui


def test_extraction_prompt_requires_exact_three_way_domain_in_both_cores():
    for filename in ("app.py", "app_core.py"):
        src = (ROOT / filename).read_text(encoding="utf-8")
        assert '"invention_domain":"Elektrik-Elektronik / Yazılım | Kimya / Biyoloji | Mekanik"' in src
        assert "BULUŞ ALANI SINIFLANDIRMASI" in src
        assert "Yardımcı nitelikte elektronik bulunması tek başına alanı Elektrik-Elektronik / Yazılım yapmaz." in src
