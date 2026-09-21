from pathlib import Path
import pytest
import app_core as app
from source_guards import validate_invention_domain


def test_extra_instruction_is_bounded_and_repo_subordinate():
    p = app._with_extra_instruction("ANA PROMPT", "özellikle sensör yapısını ayrıntılandır")
    assert "özellikle sensör yapısını ayrıntılandır" in p
    assert "Repo kuralları" in p
    with pytest.raises(ValueError, match="en fazla 500"):
        app._with_extra_instruction("ANA", "x" * 501)


def test_extra_instruction_empty_does_not_change_prompt():
    assert app._with_extra_instruction("ANA", "") == "ANA"


def test_invention_domain_three_way_validation_restored():
    payload = {"invention_domain": "kimya/biyoloji"}
    assert validate_invention_domain(payload) == "Kimya / Biyoloji"


def test_three_workflows_expose_bounded_extra_instruction_fields():
    ui = Path("app.py").read_text(encoding="utf-8")
    assert 'key="tar_extra_instruction"' in ui
    assert 'key="gor_extra_instruction"' in ui
    assert 'key="res_extra_instruction"' in ui
    assert ui.count('max_chars=MAX_EXTRA_INSTRUCTION_CHARS') >= 3


def test_tarifname_reference_label_and_domain_notice_present():
    ui = Path("app.py").read_text(encoding="utf-8")
    assert 'st.text_input("Referans Numarası"' in ui
    assert 'Buluş alanı belirlendi' in ui
