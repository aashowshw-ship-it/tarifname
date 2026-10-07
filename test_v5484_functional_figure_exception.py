from pathlib import Path

from rules import APP_VERSION, RULESET_VERSION, TARIFNAME_RULES

ROOT = Path(__file__).resolve().parent


def test_version_bumped_for_functional_figure_rule():
    assert APP_VERSION == "v5.4.84"
    assert RULESET_VERSION == "2026-10-07.v76"


def test_rules_distinguish_physical_functional_and_method_figures():
    low = TARIFNAME_RULES.casefold()
    assert "65f-1-şeki̇l" in low or "65f-1-şekil" in low
    assert "group_container" in TARIFNAME_RULES
    assert "boş kutu" in low
    assert "her yöntem işlem adımı kendi ayrı kutucuğunda" in low


def test_app_prompts_preserve_functional_text_and_group_container_exception():
    for name in ("app.py", "app_core.py"):
        text = (ROOT / name).read_text(encoding="utf-8")
        assert "functional_architecture" in text
        assert "process_decision_flow" in text
        assert "group_container" in text
        assert "reference_layout_valid" in text
        assert "BOŞ KUTU" in text
        assert "teknik akışı anlaşılır kılan Türkçe kutu metinlerini" in text


def test_physical_single_reference_rule_not_removed():
    rules = TARIFNAME_RULES.casefold()
    assert "tek fi̇zi̇ksel unsur/çağri = tek referans" in rules or "tek fiziksel unsur/çağrı = tek referans" in rules
