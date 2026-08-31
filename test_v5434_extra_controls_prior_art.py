from pathlib import Path

from rules import (
    APP_VERSION, RULESET_VERSION, TARIFNAME_RULES, EXTRA_CONTROLS_NOTICE,
    tarifname_extra_controls_completed,
)

ROOT = Path(__file__).resolve().parent

def test_version_and_notice_rule():
    assert APP_VERSION == "v5.4.39"
    assert RULESET_VERSION == "2026-08-31.v29"
    assert EXTRA_CONTROLS_NOTICE == "EKSTRA KONTROLLER YAPILDI"
    assert "en az üç gelişmiş paragraf" in TARIFNAME_RULES
    assert "en az 2400 karakter" in TARIFNAME_RULES
    assert "uyarısı yalnız" in TARIFNAME_RULES

def test_extra_notice_is_strict_boolean_gate():
    gates = {
        "source_completeness": True, "prior_art": True, "draft_quality": True,
        "claims": True, "references": True, "template": True,
        "element_step_language": True, "formula_format": True, "how_test": True,
    }
    assert not tarifname_extra_controls_completed(gates, render_passed=False)
    assert tarifname_extra_controls_completed(gates, render_passed=True)
    assert not tarifname_extra_controls_completed(gates, render_passed=True, figures_required=True, figures_passed=False)
    assert tarifname_extra_controls_completed(gates, render_passed=True, figures_required=True, figures_passed=True)
    gates["prior_art"] = False
    assert not tarifname_extra_controls_completed(gates, render_passed=True)

def test_app_surfaces_notice_only_through_gate():
    app=(ROOT/"app.py").read_text(encoding="utf-8")
    assert "figures_required=bool(separate_figures)" in app
    assert "figures_passed=figures_gate_passed" in app
    assert "if extra_controls_done:" in app
    assert "st.warning(EXTRA_CONTROLS_NOTICE)" in app

def test_post_generation_exposes_prior_art_and_draft_status():
    for name in ("app.py", "app_core.py"):
        txt=(ROOT/name).read_text(encoding="utf-8")
        assert '"prior_art": True' in txt
        assert '"draft_quality": True' in txt
        assert "_validate_prior_art_source_placement(draft, extracted, language)" in txt
        assert "_validate_prior_art_bridge_and_depth(draft, extracted, language)" in txt
