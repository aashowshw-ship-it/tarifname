from pathlib import Path

from rules import APP_VERSION, RULESET_VERSION, TARIFNAME_RULES
from tarifname_figure_generation import build_method_flow_svg

ROOT = Path(__file__).resolve().parent


def test_version_bumped():
    assert APP_VERSION == "v5.4.87"
    assert RULESET_VERSION == "2026-10-09.v79"


def test_method_flow_refs_are_unique_and_ordered():
    steps = [{"number": str(n), "text": f"step {n}"} for n in range(1001, 1007)]
    svg = build_method_flow_svg(steps).decode("utf-8")
    for n in range(1001, 1007):
        assert svg.count(f">{n}</text>") == 1
    assert "Şekil" not in svg and "Yöntem Akışı" not in svg
    positions = [svg.index(f">{n}</text>") for n in range(1001, 1007)]
    assert positions == sorted(positions)


def test_rules_require_single_canonical_method_text_and_unique_method_figure():
    assert "TEK KANONİK YÖNTEM-ADIMI KAYNAĞI" in TARIFNAME_RULES
    assert "AYRI YÖNTEM AKIŞ ŞEKLİ" in TARIFNAME_RULES
    assert "tam bir kez" in TARIFNAME_RULES.casefold()


def test_app_has_final_word_triple_sync_and_method_figure_duplicate_gate():
    app = (ROOT / "app.py").read_text(encoding="utf-8")
    core = (ROOT / "app_core.py").read_text(encoding="utf-8")
    for text in (app, core):
        assert "Word yöntem senkron kontrolü" in text
        assert "method_reference_sequence" in text
        assert "method_reference_counts" in text
        assert "method_refs_unique_and_ordered" in text
        assert "replace_with_generated_method_flow" in text
