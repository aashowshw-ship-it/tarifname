from pathlib import Path
from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES

ROOT = Path(__file__).resolve().parent

def test_v546_versions_and_rules():
    assert APP_VERSION == "v5.4.49"
    assert RULESET_VERSION == "2026-09-04.v39"
    low = GORUS_RULES.casefold()
    assert "başvuru sahibi yazılır" in low
    assert "kategori işaretleri yalnız iç savunma" in low
    assert "şekil kullanımı zorunlu değildir" in low
    assert "tüm patent sayfası verilmez" in low

def test_category_suffix_not_rendered_and_figures_are_selected_only():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert 'suffix = f" ({cat})" if cat else ""' not in src
    assert 'requested_figure_docs = []' in src
    assert 'bool(sec.get("use_figure", False))' in src
    assert '"use_figure":false' in src
