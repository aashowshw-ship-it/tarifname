from pathlib import Path
from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES

ROOT = Path(__file__).resolve().parent

def test_v551_versions_and_rules():
    assert APP_VERSION == "v5.4.78"
    assert RULESET_VERSION == "2026-09-25.v70"
    low = GORUS_RULES.casefold()
    assert "başvuru sahibi yazılır" in low
    assert "kategori işaretleri yalnız iç savunma" in low
    assert "şekil kullanımı zorunlu" in low
    assert "çince/han" in low
    assert "tüm patent sayfası verilmez" in low

def test_category_suffix_not_rendered_and_figures_are_mandatory_when_usable():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert 'suffix = f" ({cat})" if cat else ""' not in src
    assert 'requested_figure_docs = []' in src
    assert 'has_usable_non_chinese_figure' in src
    assert 'usable_non_chinese and not bool(sec.get("use_figure", False))' in src
    assert '"use_figure":true' in src

# ---- merged verbatim test logic from test_v5470_opinion_delivery_ratio_cost.py ----
def test_v5470_versions_and_binding_rules():
    assert APP_VERSION == "v5.4.78"
    assert RULESET_VERSION == "2026-09-25.v70"
    assert "persuasion_probability" in GORUS_RULES
    assert "Tahmini AI kullanım maliyeti" in GORUS_RULES
    assert "Word Response Letter/Görüş Metni içine yazılmaz" in GORUS_RULES


def test_opinion_delivery_ui_requires_ratio_and_cost_summary():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert 'st.metric("Mevcut uzman itirazını geri çektirme olasılığı (tahmini)"' in src
    assert 'st.metric("Tahmini AI kullanım maliyeti", cost_text)' in src
    assert 'Bağımsız uzman ikna oranı bulunamadı; görüş teslim kapısı nedeniyle Word indirilemez.' in src
    assert 'Görüş AI telemetrisi bulunamadı; oran/ücret teslim kapısı nedeniyle Word indirilemez.' in src
    assert '_show_gorus_ai_metrics(str(source_state.get("workflow_signature")' in src


def test_opinion_ai_calls_are_metered():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    required_stages = [
        "Görüş teknik analiz",
        "Görüş taslağı",
        "Görüş ikinci okuma kalite denetimi",
        "Görüş bağımsız uzman ikna değerlendirmesi",
    ]
    for stage in required_stages:
        assert f'metric_stage="{stage}"' in src
    assert 'metric_context=gorus_metric_context' in src
