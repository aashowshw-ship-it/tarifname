from pathlib import Path

from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES

ROOT = Path(__file__).resolve().parent


def test_v5470_versions_and_binding_rules():
    assert APP_VERSION == "v5.4.73"
    assert RULESET_VERSION == "2026-09-21.v65"
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
