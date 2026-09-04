from pathlib import Path
import fitz

from gorus_audit import has_usable_non_chinese_figure, extract_cited_original_figure_pages
from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES

ROOT = Path(__file__).resolve().parent


def _pdf_with_text(lines, fontname="helv"):
    doc = fitz.open()
    page = doc.new_page()
    y = 72
    for text in lines:
        page.insert_text((72, y), text, fontsize=12, fontname=fontname)
        y += 24
    return doc.tobytes()


def test_v551_versions_and_mandatory_figure_rule():
    assert APP_VERSION == "v5.4.51"
    assert RULESET_VERSION == "2026-09-04.v41"
    low = GORUS_RULES.casefold()
    assert "kullanılabilir teknik şekil mevcutsa şekil kullanımı zorunludur" in low
    assert "çince/han" in low
    assert "teknik içeriği tamamen korunur" in low


def test_non_chinese_figure_is_detected_and_extracted():
    asset = {"name": "US20210000001A1.pdf", "data": _pdf_with_text(["FIG. 1", "sensor 10", "controller 20"])}
    assert has_usable_non_chinese_figure("US20210000001A1", [asset]) is True
    out = extract_cited_original_figure_pages([
        {"label": "D1", "number": "US20210000001A1", "figure_reference": "Figure 1"}
    ], [asset])
    assert "D1" in out and out["D1"].startswith(b"\x89PNG")


def test_chinese_written_figure_is_not_usable():
    asset = {"name": "CN123456789A.pdf", "data": _pdf_with_text(["图 1", "传感器", "控制器"], fontname="china-s")}
    assert has_usable_non_chinese_figure("CN123456789A", [asset]) is False
    out = extract_cited_original_figure_pages([
        {"label": "D1", "number": "CN123456789A", "figure_reference": "图 1"}
    ], [asset])
    assert out == {}


def test_prompt_and_word_gate_require_usable_figure():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert "Çince/Han karakterli açıklama veya etiket İÇERMEYEN" in src
    assert "Görüş şekil kapısı" in src
    assert "chinese_text_only" in src
