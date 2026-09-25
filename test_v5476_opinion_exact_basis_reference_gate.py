from pathlib import Path
import fitz
import pytest

from gorus_audit import (
    build_page_line_index,
    clear_page_line_index_cache,
    locate_quote_page_line_span,
    validate_line_reference_authority,
    validate_mandatory_spec_basis_coverage,
    validate_opinion_reference_binding,
)
from rules import APP_VERSION, RULESET_VERSION, FINAL_COMPLIANCE_REQUIRED_CHECKS

ROOT = Path(__file__).resolve().parent


def _exact_line_pdf() -> bytes:
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    # Physical line pitch = 17 pt. Printed anchors share the same baseline as
    # the corresponding body line, as in the filing template.
    for n, y in [(5, 100), (10, 185), (15, 270)]:
        page.insert_text((34, y), str(n), fontsize=10)
    rows = {
        8: "First quoted sentence starts on physical line eight and continues with technical text",
        9: "second physical line carries the middle portion of the exact specification quotation",
        10: "third physical line continues the quotation without any paragraph approximation",
        11: "last quoted characters end on physical line eleven.",
    }
    for line_no, text in rows.items():
        y = 100 + (line_no - 5) * 17
        page.insert_text((116, y), text, fontsize=9)
    data = doc.tobytes()
    doc.close()
    return data


def test_v5476_version_and_required_gorus_receipts():
    assert APP_VERSION == "v5.4.78"
    assert RULESET_VERSION == "2026-09-25.v70"
    assert (ROOT / "README.md").read_text(encoding="utf-8").startswith("# Patent Atölyesi v5.4.78")
    required = FINAL_COMPLIANCE_REQUIRED_CHECKS["gorus"]
    assert "spec_basis_coverage" in required
    assert "reference_binding" in required
    assert "exact_physical_lines" in required


def test_reference_metadata_must_use_main_file_reference_not_opinion_filename_reference():
    validate_opinion_reference_binding({"reference": "178930"}, "178930")
    with pytest.raises(ValueError, match="ana dosya referansı"):
        validate_opinion_reference_binding({"reference": "699149"}, "178930")


def test_substantive_defence_cannot_ship_without_spec_quote():
    x_without_quote = {
        "cited_documents": [{"label": "D1", "category": "X"}],
        "sections": [{"label": "D1", "blocks": [{"type": "paragraph", "text": "Somut teknik savunma paragrafı."}]}],
        "combined_assessment": {"groups": []},
    }
    with pytest.raises(ValueError, match="birebir tarifname dayanağı zorunludur"):
        validate_mandatory_spec_basis_coverage(x_without_quote)

    x_with_quote = {
        "cited_documents": [{"label": "D1", "category": "X"}],
        "sections": [{"label": "D1", "blocks": [
            {"type": "paragraph", "text": "Somut teknik savunma paragrafı."},
            {"type": "quote", "text": "Tarifnamedeki birebir teknik cümle."},
        ]}],
        "combined_assessment": {"groups": []},
    }
    validate_mandatory_spec_basis_coverage(x_with_quote)


def test_real_y_combination_group_requires_its_own_quote():
    base = {
        "cited_documents": [
            {"label": "D1", "category": "Y1"},
            {"label": "D2", "category": "Y1"},
        ],
        "sections": [
            {"label": "D1", "blocks": [{"type": "paragraph", "text": "D1 objektif tanıtım."}]},
            {"label": "D2", "blocks": [{"type": "paragraph", "text": "D2 objektif tanıtım."}]},
        ],
        "combined_assessment": {"groups": [{
            "heading": "D1 ve D2 Birlikte Değerlendirildiğinde",
            "blocks": [{"type": "paragraph", "text": "Birlikte esas savunma."}],
        }]},
    }
    with pytest.raises(ValueError, match="birlikte değerlendirme"):
        validate_mandatory_spec_basis_coverage(base)

    base["combined_assessment"]["groups"][0]["blocks"].append(
        {"type": "quote", "text": "Tarifnamedeki birebir birleşik teknik dayanak."}
    )
    validate_mandatory_spec_basis_coverage(base)


def test_word_source_uses_internal_pdf_authority():
    with pytest.raises(ValueError, match="doğrulama kaynağı PDF"):
        validate_line_reference_authority("Tarifname.docx", "Tarifname.docx")
    validate_line_reference_authority("Tarifname.docx", "Tarifname__internal_page_line.pdf")
    validate_line_reference_authority("Tarifname.pdf", "Tarifname.pdf")


def test_quote_span_is_first_character_line_to_last_character_line_exactly():
    clear_page_line_index_cache()
    pdf = _exact_line_pdf()
    idx = build_page_line_index("Tarifname_WordExport.pdf", pdf)
    quote = (
        "First quoted sentence starts on physical line eight and continues with technical text "
        "second physical line carries the middle portion of the exact specification quotation "
        "third physical line continues the quotation without any paragraph approximation "
        "last quoted characters end on physical line eleven."
    )
    assert locate_quote_page_line_span("Tarifname_WordExport.pdf", pdf, quote, idx) == (1, 8, 1, 11)


def test_ui_requires_word_export_pdf_and_separates_two_references():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert "Tarifname sayfa/satır doğrulama PDF'si (zorunlu)" not in src
    assert "prepare_line_reference_source" in src
    assert "arka planda PDF'ye çevrilir" in src
    assert "Görüş referansı (çıktı dosyası için)" in src
    assert 'opinion["reference"] = source_state.get("reference") or ""' in src
    assert "validate_mandatory_spec_basis_coverage" in src
    assert "validate_line_reference_authority" in src
