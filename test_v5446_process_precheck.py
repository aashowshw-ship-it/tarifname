from pathlib import Path

from docx import Document

from processes import (
    application_precheck_missing,
    build_epats_application_package,
    count_claims_from_docx,
    epats_document_metrics,
    extract_application_source_text,
    normalize_application_information,
    split_patent_docx,
    strip_template_colored_text,
)

BASE = Path(__file__).resolve().parent


def test_colored_template_guidance_is_removed():
    data = (BASE / "Tarifname_181176_template.docx").read_bytes()
    cleaned = strip_template_colored_text(data)
    doc = Document(__import__("io").BytesIO(cleaned))
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "Araştırma raporunun düzenlenebilmesi için" not in text
    assert "Eğer istemler başlığı altında" not in text
    assert "TARİFNAME" in text and "İSTEMLER" in text and "ÖZET" in text


def test_claim_count_uses_claim_list_not_nested_elements():
    data = (BASE / "Tarifname_181176_template.docx").read_bytes()
    parts = split_patent_docx(data)
    assert count_claims_from_docx(parts["Istemler.docx"]) == 3


def test_document_metrics_use_pages_and_claim_count():
    data = (BASE / "Tarifname_181176_template.docx").read_bytes()
    _package, pdfs = build_epats_application_package(data)
    metrics = epats_document_metrics(data, pdfs)
    assert metrics["specification_pages"] > 0
    assert metrics["codes"]["specification"].startswith("T-")
    assert metrics["claim_count"] == 3
    assert metrics["codes"]["claims"] == "İ-3"
    assert metrics["codes"]["abstract"] == "Ö"


def test_eml_source_text_reads_headers_and_body():
    eml = b"From: client@example.com\nTo: patent@example.com\nSubject: Patent basvurusu\nContent-Type: text/plain; charset=utf-8\n\nHak sahibi: ABC A.S.\nBulus sahibi: Ali Veli"
    text = extract_application_source_text("mail.eml", eml)
    assert "Patent basvurusu" in text
    assert "ABC A.S." in text
    assert "Ali Veli" in text


def test_missing_gate_blocks_incomplete_application():
    metadata = normalize_application_information({
        "application_kind": "Patent",
        "invention_title": "Deneme",
        "applicants": [{"name": "ABC A.Ş.", "country": "Türkiye", "address": "Bursa"}],
        "inventors": [{"name": "Ali Veli", "country": "Türkiye", "address": "Bursa"}],
        "priority": {"status": "Yok"},
    })
    metrics = {"specification_pages": 7, "claim_count": 12, "abstract_present": True, "figures_pages": 0}
    assert application_precheck_missing(metadata, metrics) == []
    metadata["applicants"][0]["address"] = ""
    assert any("adres" in item for item in application_precheck_missing(metadata, metrics))
