"""v5.4.87: Word and LibreOffice page/line divergences must NEVER silently pass."""
import io
from pathlib import Path

import fitz
import pytest
from docx import Document

from gorus_audit import (
    validate_word_origin_pdf_authority,
    locate_quote_page_line_span,
    prepare_line_reference_source,
    build_page_line_index,
)
from rules import APP_VERSION, RULESET_VERSION, FINAL_COMPLIANCE_REQUIRED_CHECKS

QUOTE = 'Güneş ışığı, RF sinyalleri ve çevresel titreşim kaynaklarından enerji toplayan sisteme enerji veren bir enerji modülü (3),'


def word_docx():
    word = Document()
    word.add_paragraph(QUOTE)
    out = io.BytesIO()
    word.save(out)
    return out.getvalue()


def printed_pdf(first_line):
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    for num, y in ((5, 120), (10, 205), (15, 290)):
        page.insert_text((30, y), str(num), fontsize=10)
    page.insert_font(fontname="DejaVu", fontfile="/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
    a, b = QUOTE.split(' toplayan ')
    for offset, text in ((0, a), (1, 'toplayan '+b)):
        y = 120 + (first_line + offset - 5) * 17
        page.insert_text((110, y), text, fontsize=9, fontname="DejaVu")
    pdf = doc.tobytes(); doc.close()
    return pdf


def test_versions_and_rules():
    assert APP_VERSION == 'v5.4.87'
    assert RULESET_VERSION == '2026-10-09.v79'
    assert 'word_origin_pdf' in FINAL_COMPLIANCE_REQUIRED_CHECKS['gorus']


def test_auto_libreoffice_pdf_cannot_certify_word_layout_even_if_anchors_are_plausible():
    b = word_docx()
    auto_name, auto_pdf = prepare_line_reference_source('source.docx', b)
    with pytest.raises(ValueError, match='LibreOffice PDF'):
        validate_word_origin_pdf_authority('source.docx',b,auto_name,auto_pdf,uploaded_by_user=False)


def test_external_word_pdf_anchors_control_exact_699151_example():
    b = word_docx()
    pdf = printed_pdf(7)
    validate_word_origin_pdf_authority('source.docx', b, 'Word_export.pdf', pdf, uploaded_by_user=True)
    assert locate_quote_page_line_span('Word_export.pdf', pdf, QUOTE) == (1,7,1,8)
    # The internal/alternate 5-6 grid must never override the authority PDF.
    lo_style_pdf = printed_pdf(5)
    assert locate_quote_page_line_span('lo_style.pdf',lo_style_pdf,QUOTE)==(1,5,1,6)


def test_stale_or_unreadable_word_pdf_is_refused():
    b = word_docx()
    with pytest.raises(ValueError,match='eşleşmiyor'):
        validate_word_origin_pdf_authority('source.docx',b,'other.pdf',printed_other(),uploaded_by_user=True)
    with pytest.raises(ValueError,match='geçersiz'):
        validate_word_origin_pdf_authority('source.docx',b,'bad.pdf',b'%PDF-not-real',uploaded_by_user=True)


def printed_other():
    doc=fitz.open();p=doc.new_page()
    p.insert_text((30,100),'5');p.insert_text((30,185),'10')
    p.insert_text((150,120),'Tamamen farkli bir tarifname metni.',fontsize=9)
    out=doc.tobytes();doc.close();return out


def test_source_ui_and_final_gate_both_require_uploaded_word_pdf():
    app = (Path(__file__).resolve().parent/'app.py').read_text(encoding='utf-8')
    assert 'gor_native_word_pdf' in app
    assert 'gor_revised_native_pdf' in app
    assert 'validate_word_origin_pdf_authority(' in app
    assert 'line_spec_user_word_export' in app
    assert 'revised_line_spec_user_word_export' in app
