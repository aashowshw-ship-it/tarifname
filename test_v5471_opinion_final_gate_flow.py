from __future__ import annotations

import io
from pathlib import Path

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from app_core import build_gorus_docx
from gorus_audit import validate_gorus_template_fidelity, validate_opinion_narrative_rules
from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES

ROOT = Path(__file__).resolve().parent


def _opinion():
    return {
        "application_no": "2024/005033",
        "applicant": "OSTİM TEKNİK ÜNİVERSİTESİ",
        "reference": "698947",
        "intro": "14.07.2026 tarihli Birinci Bildirimde, 1-3 numaralı istemlerin yenilik ve sanayiye uygulanabilirlik kriterlerini sağladığı, ancak buluş basamağı kriterini sağlamadığı belirtilmiştir. Başvuru sahibinin görüşleri aşağıda dikkatinize sunulmaktadır.",
        "cited_documents": [],
        "sections": [],
        "combined_assessment": {"heading": "", "paragraphs": []},
        "conclusion": ["Ayırt edici teknik fark, sensör verilerinin belirli bir işlevsel ilişki içinde değerlendirilmesidir. Teknik etki açıklanmıştır. Objektif teknik problem bu işlevsel ilişkinin nasıl sağlanacağıdır. Önceki teknikte bu yönde motivasyon veya yönlendirme bulunmamaktadır."],
        "signoff": "Saygılarımızla,\nDESTEK PATENT A.Ş.",
        "_output_language": "Türkçe",
        "_recipient_office": "TURKISH PATENT AND TRADEMARK OFFICE",
    }


def test_version_bumped():
    assert APP_VERSION == "v5.4.73"
    assert RULESET_VERSION == "2026-09-21.v65"


def test_binding_title_paragraphs_are_centered_and_gate_enforces_it():
    op = _opinion()
    data = build_gorus_docx(op)
    doc = Document(io.BytesIO(data))
    assert doc.paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert doc.paragraphs[1].alignment == WD_ALIGN_PARAGRAPH.CENTER
    validate_gorus_template_fidelity(data, ROOT / "Gorus_metni_696809_template.docx", op)

    doc.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    out = io.BytesIO(); doc.save(out)
    with pytest.raises(ValueError, match="hizalama"):
        validate_gorus_template_fidelity(out.getvalue(), ROOT / "Gorus_metni_696809_template.docx", op)


def test_internal_uppercase_how_label_is_rejected_but_natural_lowercase_is_allowed():
    op = _opinion()
    op["conclusion"] = ["Ayırt edici teknik farkın nasıl ortaya çıktığı unsur ilişkilerinden anlaşılmaktadır. Teknik etki vardır. Objektif teknik problem açıklanmıştır. Motivasyon ve yönlendirme bulunmamaktadır."]
    validate_opinion_narrative_rules(op, "buluş basamağı", "teknik katkı teknik etki objektif teknik problem motivasyon")
    op["conclusion"] = ["İstem 1'in elemanları arasındaki NASIL ilişkisi savunmayı desteklemektedir. Teknik etki vardır. Objektif teknik problem açıklanmıştır. Motivasyon ve yönlendirme bulunmamaktadır."]
    with pytest.raises(ValueError, match="iç-süreç etiketi"):
        validate_opinion_narrative_rules(op, "buluş basamağı", "teknik katkı teknik etki objektif teknik problem motivasyon")


def test_normal_flow_is_automatic_and_pre_generation_claim_revision_decision_is_not_blocking():
    app = (ROOT / "app.py").read_text(encoding="utf-8")
    assert 'revision_status = "Mevcut istem seti üzerinden otomatik revizyonsuz görüş"' in app
    assert 'if not (st.session_state.gorus_opinion_data and st.session_state.gorus_opinion_status == revision_status):' in app
    assert 'st.radio(\n                "İstem revizyonu kararı"' not in app
    assert 'st.button(f"{opinion_step}. Görüş metnini oluştur"' not in app
    assert "### Görüşü revize et" in app


def test_rules_require_auto_flow_and_internal_label_gate():
    assert "GÖRÜŞ OTOMATİK AKIŞ KAPISI" in GORUS_RULES
    assert "NİHAİ GÖRÜŞ İÇ-SÜREÇ ETİKETİ KAPISI" in GORUS_RULES
    assert "GÖRÜŞ KURUM BAŞLIĞI HİZALAMA KAPISI" in GORUS_RULES
