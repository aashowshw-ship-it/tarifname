import io
import zipfile
from pathlib import Path

import pytest
from docx import Document
from lxml import etree

from claim_amendment import (
    validate_atomic_amendment,
    validate_amendment_audit,
    validate_selected_candidate_is_structurally_best,
)
from app_core import build_claim_revision_pair, build_gorus_docx
from gorus_audit import validate_gorus_docx_content_flow, validate_tracked_changes_against_plan
from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES


def _amendment(old, new, *, deletion_required=False, deletion_reason=""):
    return {
        "claim_number": "1",
        "reason": "OBJ-1 kapsamındaki açıklık itirazını teknik ölçüm yapısıyla karşılamak.",
        "basis_quote": "Cihazda bulunan koagülometrede ışık kaynağı ve algılayıcıdan oluşan sensörler bulunur.",
        "change_type": "technical",
        "edit_mode": "replace_minimal" if deletion_required else "insert_only",
        "deletion_required": deletion_required,
        "deletion_reason": deletion_reason,
        "addresses_objection_ids": ["OBJ-1"],
        "clarity_effect": "Ölçüm yapısını teknik unsurla somutlaştırır.",
        "scope_effect": "narrows",
        "prior_art_effect": "Savunulan teknik zinciri daha belirgin hale getirir.",
        "remaining_risk": "Parametrelerin ayrı hesap modeli tarifnamede ayrıca açıklanmamıştır.",
        "direct_support": [{
            "added_feature": "algılayıcı",
            "basis_quote": "Cihazda bulunan koagülometrede ışık kaynağı ve algılayıcıdan oluşan sensörler bulunur.",
        }],
        "old_text": old,
        "new_text": new,
    }


def _source_docx(claim_text: str) -> bytes:
    d = Document()
    d.add_paragraph("TARİFNAME")
    d.add_paragraph("Cihazda bulunan koagülometrede ışık kaynağı ve algılayıcıdan oluşan sensörler bulunur.")
    d.add_paragraph("İSTEMLER")
    d.add_paragraph("1. " + claim_text)
    d.add_paragraph("ÖZET")
    d.add_paragraph("Özet metni")
    b = io.BytesIO(); d.save(b); return b.getvalue()


def _revision_texts(docx_bytes: bytes):
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
        root = etree.fromstring(zf.read("word/document.xml"))
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    ins = ["".join(x.xpath('.//w:t/text()', namespaces=ns)) for x in root.xpath('.//w:ins', namespaces=ns)]
    dele = ["".join(x.xpath('.//w:delText/text()', namespaces=ns)) for x in root.xpath('.//w:del', namespaces=ns)]
    return ins, dele


def test_v5477_versions_and_hard_rules():
    assert APP_VERSION == "v5.4.78"
    assert RULESET_VERSION == "2026-09-25.v70"
    low = GORUS_RULES.casefold()
    assert "95. ATOMİK İSTEM REVİZYONU" in GORUS_RULES
    assert "insertion-first" in low
    assert "bağımsız amendment auditor" in low
    assert "üçlü bütünlük" in low


def test_insertion_only_preserves_existing_claim_words():
    spec = "Cihazda bulunan koagülometrede ışık kaynağı ve algılayıcıdan oluşan sensörler bulunur. üzerinde sensörler ve ışık kaynağı bulunan"
    item = _amendment(
        "üzerinde sensörler ve ışık kaynağı bulunan",
        "üzerinde algılayıcı içeren sensörler ve ışık kaynağı bulunan",
    )
    receipt = validate_atomic_amendment(item, spec, known_objection_ids=["OBJ-1"], require_receipts=True)
    assert receipt["has_deletion"] is False
    assert receipt["has_insertion"] is True


def test_rewrite_is_rejected_when_original_words_can_be_preserved():
    spec = "Cihazda bulunan koagülometrede ışık kaynağı ve algılayıcıdan oluşan sensörler bulunur. üzerinde sensörler ve ışık kaynağı bulunan"
    item = _amendment(
        "üzerinde sensörler ve ışık kaynağı bulunan",
        "üzerinde ışık kaynağı ve algılayıcı içeren sensörler bulunan",
        deletion_required=True,
        deletion_reason="Uzman açıklık itirazını karşılamak için mevcut sıralamanın değiştirilmesi gerektiği ileri sürülmektedir.",
    )
    with pytest.raises(ValueError):
        validate_atomic_amendment(item, spec, known_objection_ids=["OBJ-1"], require_receipts=True)


def test_large_phrase_rewrite_is_blocked_even_if_deletion_claimed_required():
    old = "üzerinde sensörler ve ışık kaynağı bulunan sensör aparatının arasına parmak yerleştirilerek ölçüm sağlayan üzerinden kızılötesi ışın yayarak kanın değerlerini ölçen"
    spec = "Cihazda bulunan koagülometrede ışık kaynağı ve algılayıcıdan oluşan sensörler bulunur. " + old
    new = "algılayıcı ve ışık kaynağı taşıyan çoklu dalga boylu optik sensör düzeniyle parmak dokusunu analiz ederek koagülasyon verisi oluşturan"
    item = _amendment(old, new, deletion_required=True, deletion_reason="Uzmanın açıklık itirazında sonuç dilinin teknik yapı ile değiştirilmesi gerektiği için bu ifade zorunlu olarak değiştirilmelidir.")
    with pytest.raises(ValueError):
        validate_atomic_amendment(item, spec, known_objection_ids=["OBJ-1"], require_receipts=True)


def test_track_changes_pair_has_only_the_inserted_phrase():
    claim = "Bir cihaz olup, özelliği; üzerinde sensörler ve ışık kaynağı bulunan koagülometre içermesidir."
    source = _source_docx(claim)
    amendment = {
        "claim_number": "1",
        "old_text": "üzerinde sensörler ve ışık kaynağı bulunan",
        "new_text": "üzerinde algılayıcı içeren sensörler ve ışık kaynağı bulunan",
    }
    markup, clean = build_claim_revision_pair(source, [amendment])
    ins, dele = _revision_texts(markup)
    assert dele == []
    assert ins == ["algılayıcı içeren "]
    validate_tracked_changes_against_plan(source, markup, clean, [amendment])


def test_amendment_auditor_must_confirm_candidate_selection():
    amendment = _amendment("üzerinde sensörler bulunan", "üzerinde algılayıcı içeren sensörler bulunan")
    analysis = {
        "amendments": [amendment],
        "amendment_candidates": [
            {"candidate_id": "A1", "addresses_objection_ids": ["OBJ-1"], "amendments": [amendment]},
            {"candidate_id": "A2", "addresses_objection_ids": ["OBJ-1"], "amendments": [amendment]},
        ],
        "selected_candidate_id": "A1",
    }
    audit = {
        "overall_pass": True,
        "overall_assessment": "",
        "selected_candidate_confirmed": False,
        "selection_basis": "",
        "candidate_comparison": [],
        "items": [{
            "claim_number": "1", "direct_basis_pass": True, "addresses_examiner_issue": "yes",
            "minimal_edit_pass": True, "unnecessary_deletion": False, "clarity_effect": "clear",
            "scope_effect": "narrows", "prior_art_effect": "useful", "new_ambiguity": False,
            "remaining_risk": "residual",
        }],
        "blocking_reasons": [],
    }
    with pytest.raises(ValueError):
        validate_amendment_audit(audit, analysis)


def test_gorus_word_order_is_docs_then_amendment_then_defence():
    opinion = {
        "application_no": "2025/1", "applicant": "A", "reference": "R", "intro": "Araştırma raporunda gösterilen benzer dokümanlar aşağıdadır:",
        "cited_documents": [{"label": "D1", "number": "US123"}],
        "amendment_assessment": {"heading": "İstemlerde Yapılan Değişiklikler ve Dayanakları", "blocks": [{"type": "paragraph", "text": "İstem 1'e algılayıcı özelliği eklenmiştir."}]},
        "sections": [{"label": "D1", "heading": "D1 (US123) dokümanı:", "use_figure": False, "blocks": [{"type": "paragraph", "text": "Teknik değerlendirme paragrafıdır."}], "novelty_paragraphs": [], "inventive_step_paragraphs": []}],
        "combined_assessment": {"heading": "", "paragraphs": [], "groups": []},
        "conclusion": ["Sonuç paragrafıdır."], "signoff": "Saygılarımızla,\nDESTEK PATENT A.Ş.",
        "_output_language": "Türkçe", "_recipient_office": "TURKISH PATENT AND TRADEMARK OFFICE",
    }
    data = build_gorus_docx(opinion)
    validate_gorus_docx_content_flow(data)
    d = Document(io.BytesIO(data))
    texts = [p.text.strip() for p in d.paragraphs]
    assert texts.index("D1: US123") < texts.index("İstemlerde Yapılan Değişiklikler ve Dayanakları") < texts.index("D1 (US123) dokümanı:")


def test_word_flow_gate_rejects_amendment_before_d_bibliography():
    d = Document()
    d.add_paragraph("İstemlerde Yapılan Değişiklikler ve Dayanakları")
    d.add_paragraph("İstem 1'e algılayıcı özelliği eklenmiştir.")
    d.add_paragraph("D1: US123")
    d.add_paragraph("D1 (US123) dokümanı:")
    b = io.BytesIO(); d.save(b)
    with pytest.raises(ValueError, match="D1/D2/.+bibliyografik"):
        validate_gorus_docx_content_flow(b.getvalue())


def test_word_flow_gate_rejects_internal_amendment_meta_language():
    d = Document()
    d.add_paragraph("D1: US123")
    d.add_paragraph("İstemlerde Yapılan Değişiklikler ve Dayanakları")
    d.add_paragraph("İstem 1 minimum ölçüde değiştirilmiştir.")
    d.add_paragraph("D1 (US123) dokümanı:")
    b = io.BytesIO(); d.save(b)
    with pytest.raises(ValueError, match="iç süreç/meta"):
        validate_gorus_docx_content_flow(b.getvalue())


def test_claim_amendment_has_dedicated_final_compliance_receipts():
    from rules import FINAL_COMPLIANCE_REQUIRED_CHECKS, final_compliance_gate
    assert FINAL_COMPLIANCE_REQUIRED_CHECKS["claim_amendment"] == ("atomic_plan", "amendment_integrity")
    source = _source_docx("Bir cihaz olup, özelliği; üzerinde sensörler ve ışık kaynağı bulunan koagülometre içermesidir.")
    amendment = {"claim_number": "1", "old_text": "üzerinde sensörler ve ışık kaynağı bulunan", "new_text": "üzerinde algılayıcı içeren sensörler ve ışık kaynağı bulunan"}
    markup, _ = build_claim_revision_pair(source, [amendment])
    with pytest.raises(ValueError):
        final_compliance_gate("claim_amendment", data=markup, output_name="x.docx", default_name="x.docx", checks={"atomic_plan": True})
    assert final_compliance_gate("claim_amendment", data=markup, output_name="x.docx", default_name="x.docx", checks={"atomic_plan": True, "amendment_integrity": True}) == "x.docx"


def test_revised_opinion_delivery_rebuilds_page_line_authority_from_final_markup():
    app = (Path(__file__).resolve().parent / "app.py").read_text(encoding="utf-8")
    assert 'if revision_status.startswith("Kullanıcı tarafından onaylanmış revize") and st.session_state.gorus_markup_data:' in app
    assert '_delivery_authority_name, _delivery_authority_bytes = prepare_line_reference_source(' in app
    assert '_delivery_final_spec_name = "son_markup_tarifname.docx"' in app
