from pathlib import Path
import io
from docx import Document
from docx.oxml.ns import qn

import app_core as app


def test_sentence_start_after_newline_not_false_mid_sentence_case():
    text = "Önceki cümle.\nDerin paket inceleme motoru çalışır.\nTaşıyıcı sınıfı ağ adresi çevirisi uygulanır."
    assert app._mid_sentence_technical_case_violations(text) == []


def test_mid_sentence_generic_technical_terms_are_normalized():
    text = "Sistem Derin Paket İnceleme ve Control Plane bilgisini Katman 7 seviyesinde işler."
    out = app._normalize_known_mid_sentence_technical_case(text)
    assert "derin paket inceleme" in out
    assert "control plane" in out
    assert "katman 7" in out


def test_numbered_claim_has_direct_hanging_and_num_tab():
    doc = Document()
    template = Document(str(app.TARIFNAME_TEMPLATE))
    p = app.add_numbered_claim(doc, template, "Uzun bir istem giriş metni olup, özelliği;")
    ppr = p._p.pPr
    ind = ppr.find(qn("w:ind"))
    assert ind is not None
    assert ind.get(qn("w:left")) == "720"
    assert ind.get(qn("w:hanging")) == "360"
    tabs = ppr.find(qn("w:tabs"))
    assert tabs is not None
    num_tabs = [t for t in tabs.findall(qn("w:tab")) if t.get(qn("w:val")) == "num"]
    assert len(num_tabs) == 1
    assert num_tabs[0].get(qn("w:pos")) == "720"


def test_text_heavy_nonreference_numbering_is_cleaned_not_auto_excluded():
    audit = {
        "final_use": "include",
        "text_heavy_nonreference_numbering": True,
        "nonreference_numeric_or_step_marks": ["1", "A1", "B2"],
    }
    assert app._figure_audit_requests_exclusion(audit) is False


def test_dependent_system_claim_rejects_nominative_subject_before_possession_closure():
    import pytest
    bad = ["İstem 1’e uygun sistem olup, özelliği; derin paket inceleme motoru (20), paket başlıklarını inceleyen yapıya sahip olmasıdır."]
    with pytest.raises(ValueError, match="tamlayan/iyelik"):
        app._validate_dependent_system_claim_possessive_grammar(bad, "Türkçe")


def test_dependent_system_claim_accepts_genitive_subject_before_possession_closure():
    good = ["İstem 1’e uygun sistem olup, özelliği; derin paket inceleme motorunun (20), paket başlıklarını inceleyen yapıya sahip olmasıdır."]
    app._validate_dependent_system_claim_possessive_grammar(good, "Türkçe")


def test_dependent_system_claim_rejects_nominative_subject_before_contains_closure():
    import pytest
    bad = ["İstem 1’e uygun sistem olup, özelliği; taşıyıcı sınıfı ağ adresi çevirisi yönlendiricisi (40), dinamik port havuzu yapısı içermesidir."]
    with pytest.raises(ValueError, match="tamlayan/iyelik"):
        app._validate_dependent_system_claim_possessive_grammar(bad, "Türkçe")


def test_app_core_requires_sentence_case_and_method_how_audit_flags():
    source = Path(app.__file__).read_text(encoding="utf-8")
    assert '"method_how_steps_passed", "sentence_case_clean"' in source
