import io, zipfile
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pytest
from gorus_audit import validate_minimal_tracked_changes, validate_ep_prior_art_markup_text
from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES


def _docx_with_pair(old: str, new: str) -> bytes:
    d=Document(); p=d.add_paragraph()
    delete=OxmlElement('w:del'); delete.set(qn('w:id'),'1')
    r=OxmlElement('w:r'); t=OxmlElement('w:delText'); t.text=old; r.append(t); delete.append(r)
    ins=OxmlElement('w:ins'); ins.set(qn('w:id'),'2')
    r2=OxmlElement('w:r'); t2=OxmlElement('w:t'); t2.text=new; r2.append(t2); ins.append(r2)
    p._p.append(delete); p._p.append(ins)
    b=io.BytesIO(); d.save(b); return b.getvalue()


def test_minimal_redline_accepts_article_only():
    validate_minimal_tracked_changes(_docx_with_pair('the','a'))


def test_minimal_redline_rejects_unchanged_trigger_rewritten():
    with pytest.raises(ValueError):
        validate_minimal_tracked_changes(_docx_with_pair('the trigger','a trigger'))


def test_ep_prior_art_requires_however_and_no_d_labels():
    spec='trigger embedded in the actors generates zero-knowledge proofs application authentication authorization interaction'
    good=['As a result of the research on the subject, application numbered US2020/0145229 A1 has been found. The application is related to blockchain authentication. However, the application does not mention a trigger embedded in the actors that generates zero-knowledge proofs.']
    validate_ep_prior_art_markup_text(good,spec)
    with pytest.raises(ValueError):
        validate_ep_prior_art_markup_text(['As a result of the research on the subject, D2 US2020/0145229 A1 has been found. However, trigger embedded in actors.'],spec)
    with pytest.raises(ValueError):
        validate_ep_prior_art_markup_text(['As a result of the research on the subject, application numbered US2020/0145229 A1 has been found. The application is related to blockchain authentication.'],spec)


def test_v5427_rules_present():
    assert APP_VERSION=='v5.4.41'
    assert RULESET_VERSION=='2026-09-01.v31'
    low=GORUS_RULES.casefold()
    assert 'minimum track changes' in low
    assert 'however,' in low
    assert 'yalnız araştırma raporunda x veya y' in low
