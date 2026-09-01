from pathlib import Path
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent

def test_version_and_rules():
    from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES
    assert APP_VERSION == "v5.4.41"
    assert RULESET_VERSION == "2026-09-01.v31"
    low = GORUS_RULES.casefold()
    assert "savunma dokümanları görülmeden" in low
    assert "yapılan değişiklikler ve dayanakları" in low
    assert "son markup" in low
    assert "destek patent" in low
    assert "işlem adımını" in low

def test_prompt_and_builder_order_contract():
    for fn in ["app_core.py", "app.py"]:
        src = (ROOT / fn).read_text(encoding="utf-8")
        assert '"amendment_assessment"' in src
        assert 'İstemlerde Yapılan Değişiklikler ve Dayanakları' in src
        assert 'wrapper.set(qn("w:author"), "Destek Patent")' in src
        build_pos = src.index('amendment = opinion.get("amendment_assessment") or {}', src.index('def build_gorus_docx'))
        docs_pos = src.index('docs = opinion.get("cited_documents") or []', src.index('def build_gorus_docx'))
        assert build_pos < docs_pos

def test_revision_author_is_destek_patent():
    from app_core import _append_revision
    from docx.oxml import OxmlElement
    parent = OxmlElement("w:p")
    _append_revision(parent, "x", kind="insert", change_id=1)
    ins = next(iter(parent))
    assert ins.get(qn("w:author")) == "Destek Patent"

def test_amendment_quotes_are_in_location_iterator():
    from gorus_audit import _iter_quote_objects
    opinion = {
        "amendment_assessment": {"blocks": [{"type": "quote", "text": "basis"}]},
        "sections": [{"blocks": [{"type": "quote", "text": "defence"}]}],
    }
    assert [x["text"] for x in _iter_quote_objects(opinion)] == ["basis", "defence"]
