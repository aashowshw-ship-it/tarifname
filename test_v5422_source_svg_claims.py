from __future__ import annotations

import io
import zipfile

import pytest

from pathlib import Path
import cairosvg
from source_guards import build_source_passage_registry, validate_source_passage_audit, resolve_tarifname_claim_mode
from validators import validate_draft


def test_auto_mode_forces_system_and_method_when_both_have_source_basis():
    extracted = {
        "recommended_claim_mode": "Yalnızca yöntem",
        "has_system_basis": True,
        "has_method_basis": True,
        "elements": [{"number": "1", "name": "Platform karakterizasyon modülü"}, {"number": "2", "name": "Hedef tanımlama modülü"}],
        "method_steps": [{"number": "1001", "text": "Parametrelerin elde edilmesi"}, {"number": "1002", "text": "Sonucun hesaplanması"}],
    }
    assert resolve_tarifname_claim_mode(extracted, "BBF'ye göre otomatik belirle") == "Sistem ve yöntem"


def _minimal_method_draft(dependent: str):
    return {
        "elements": [],
        "method_steps": [
            {"number": "1001", "text": "Bir elektronik işlem birimi vasıtasıyla verinin elde edilmesi"},
            {"number": "1002", "text": "Bir elektronik işlem birimi vasıtasıyla sonucun hesaplanması"},
        ],
        "method_claim": {
            "preamble": "Bir elektronik işlem birimi vasıtasıyla gerçekleştirilen yöntem",
            "steps": [
                "Bir elektronik işlem birimi vasıtasıyla verinin elde edilmesi (1001)",
                "Bir elektronik işlem birimi vasıtasıyla sonucun hesaplanması (1002)",
            ],
            "closing": "işlem adımlarını içermesidir.",
        },
        "dependent_method_claims": [dependent],
        "dependent_system_claims": [],
        "system_claim": None,
        "tables": [],
    }


def test_dependent_method_bad_ending_is_rejected():
    findings = validate_draft(_minimal_method_draft("İstem 1'e uygun yöntem olup, özelliği; sonucun ayrıca sınıflandırılmasıdır."))
    assert any("Bağımlı yöntem istemi" in x["message"] and "işlem adımını" in x["message"] for x in findings)


@pytest.mark.parametrize("ending", ["işlem adımını içermesidir.", "işlem adımlarını içermesidir."])
def test_dependent_method_valid_step_endings_are_not_rejected_for_ending(ending):
    claim = f"İstem 1'e uygun yöntem olup, özelliği; sonucun ayrıca sınıflandırılması {ending}"
    findings = validate_draft(_minimal_method_draft(claim))
    assert not any("Bağımlı yöntem istemi" in x["message"] and "işlem adımını" in x["message"] for x in findings)


def test_active_app_keeps_svg_in_zip_whitelist_and_cairosvg_renders():
    app_source = (Path(__file__).with_name("app.py")).read_text(encoding="utf-8")
    assert '".svg"' in app_source
    assert 'inner_suffix not in {".pdf", ".docx", ".doc", ".txt", ".md", ".png", ".jpg", ".jpeg", ".webp", ".svg"}' in app_source
    assert 'technical_figure_assets.extend(extract_embedded_images(asset))' in app_source
    svg = b'<svg xmlns="http://www.w3.org/2000/svg" width="300" height="200"><rect width="300" height="200" fill="white"/><circle cx="150" cy="100" r="50" fill="none" stroke="black"/></svg>'
    png = cairosvg.svg2png(bytestring=svg, output_width=600)
    assert png.startswith(b"\x89PNG")

def test_source_passage_audit_requires_every_raw_passage_and_valid_mapping():
    registry = build_source_passage_registry("Görev hızı hesaplanır.\nİmza:", "Jamming seviyesi kullanılır.")
    extracted = {"technical_facts": [{"id": "T001"}, {"id": "T002"}], "source_passage_audit": []}
    with pytest.raises(ValueError, match="eksik"):
        validate_source_passage_audit(extracted, registry)

    rows=[]
    for rec in registry:
        if "İmza" in rec["text"]:
            rows.append({"passage_id":rec["passage_id"],"classification":"nontechnical","fact_ids":[],"reason":"İdari imza alanı"})
        elif "Jamming" in rec["text"]:
            rows.append({"passage_id":rec["passage_id"],"classification":"technical","fact_ids":["T002"],"reason":""})
        else:
            rows.append({"passage_id":rec["passage_id"],"classification":"technical","fact_ids":["T001"],"reason":""})
    extracted["source_passage_audit"] = rows
    validate_source_passage_audit(extracted, registry)


def test_claim_and_abstract_page_break_regression():
    src = Path(__file__).with_name("app.py").read_text(encoding="utf-8")
    # Şablon paragrafı 98 manuel page break içerir; istemler arasında veya ÖZET
    # öncesindeki normal boşluk olarak kullanılmamalıdır.
    assert "tpl_blank(98)" not in src
    assert "working_p = tpl_text(74" in src
    assert 'br.get(qn("w:type")) == "page"' in src


def test_reference_list_preserves_exact_element_name_case():
    src = Path(__file__).with_name("app.py").read_text(encoding="utf-8")
    assert "_reference_sentence_case(element.get('name',''))" not in src
    assert "str(element.get('name','') or '').strip()" in src


def test_svg_rasterization_normalizes_unicode_minus_without_changing_values():
    src = Path(__file__).with_name("app.py").read_text(encoding="utf-8")
    assert '.replace("−".encode("utf-8"), b"-")' in src
    assert '.replace("–".encode("utf-8"), b"-")' in src


def test_short_description_has_physical_blank_paragraph_regression():
    src = Path(__file__).with_name("app.py").read_text(encoding="utf-8")
    assert 'tpl_blank(15)\n\n    # BULUŞUN KISA AÇIKLAMASI' in src
