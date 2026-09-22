from __future__ import annotations

import io
from docx import Document
from docx.oxml.ns import qn

from rules import APP_VERSION, RULESET_VERSION, TARIFNAME_RULES
import app_core
from tarifname_figure_generation import protect_turkish_claim_transition


def test_versions_and_v5437_rules():
    assert APP_VERSION == "v5.4.75"
    assert RULESET_VERSION == "2026-09-22.v67"
    low = TARIFNAME_RULES.casefold()
    assert "ascii" in low and "hansi" in low and "eastasia" in low and "cs" in low
    assert "kısa/orphan" in low
    assert "doğal satır kaydırması" in low
    assert "non-breaking" in low and "yasak" in low


def test_claim_transition_preserves_natural_word_spacing():
    text = protect_turkish_claim_transition(
        "Hücrelerin haftalık davranışına göre anormal durum tespit sistemi olup, özelliği;"
    )
    assert "\u00a0" not in text
    assert text.endswith("tespit sistemi olup, özelliği;")


def test_figure_page_counter_fields_have_literal_arial_all_scripts():
    # One minimal raster figure is enough to exercise the Figures DOCX header.
    from PIL import Image
    img = Image.new("RGB", (300, 120), "white")
    out = io.BytesIO(); img.save(out, format="PNG")
    data = app_core.build_figures_docx([app_core.UploadedAsset("figure.png", out.getvalue(), "image/png")], "Türkçe")
    app_core.validate_figures_docx_structure(data, {})
    doc = Document(io.BytesIO(data))
    hdr = doc.sections[0].header._element
    fld_nodes = [n for n in hdr.iter() if str(n.tag).endswith("}fldSimple")]
    assert len(fld_nodes) >= 2
    for fld in fld_nodes:
        rfonts = [n for n in fld.iter() if str(n.tag).endswith("}rFonts")]
        assert rfonts
        for rf in rfonts:
            assert all(rf.get(qn(attr)) == "Arial" for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"))
