from __future__ import annotations

import io
from PIL import Image, ImageDraw
from docx import Document
from docx.oxml.ns import qn
import pytest

from rules import APP_VERSION, RULESET_VERSION, TARIFNAME_RULES
import app_core
from source_guards import collect_required_exact_technical_phrases, validate_required_exact_technical_phrases
from tarifname_figure_generation import has_material_gray_fill, normalize_grayscale_line_art


def test_v5465_versions_and_binding_rules():
    assert APP_VERSION == "v5.4.75"
    assert RULESET_VERSION == "2026-09-22.v67"
    low = TARIFNAME_RULES.casefold()
    assert "source_exact_phrases" in low or "open gateway / api sunucusu" in low
    assert "açık ağ geçidi (open gateway)" in low
    assert "utm sunucusu" in low
    assert "gri/gölgeli kutu dolgulu" in low
    assert "mergeformat" in low


def test_turkish_i_and_acronym_canonicalization():
    assert app_core._reference_sentence_case("İnsansız Hava Aracı") == "İnsansız hava aracı"
    assert app_core._inline_reference_name("İnsansız Hava Aracı") == "insansız hava aracı"
    assert app_core._reference_sentence_case("utm sunucusu") == "UTM sunucusu"
    assert app_core._reference_sentence_case("Utm Sunucusu") == "UTM sunucusu"
    title = app_core._normalize_turkish_invention_title(
        "6G ağlarında İnsansız hava aracı trafik yönetimi için dinamik QoD sistemi ve yöntemi"
    )
    assert "İnsansız" in title
    assert "QoD" in title
    assert "6G" in title


def test_turkish_terminology_requires_utm_expansion_and_open_gateway_translation():
    draft = {
        "technical_field": "Buluş, insansız hava aracı trafik yönetimi (UTM) için açık ağ geçidi (Open Gateway) üzerinden çalışan bir sistem ile ilgilidir.",
        "prior_art_general_paragraphs": [], "short_description_intro": "", "objectives": [],
        "unumbered_invention_definition": "", "unumbered_invention_features": [],
        "detailed_paragraphs": ["UTM sunucusu API çağrısı üretmektedir."], "working_principle": "",
        "system_claim": {}, "dependent_system_claims": [], "method_claim": {}, "dependent_method_claims": [], "abstract": "",
    }
    app_core._validate_turkish_terminology(draft, "Türkçe")
    bad = dict(draft)
    bad["detailed_paragraphs"] = ["utm sunucusu API çağrısı üretmektedir."]
    with pytest.raises(ValueError, match="kısaltma"):
        app_core._validate_turkish_terminology(bad, "Türkçe")
    bad2 = dict(draft)
    bad2["technical_field"] = "Buluş, insansız hava aracı trafik yönetimi (UTM) için Open Gateway üzerinden çalışan bir sistem ile ilgilidir."
    with pytest.raises(ValueError, match="açık ağ geçidi"):
        app_core._validate_turkish_terminology(bad2, "Türkçe")


def test_exact_customer_technical_phrase_is_preserved_fail_closed():
    registry = [
        {"passage_id":"B0001", "source":"BBF", "text":"Sistem büyük bir \"Milli İHA Yönetim ve Telekomünikasyon Ağı\" mimarisinde kullanılmaktadır."},
        {"passage_id":"B0002", "source":"BBF", "text":"Formu doldurunuz."},
    ]
    extracted = {
        "source_passage_audit":[
            {"passage_id":"B0001","classification":"technical","fact_ids":["T001"],"reason":""},
            {"passage_id":"B0002","classification":"nontechnical","fact_ids":[],"reason":"form talimatı"},
        ],
        "technical_facts":[{"id":"T001","statement":"Büyük ağ içinde kullanım","mandatory":True}],
    }
    phrases = collect_required_exact_technical_phrases(extracted, registry)
    assert "Milli İHA Yönetim ve Telekomünikasyon Ağı" in phrases
    validate_required_exact_technical_phrases(extracted, "Buluş, Milli İHA Yönetim ve Telekomünikasyon Ağı içerisinde kullanılabilir.")
    with pytest.raises(ValueError, match="teknik isimlendirme"):
        validate_required_exact_technical_phrases(extracted, "Buluş daha büyük bir ağ içerisinde kullanılabilir.")


def _gray_box_png() -> bytes:
    im = Image.new("RGB", (800, 360), "white")
    d = ImageDraw.Draw(im)
    d.rectangle((90, 70, 710, 290), fill=(238,238,238), outline=(0,0,0), width=5)
    d.line((90,180,710,180), fill=(0,0,0), width=4)
    out = io.BytesIO(); im.save(out, format="PNG"); return out.getvalue()


def test_gray_decorative_fill_is_detected_and_removed():
    raw = _gray_box_png()
    assert has_material_gray_fill(raw)
    fixed = normalize_grayscale_line_art(raw)
    assert not has_material_gray_fill(fixed)
    with Image.open(io.BytesIO(fixed)).convert("L") as im:
        assert im.getpixel((200, 120)) == 255
        assert im.getpixel((90, 70)) == 0


def test_figures_header_and_caption_are_explicit_arial_11_with_mergeformat():
    im = Image.new("RGB", (400, 180), "white")
    out = io.BytesIO(); im.save(out, format="PNG")
    data = app_core.build_figures_docx([app_core.UploadedAsset("figure.png", out.getvalue(), "image/png")], "Türkçe")
    app_core.validate_figures_docx_structure(data, {})
    doc = Document(io.BytesIO(data))
    hdr = doc.sections[0].header._element
    fields = [n for n in hdr.iter() if str(n.tag).endswith("}fldSimple")]
    assert fields
    assert all("MERGEFORMAT" in str(n.get(qn("w:instr")) or "") for n in fields)
    style = doc.styles["Header"]
    assert style.font.name == "Arial"
    caps = [p for p in doc.paragraphs if p.text.startswith("ŞEKİL ")]
    assert caps and caps[0].runs
    rpr = caps[0].runs[0]._r.rPr
    fonts = rpr.find(qn("w:rFonts"))
    assert fonts is not None
    assert all(fonts.get(qn(a)) == "Arial" for a in ("w:ascii","w:hAnsi","w:eastAsia","w:cs"))
