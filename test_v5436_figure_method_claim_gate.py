from __future__ import annotations

import io
from pathlib import Path

import pytest
from PIL import Image
from docx import Document

from rules import APP_VERSION, RULESET_VERSION, TARIFNAME_RULES
import app_core
from tarifname_figure_generation import (
    build_method_flow_png,
    protect_turkish_claim_transition,
    is_monochrome_enough,
)


def _bw_source_png() -> bytes:
    im = Image.new("RGB", (600, 220), "white")
    # a minimal black line and system refs; no method refs
    for x in range(80, 520):
        im.putpixel((x, 110), (0, 0, 0))
    out = io.BytesIO(); im.save(out, format="PNG"); return out.getvalue()


def test_versions_and_new_figure_rules():
    assert APP_VERSION == "v5.4.49"
    assert RULESET_VERSION == "2026-09-04.v39"
    assert "sistem şekline" in TARIFNAME_RULES.casefold()
    assert "ayrı yöntem/akış" in TARIFNAME_RULES.casefold()
    assert "arial 11" in TARIFNAME_RULES.casefold() and "kalın" in TARIFNAME_RULES.casefold()


def test_method_flow_is_separate_monochrome_and_referenced_in_docx_metadata():
    steps = [
        {"number": "1001", "text": "Verinin alınması"},
        {"number": "1002", "text": "Verinin düzenlenmesi"},
        {"number": "1003", "text": "Verinin iletilmesi"},
    ]
    method_png = build_method_flow_png(steps)
    assert is_monochrome_enough(method_png)
    source = app_core.UploadedAsset("source.png", _bw_source_png(), "image/png")
    method = app_core.UploadedAsset("generated_method_flow_1001_1002_1003.png", method_png, "image/png")
    data = app_core.build_figures_docx([source, method], "Türkçe")
    draft = {"method_steps": steps}
    app_core.validate_figures_docx_structure(data, draft)
    doc = Document(io.BytesIO(data))
    descr = [str(n.get("descr") or "") for n in doc._element.iter() if str(n.tag).endswith("}docPr")]
    assert any(x == "method_flow:1001,1002,1003" for x in descr)


def test_method_gate_rejects_figures_docx_without_separate_method_figure():
    source = app_core.UploadedAsset("source.png", _bw_source_png(), "image/png")
    data = app_core.build_figures_docx([source], "Türkçe")
    with pytest.raises(ValueError, match="ayrı yöntem/akış şekli"):
        app_core.validate_figures_docx_structure(data, {"method_steps": [{"number":"1001","text":"x"}]})


def test_claim_transition_uses_non_breaking_spaces():
    text = protect_turkish_claim_transition("Anormal durum tespit sistemi olup, özelliği;")
    assert "sistemi\u00a0olup,\u00a0özelliği;" in text
    assert "sistemi olup, özelliği;" not in text
