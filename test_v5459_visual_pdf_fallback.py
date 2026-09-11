import io

import fitz
import pytest
from PIL import Image

import app_core


def _image_only_pdf(page_count=2):
    im = Image.new("RGB", (800, 1100), "white")
    bio = io.BytesIO(); im.save(bio, format="PNG")
    png = bio.getvalue()
    pdf = fitz.open()
    for _ in range(page_count):
        page = pdf.new_page(width=595, height=842)
        page.insert_image(page.rect, stream=png)
    data = pdf.tobytes(); pdf.close()
    return data


def _text_pdf(text):
    pdf = fitz.open(); page = pdf.new_page()
    page.insert_text((72,72), text)
    data = pdf.tobytes(); pdf.close(); return data


def test_image_only_pdf_uses_visual_fallback(monkeypatch):
    calls=[]
    def fake_ask(prompt, *, web_search=False, images=None):
        calls.append((prompt, list(images or [])))
        return {"pages":[{"page":1,"text":"PATENT SAYFA BIR"},{"page":2,"text":"CLAIM 1 TEST"}]}
    monkeypatch.setattr(app_core, "ask_json", fake_ask)
    asset=app_core.UploadedAsset("scan.pdf", _image_only_pdf(2), "application/pdf")
    out=app_core.extract_text_from_asset(asset)
    assert "PATENT SAYFA BIR" in out and "CLAIM 1 TEST" in out
    assert len(calls)==1 and len(calls[0][1])==2


def test_text_pdf_does_not_use_visual_fallback(monkeypatch):
    def fail(*args, **kwargs):
        raise AssertionError("visual fallback should not run")
    monkeypatch.setattr(app_core, "_visual_pdf_text", fail)
    text = "This is selectable patent text. " * 30
    asset=app_core.UploadedAsset("normal.pdf", _text_pdf(text), "application/pdf")
    out=app_core.extract_text_from_asset(asset)
    assert "selectable patent text" in out


def test_visual_fallback_is_fail_closed_when_model_returns_no_text(monkeypatch):
    monkeypatch.setattr(app_core, "ask_json", lambda *a, **k: {"pages":[]})
    asset=app_core.UploadedAsset("scan.pdf", _image_only_pdf(1), "application/pdf")
    with pytest.raises(ValueError, match="görsel PDF fallback"):
        app_core.extract_text_from_asset(asset)


def test_sparse_pdf_threshold_scales_with_page_count():
    assert app_core._pdf_needs_visual_fallback("", 10)
    assert app_core._pdf_needs_visual_fallback("a"*200, 10)
    assert not app_core._pdf_needs_visual_fallback("a"*500, 10)
