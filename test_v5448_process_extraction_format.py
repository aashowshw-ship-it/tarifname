import io
import zipfile
from pathlib import Path

from docx import Document

from processes import (
    extract_application_information_rule_based,
    extract_application_source_text,
    reference_from_filename,
    split_patent_docx,
    strip_template_colored_text,
)
from rules import APP_VERSION

ROOT = Path(__file__).resolve().parent


def _sample_form() -> bytes:
    doc = Document()
    table = doc.add_table(rows=0, cols=2)
    rows = [
        ("HAK SAHİBİ BİLGİLERİ", ""),
        ("Adı / Unvanı", "ÖRNEK OTOMOTİV SANAYİ VE TİCARET A.Ş."),
        ("VKN", "1234567890"),
        ("Adres", "Nilüfer / Bursa / Türkiye"),
        ("Ülke", "Türkiye"),
        ("İl", "Bursa"),
        ("BULUŞU YAPAN BİLGİLERİ", ""),
        ("Adı Soyadı", "Ayşe Yılmaz"),
        ("T.C. Kimlik No", "12345678901"),
        ("Adres", "Osmangazi / Bursa / Türkiye"),
        ("Ülke", "Türkiye"),
        ("İl", "Bursa"),
    ]
    for a, b in rows:
        cells = table.add_row().cells
        cells[0].text = a
        cells[1].text = b
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def test_v5448_structured_form_roles_and_filename_reference():
    assert APP_VERSION == "v5.4.52"
    form = _sample_form()
    text = extract_application_source_text("Beyan_Formu.docx", form)
    data = extract_application_information_rule_based(
        [("Beyan_Formu.docx", text), ("mail.txt", "Başvuru türü: Patent\nRüçhan: Yok\nE-posta: patent@example.com")],
        specification_text="TARİFNAME\nÖrnek Buluş Başlığı\nTEKNİK ALAN\n...",
        specification_filename="Tarifname_181176_rev6_SON.docx",
    )
    assert data["reference"] == "181176"
    assert data["field_sources"]["reference"] == "Tarifname dosya adı"
    assert data["application_kind"] == "Patent"
    assert data["priority"]["status"] == "Yok"
    assert data["applicants"][0]["name"] == "ÖRNEK OTOMOTİV SANAYİ VE TİCARET A.Ş."
    assert data["applicants"][0]["identity"] == "1234567890"
    assert data["applicants"][0]["country"] == "Türkiye"
    assert data["inventors"][0]["name"] == "Ayşe Yılmaz"
    assert data["inventors"][0]["identity"] == "12345678901"
    assert any(x["label"] == "E-posta" and x["value"] == "patent@example.com" for x in data["other_information"])


def test_v5448_reference_filename_variants():
    assert reference_from_filename("Tarifname_DP-696809_rev2.docx") == "696809"
    assert reference_from_filename("181612_Tarifname.docx") == "181612"
    assert reference_from_filename("Tarifname_final.docx") == ""


def test_v5448_docx_package_style_parts_are_untouched():
    src = (ROOT / "Tarifname_181176_template.docx").read_bytes()
    cleaned = strip_template_colored_text(src)
    split = split_patent_docx(src)
    protected = ["word/styles.xml", "word/numbering.xml", "word/settings.xml"]
    with zipfile.ZipFile(io.BytesIO(src)) as zsrc:
        for payload in [cleaned, *split.values()]:
            with zipfile.ZipFile(io.BytesIO(payload)) as zout:
                for name in protected:
                    if name in zsrc.namelist() and name in zout.namelist():
                        assert zout.read(name) == zsrc.read(name)
                for name in zsrc.namelist():
                    if name.startswith("word/header") or name.startswith("word/footer"):
                        assert zout.read(name) == zsrc.read(name)


def test_v5448_render_installs_real_arial_for_pdf_conversion():
    docker = (ROOT / "Dockerfile").read_text(encoding="utf-8")
    assert "ttf-mscorefonts-installer" in docker
    assert "fc-match Arial" in docker
