import io
import zipfile
from pathlib import Path

from docx import Document

from processes import (
    extract_application_information_rule_based,
    extract_application_source_text,
    remove_word_line_numbering,
    reference_from_filename,
)
from rules import APP_VERSION

ROOT = Path(__file__).resolve().parent


def _realistic_beyan() -> bytes:
    doc = Document()
    t = doc.add_table(rows=8, cols=4)
    rows = [
        [
            "HAK SAHİBİ / BAŞVURU SAHİBİ (SAHİPLERİ):",
            "Not: Başvuru sahibinin birden fazla olması durumunda her bir başvuru sahibi için ayrı ayrı düzenlenmelidir.",
            "",
            "",
        ],
        ["Unvanı", "TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ", "VKN", "1234567890"],
        ["Adres", "GAYRETTEPE MAH. YILDIZ POSTA CAD. TÜRK TELEKOM GENEL MÜDÜRLÜĞÜ 40 TÜRKİYE", "İl", "İstanbul"],
        ["E-posta", "gurkan.erkoc@turktelekom.com.tr", "Telefon", "5552550437"],
        ["BULUŞ SAHİBİ BİLGİLERİ", "", "", ""],
        ["TCKN / Kimlik", "12345678901", "Doğum Tarihi", "01.02.1980"],
        ["Ad Soyad", "Gürkan Erkoç", "Ev/İş Telefonu", "5552550437"],
        ["İl/İlçe", "İstanbul/Ümraniye", "Adres", "Atakent Mah. Örnek Sok. No:1 İstanbul TÜRKİYE"],
    ]
    for row, vals in zip(t.rows, rows):
        for cell, value in zip(row.cells, vals):
            cell.text = value
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def test_v5449_real_beyan_note_is_not_person_and_fields_attach_to_same_people():
    assert APP_VERSION == "v5.4.55"
    raw = _realistic_beyan()
    text = extract_application_source_text("Hologramli Gorusme - b.beyan.docx", raw)
    data = extract_application_information_rule_based(
        [("Hologramli Gorusme - b.beyan.docx", text), ("mail.txt", "Buluş başlığı: Hologramlı Görüşme /\nBaşvuru türü: Patent\nRüçhan: Yok")],
        specification_text="TARİFNAME\nHologramlı görüşme yöntemi\nTEKNİK ALAN\n...",
        specification_filename="Tarifname_696809.docx",
    )

    assert data["reference"] == "696809"
    assert data["invention_title"] == "Hologramlı görüşme yöntemi"
    assert not any("Buluş başlığı" in x for x in data["conflicts"])

    assert len(data["applicants"]) == 1
    applicant = data["applicants"][0]
    assert applicant["name"] == "TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ"
    assert applicant["identity"] == "1234567890"
    assert applicant["country"] == "Türkiye"
    assert applicant["city"] == "İstanbul"
    assert "GAYRETTEPE" in applicant["address"]
    assert applicant["email"] == "gurkan.erkoc@turktelekom.com.tr"
    assert applicant["phone"] == "5552550437"
    assert "Not: Başvuru sahibinin" not in applicant["name"]

    assert len(data["inventors"]) == 1
    inventor = data["inventors"][0]
    assert inventor["identity"] == "12345678901"
    assert inventor["name"] == "Gürkan Erkoç"
    assert inventor["birth_date"] == "01.02.1980"
    assert inventor["city"] == "İstanbul"
    assert inventor["district"] == "Ümraniye"
    assert inventor["country"] == "Türkiye"
    assert inventor["phone"] == "5552550437"
    assert "Atakent" in inventor["address"]


def test_v5449_word_line_numbers_removed_without_touching_style_parts():
    src = (ROOT / "Tarifname_181176_template.docx").read_bytes()
    cleaned = remove_word_line_numbering(src)
    with zipfile.ZipFile(io.BytesIO(src)) as zsrc, zipfile.ZipFile(io.BytesIO(cleaned)) as zout:
        assert b"lnNumType" in zsrc.read("word/document.xml")
        assert b"lnNumType" not in zout.read("word/document.xml")
        for name in ["word/styles.xml", "word/numbering.xml", "word/settings.xml"]:
            assert zout.read(name) == zsrc.read(name)
        for name in zsrc.namelist():
            if name.startswith("word/header") or name.startswith("word/footer"):
                assert zout.read(name) == zsrc.read(name)


def test_v5449_filename_reference_still_drives_dp_ref():
    assert reference_from_filename("Tarifname_696809_rev3.pdf") == "696809"
