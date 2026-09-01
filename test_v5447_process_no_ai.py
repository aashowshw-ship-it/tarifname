from pathlib import Path
from docx import Document

from processes import (
    extract_application_information_rule_based,
    extract_application_source_text,
    application_precheck_missing,
)

ROOT = Path(__file__).resolve().parent


def test_rule_based_application_extraction_without_ai():
    source = """Başvuru Türü: Patent
DP Referans: 181999
Buluş Başlığı: Otomatik Test Sistemi
Hak Sahibi: ABC Otomotiv San. ve Tic. A.Ş.
VKN: 1234567890
Ülke: Türkiye
İl: Bursa
Adres: Nilüfer Bursa
Buluş Sahibi: Ali Veli
TCKN: 12345678901
Ülke: Türkiye
İl: Bursa
Adres: Osmangazi Bursa
Rüçhan: Yok
"""
    data = extract_application_information_rule_based(
        [("beyan.txt", source)],
        specification_text="TARİFNAME\nOtomatik Test Sistemi\nTEKNİK ALAN",
    )
    assert data["application_kind"] == "Patent"
    assert data["reference"] == "181999"
    assert data["invention_title"] == "Otomatik Test Sistemi"
    assert data["applicants"][0]["name"] == "ABC Otomotiv San. ve Tic. A.Ş."
    assert data["applicants"][0]["city"] == "Bursa"
    assert data["inventors"][0]["name"] == "Ali Veli"
    assert data["priority"]["status"] == "Yok"


def test_docx_four_column_form_is_split_into_pairs(tmp_path):
    doc = Document()
    table = doc.add_table(rows=3, cols=4)
    table.rows[0].cells[0].text = "Hak Sahibi"
    table.rows[0].cells[1].text = "ABC A.Ş."
    table.rows[0].cells[2].text = "Ülke"
    table.rows[0].cells[3].text = "Türkiye"
    table.rows[1].cells[0].text = "Adres"
    table.rows[1].cells[1].text = "Bursa"
    table.rows[1].cells[2].text = "İl"
    table.rows[1].cells[3].text = "Bursa"
    table.rows[2].cells[0].text = "Buluş Sahibi"
    table.rows[2].cells[1].text = "Ali Veli"
    table.rows[2].cells[2].text = "Rüçhan"
    table.rows[2].cells[3].text = "Yok"
    path = tmp_path / "beyan.docx"
    doc.save(path)
    text = extract_application_source_text(path.name, path.read_bytes())
    data = extract_application_information_rule_based([(path.name, text)])
    assert data["applicants"][0]["name"] == "ABC A.Ş."
    assert data["applicants"][0]["country"] == "Türkiye"
    assert data["applicants"][0]["city"] == "Bursa"
    assert data["inventors"][0]["name"] == "Ali Veli"
    assert data["priority"]["status"] == "Yok"


def test_freeform_email_only_accepts_explicit_roles():
    text = "Patent başvurusu yapılacaktır. Hak sahibi olarak ABC Otomotiv A.Ş. olacaktır. Buluş sahibi Ali Veli olacaktır. Rüçhan yoktur."
    data = extract_application_information_rule_based([("mail.txt", text)])
    assert data["application_kind"] == "Patent"
    assert data["applicants"][0]["name"] == "ABC Otomotiv A.Ş."
    assert data["inventors"][0]["name"] == "Ali Veli"
    assert data["priority"]["status"] == "Yok"


def test_ambiguous_company_instruction_is_not_guessed():
    text = "Bu dosyada geçen şirket yerine yeni kurduğumuz XYZ üzerinden ilerleyelim."
    data = extract_application_information_rule_based([("mail.txt", text)])
    assert data["applicants"] == []


def test_missing_gate_still_blocks_rule_based_missing_data():
    data = extract_application_information_rule_based([("mail.txt", "Başvuru Türü: Patent\nRüçhan: Yok")])
    metrics = {"specification_pages": 7, "claim_count": 12, "abstract_present": True, "figures_pages": 0}
    missing = application_precheck_missing(data, metrics)
    assert "buluş başlığı" in missing
    assert any("hak sahibi" in x for x in missing)
