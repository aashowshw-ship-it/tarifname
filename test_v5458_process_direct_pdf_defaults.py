from pathlib import Path
import shutil

import fitz

from processes import (
    application_needs_semantic_ai,
    build_epats_application_package,
    extract_application_information_rule_based,
)
from rules import APP_VERSION

ROOT = Path(__file__).resolve().parent


def test_v5458_version():
    assert APP_VERSION == "v5.4.58"


def test_v5458_rule_parser_reads_direct_name_company_address_and_mail_options():
    applicant = """[[TABLO 1]]
Unvanı\tTT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ\tSahip Türü\tTüzel kişi
VKN\t8590380323\tÜlke\tseçilmedi
Adres\t(seçilmedi, GAYRETTEPE MAH. YILDIZ POSTA CAD. TÜRKTELEKOM GENEL MÜDÜRLÜĞÜ 40 TÜRKİYE
E-posta\tkubilay.aydin@turktelekom.com.tr\tTelefon\t2164606574
"""
    inventor = """[[TABLO 2]]
Ad Soyad\tGürkan Erkoç
TCKN\t54796123412
Doğum Tarihi\t23.05.1992
İl/İlçe\tİstanbul/Ümraniye
Adres\tFatih Sultan Mehmet Mah, Balkan Cd. No:49, 34771 Ümraniye/İstanbul
E-posta\tgurkan.erkoc@turktelekom.com.tr
Telefon\t5552550437
"""
    mail = """1. Başvuru esnasında buluşçu bilgileri gizlensin mi? (Cevap verilmemesi halinde HAYIR olarak işleme alınacaktır.) à HAYIR
2. Buluş, TÜBİTAK, KOSGEB vb. bir kamu kurum tarafından desteklenen bir proje kapsamında mı ortaya çıktı? (Cevap verilmemesi halinde destek alınmadığı varsayılarak başvuru yapılacaktır.) à HAYIR
3. Erken yayın talep ediliyor mu? (Cevap verilmemesi halinde HAYIR olarak işleme alınacaktır.) à EVET
"""
    out = extract_application_information_rule_based(
        [("Adsız.png", applicant), ("beyan.docx", inventor), ("mail.msg", mail)],
        specification_text="TARİFNAME\nHologramlı görüşme yöntemi\nTEKNİK ALAN",
        specification_filename="Tarifname_181140.docx",
    )
    app = out["applicants"][0]
    inv = out["inventors"][0]
    assert app["name"] == "TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ"
    assert app["identity"] == "8590380323"
    assert app["country"] == "Türkiye"
    assert app["address"].startswith("GAYRETTEPE MAH.")
    assert "seçilmedi" not in app["address"].casefold()
    assert inv["name"] == "Gürkan Erkoç"
    assert inv["identity"] == "54796123412"
    assert inv["phone"] == "5552550437"
    assert out["filing_options"]["inventor_hidden"]["status"] == "Hayır"
    assert out["filing_options"]["public_project"]["status"] == "Hayır"
    assert out["filing_options"]["early_publication"]["status"] == "Evet"
    assert out["priority"]["status"] == "Yok"
    assert out["priority"]["source"].startswith("Varsayılan")
    assert application_needs_semantic_ai(out) is False


def test_v5458_word_is_converted_once_then_split_as_full_pages_preserving_numbers():
    if shutil.which("libreoffice") is None:
        return
    data = (ROOT / "Tarifname_181176_template.docx").read_bytes()
    _zip, pdfs = build_epats_application_package(data, specification_name="Tarifname_181176_template.docx")

    # Full Word has pages 1..9; claims begin on Word page 7, abstract on page 9.
    assert len(fitz.open(stream=pdfs["Tarifname.pdf"], filetype="pdf")) == 6
    claims = fitz.open(stream=pdfs["Istemler.pdf"], filetype="pdf")
    abstract = fitz.open(stream=pdfs["Ozet.pdf"], filetype="pdf")
    try:
        ctext = claims[0].get_text("text")
        atext = abstract[0].get_text("text")
        # Original page numbers remain; no leading blank section page.
        assert ctext.splitlines()[0].strip() == "7"
        assert "İSTEMLER" in ctext
        assert atext.splitlines()[0].strip() == "9"
        assert "ÖZET" in atext
        # Word line numbering is preserved, while fake 1X/2X artifacts are absent.
        assert "5" in ctext.splitlines()
        assert "10" in ctext.splitlines()
        assert "1X" not in ctext and "2X" not in ctext
    finally:
        claims.close()
        abstract.close()
