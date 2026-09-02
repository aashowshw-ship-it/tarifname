import io
import re
import zipfile

import fitz

from processes import (
    _pdf_epats_cleanup,
    extract_application_information_rule_based,
    merge_verified_cpu_ner_application_information,
    remove_word_header_page_numbers,
    split_patent_docx,
)
from rules import APP_VERSION


def test_v5456_version():
    assert APP_VERSION == "v5.4.58"


def test_v5456_structured_roles_and_mail_preferences_are_not_crossed():
    text = """HAK SAHİBİ / BAŞVURU SAHİBİ
Unvanı\tHak Sahibi Adresi\tSahip Türü\tUyruk\tTC Kimlik / Vergi No\tE-Posta\tTelefon
TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ\tGAYRETTEPE MAH. YILDIZ POSTA CAD. TÜRKTELEKOM GENEL MÜDÜRLÜĞÜ 40 TÜRKİYE\tTüzel\tTürkiye\t1234567890\tkubilay.aydin@turktelekom.com.tr\t02164606574

BULUŞ SAHİBİ
Ad Soyad\tTCKN / Kimlik\tÜlke\tİl\tİlçe\tAdres\tE-posta\tTelefon\tDoğum Tarihi
Gürkan Erkoç\t54796123412\tTürkiye\tİstanbul\tÜmraniye\tFatih Sultan Mehmet Mah, Balkan Cd. No:49, 34771 Ümraniye/İstanbul\tgurkan.erkoc@turktelekom.com.tr\t5552550437\t23.05.1992

1. Başvuru esnasında buluşçu bilgileri gizlensin mi? (Cevap verilmemesi halinde HAYIR olarak işleme alınacaktır.) à HAYIR
2. Buluş, TÜBİTAK, KOSGEB vb. bir kamu kurum tarafından desteklenen bir proje kapsamında mı ortaya çıktı? Aldıysanız lütfen kurumu ve proje numarasını belirtiniz. (Cevap verilmemesi halinde destek alınmadığı varsayılarak başvuru yapılacaktır.) à HAYIR
3. Erken yayın talep ediliyor mu? (Cevap verilmemesi halinde HAYIR olarak işleme alınacaktır.) à EVET
"""
    out = extract_application_information_rule_based(
        [("mail-ve-beyan.txt", text)],
        specification_text="TARİFNAME\nHologramlı görüşme yöntemi\nTEKNİK ALAN\n...",
        specification_filename="Tarifname_181140.docx",
    )
    assert len(out["applicants"]) == 1
    assert len(out["inventors"]) == 1
    app = out["applicants"][0]
    inv = out["inventors"][0]
    assert app["name"] == "TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ"
    assert app["identity"] == "1234567890"
    assert app["phone"] == "02164606574"
    assert app["email"] == "kubilay.aydin@turktelekom.com.tr"
    assert inv["name"] == "Gürkan Erkoç"
    assert inv["identity"] == "54796123412"
    assert inv["phone"] == "5552550437"
    assert inv["email"] == "gurkan.erkoc@turktelekom.com.tr"
    assert out["filing_options"]["inventor_hidden"]["status"] == "Hayır"
    assert out["filing_options"]["inventor_hidden"]["explicit"] is True
    assert out["filing_options"]["public_project"]["status"] == "Hayır"
    assert out["filing_options"]["early_publication"]["status"] == "Evet"
    assert out["filing_options"]["early_publication"]["explicit"] is True
    # 10 haneli VKN artık diğer bilgilerde telefon diye tekrar üretilmez.
    assert not any(x["label"] == "Telefon" and re.sub(r"\D", "", x["value"]) == "1234567890" for x in out["other_information"])


def test_v5456_cpu_ner_does_not_create_ghost_role_without_explicit_role_context():
    rule = {
        "applicants": [],
        "inventors": [{
            "identity": "54796123412", "name": "", "country": "Türkiye", "city": "İstanbul",
            "district": "Ümraniye", "address": "Adres", "email": "gurkan@example.com",
            "phone": "5552550437", "birth_date": "23.05.1992", "source": "beyan.docx",
        }],
        "filing_options": {}, "conflicts": [], "field_sources": {},
    }
    ner = {"entities": [
        {"source": "beyan.docx", "label": "ORG", "word": "Teknoloji Merkezi", "score": .99,
         "before": "Adres bilgileri ", "after": " e-posta"},
        {"source": "beyan.docx", "label": "PER", "word": "Gürkan Erkoç", "score": .99,
         "before": "BULUŞ SAHİBİ\nAd Soyad: ", "after": "\nTCKN: 54796123412"},
    ]}
    out = merge_verified_cpu_ner_application_information(rule, ner, [("beyan.docx", "BULUŞ SAHİBİ\nAd Soyad: Gürkan Erkoç\nTCKN: 54796123412\nTeknoloji Merkezi")])
    assert out["applicants"] == []
    assert out["inventors"][0]["name"] == "Gürkan Erkoç"


def test_v5456_header_page_field_removed_only_on_conversion_copy():
    src = open("Tarifname_181176_template.docx", "rb").read()
    split = split_patent_docx(src)
    conversion = remove_word_header_page_numbers(split["Tarifname.docx"])
    with zipfile.ZipFile(io.BytesIO(split["Tarifname.docx"])) as z_before, zipfile.ZipFile(io.BytesIO(conversion)) as z_after:
        before = b"\n".join(z_before.read(n) for n in z_before.namelist() if n.startswith("word/header") and n.endswith(".xml"))
        after = b"\n".join(z_after.read(n) for n in z_after.namelist() if n.startswith("word/header") and n.endswith(".xml"))
        assert b"PAGE" in before
        assert b"PAGE" not in after
        # Body XML aynıdır; yalnız dönüşüm kopyasının header'ı temizlenir.
        assert z_before.read("word/document.xml") == z_after.read("word/document.xml")


def test_v5456_pdf_cleanup_removes_top_left_1x_artifact_without_body_text():
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((20, 25), "1X")
    page.insert_text((72, 120), "TEKNİK ALAN")
    raw = doc.tobytes()
    doc.close()
    cleaned = _pdf_epats_cleanup(raw)
    d = fitz.open(stream=cleaned, filetype="pdf")
    text = "\n".join(p.get_text() for p in d)
    d.close()
    assert "1X" not in text
    assert "ALAN" in text
