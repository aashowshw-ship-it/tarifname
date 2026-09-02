from processes import (
    application_precheck_missing,
    extract_application_information_rule_based,
    _html_to_text,
)
from rules import APP_VERSION


def _source_text():
    return """
HAK SAHİBİ / BAŞVURU SAHİBİ BİLGİLERİ
Unvanı\tTT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ
Hak Sahibi Adresi\tGAYRETTEPE MAH. YILDIZ POSTA CAD. TÜRK TELEKOM GENEL MÜDÜRLÜĞÜ 40 TÜRKİYE
Sahip Türü\tTüzel kişi\tUyruk\tTürkiye\tTC Kimlik / Vergi No\t1234567890
E-Posta\tgurkan.erkoc@turktelekom.com.tr\tTelefon\t5552550437

BULUŞ SAHİBİ BİLGİLERİ
TCKN / Kimlik\t12345678901\tDoğum Tarihi\t23.05.1992
Ad Soyad\tGürkan Erkoç\tCinsiyet\tErkek
İl/İlçe\tİstanbul/Ümraniye\tAdres\tFatih Sultan Mehmet Mah, Balkan Cd. No:49, 34771 Ümraniye/İstanbul TÜRKİYE
E-Posta\tgurkan.erkoc@turktelekom.com.tr\tEv/İş Telefonu\t5552550437
""".strip()


def _mail_questions():
    return """
1. Başvuru esnasında buluşçu bilgileri gizlensin mi? (Cevap verilmemesi halinde HAYIR olarak işleme alınacaktır.) à HAYIR

2. Buluş, TÜBİTAK, KOSGEB vb. bir kamu kurum tarafından desteklenen bir proje kapsamında mı ortaya çıktı? Aldıysanız lütfen kurumu ve proje numarasını belirtiniz. (Cevap verilmemesi halinde destek alınmadığı varsayılarak başvuru yapılacaktır.) à HAYIR

3. Erken yayın talep ediliyor mu? (Cevap verilmemesi halinde HAYIR olarak işleme alınacaktır.) à EVET
""".strip()


def test_v5450_person_fields_and_mail_options_are_bound_correctly():
    assert APP_VERSION == "v5.4.58"
    data = extract_application_information_rule_based(
        [("beyan.png", _source_text()), ("mail.msg", _mail_questions() + "\nBaşvuru türü: Patent\nRüçhan: Yok")],
        specification_text="TARİFNAME\nHologramlı görüşme yöntemi\nTEKNİK ALAN\n...",
        specification_filename="Tarifname_181140.docx",
    )
    assert data["reference"] == "181140"
    assert data["invention_title"] == "Hologramlı görüşme yöntemi"

    assert len(data["applicants"]) == 1
    a = data["applicants"][0]
    assert a["name"] == "TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ"
    assert a["identity"] == "1234567890"
    assert a["entity_type"] == "Tüzel kişi"
    assert a["country"] == "Türkiye"
    assert "GAYRETTEPE" in a["address"]
    assert a["email"] == "gurkan.erkoc@turktelekom.com.tr"

    assert len(data["inventors"]) == 1
    i = data["inventors"][0]
    assert i["identity"] == "12345678901"
    assert i["name"] == "Gürkan Erkoç"
    assert i["birth_date"] == "23.05.1992"
    assert i["city"] == "İstanbul"
    assert i["district"] == "Ümraniye"
    assert i["country"] == "Türkiye"
    assert "Fatih Sultan Mehmet" in i["address"]
    assert i["phone"] == "5552550437"
    assert i["identity"] != "Cinsiyet"

    opts = data["filing_options"]
    assert opts["inventor_hidden"]["status"] == "Hayır"
    assert opts["inventor_hidden"]["source"] == "mail.msg"
    assert opts["public_project"]["status"] == "Hayır"
    assert opts["early_publication"]["status"] == "Evet"


def test_v5450_html_table_boundaries_are_preserved_for_mail():
    html = "<table><tr><td>Ad Soyad</td><td>Gürkan Erkoç</td><td>Telefon</td><td>5552550437</td></tr></table>"
    text = _html_to_text(html)
    assert "\t" in text
    assert "Ad Soyad" in text and "Gürkan Erkoç" in text


def test_v5450_public_project_yes_requires_institution_and_project_number():
    data = extract_application_information_rule_based(
        [("mail.txt", "Buluş, TÜBİTAK, KOSGEB vb. bir kamu kurum tarafından desteklenen bir proje kapsamında mı ortaya çıktı? à EVET\nBaşvuru türü: Patent\nRüçhan: Yok")],
        specification_text="TARİFNAME\nTest başlığı\nTEKNİK ALAN",
        specification_filename="Tarifname_181140.docx",
    )
    metrics = {"specification_pages": 1, "claim_count": 1, "abstract_present": True, "figures_pages": 0}
    # Kişi eksikleri de gelir; burada proje alanlarının özellikle kilit oluşturduğunu doğrula.
    missing = application_precheck_missing(data, metrics)
    assert "kamu destekli proje kurumu" in missing
    assert "kamu destekli proje numarası" in missing


def test_v5450_unanswered_filing_options_use_documented_defaults():
    data = extract_application_information_rule_based(
        [("mail.txt", "Başvuru türü: Patent\nRüçhan: Yok")],
        specification_text="TARİFNAME\nTest başlığı\nTEKNİK ALAN",
        specification_filename="Tarifname_181140.docx",
    )
    assert data["filing_options"]["inventor_hidden"]["status"] == "Hayır"
    assert data["filing_options"]["public_project"]["status"] == "Hayır"
    assert data["filing_options"]["early_publication"]["status"] == "Hayır"
    assert data["filing_options"]["early_publication"]["explicit"] is False
