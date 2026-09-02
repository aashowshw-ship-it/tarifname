from processes import extract_application_information_hybrid


def test_hybrid_local_ai_repairs_misaligned_person_fields_and_keeps_spec_title():
    source = """
HAK SAHİBİ / BAŞVURU SAHİBİ
Unvanı / Ad Soyad: TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ
VKN: 1234567890
Adres: Gayrettepe Mah. Yıldız Posta Cad. Türk Telekom Genel Müdürlüğü 40 Türkiye
E-posta: kubilay.aydin@turktelekom.com.tr
Telefon: 02164606574

BULUŞ SAHİBİ
TCKN: 54796123412
Ad Soyad: Gürkan Erkoç
İl / İlçe: İstanbul / Ümraniye
Adres: Fatih Sultan Mehmet Mah, Balkan Cd. No:49, 34771 Ümraniye/İstanbul
E-posta: İmza gurkan.erkoc@turktelekom.com.tr İmza
Telefon: 5552550437
Doğum Tarihi: 23.05.1992

Buluşçu bilgileri gizlensin mi? HAYIR
Kamu kurum tarafından desteklenen proje kapsamında mı? HAYIR
Erken yayın talep ediliyor mu? EVET
"""

    def fake_runner(prompt, schema):
        assert "TT MOBİL" in prompt
        return {
            "applicants": [{
                "entity_type": "Tüzel kişi",
                "identity": "1234567890",
                "name": "TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ",
                "country": "Türkiye",
                "city": "",
                "district": "",
                "address": "Gayrettepe Mah. Yıldız Posta Cad. Türk Telekom Genel Müdürlüğü 40 Türkiye",
                "email": "kubilay.aydin@turktelekom.com.tr",
                "phone": "02164606574",
                "birth_date": "",
                "source": "beyan.docx",
            }],
            "inventors": [{
                "identity": "54796123412",
                "name": "Gürkan Erkoç",
                "country": "",
                "city": "İstanbul",
                "district": "Ümraniye",
                "address": "Fatih Sultan Mehmet Mah, Balkan Cd. No:49, 34771 Ümraniye/İstanbul",
                "email": "gurkan.erkoc@turktelekom.com.tr",
                "phone": "5552550437",
                "birth_date": "23.05.1992",
                "source": "beyan.docx",
            }],
            "filing_options": {
                "inventor_hidden": {"status": "Hayır", "source": "beyan.docx"},
                "public_project": {"status": "Hayır", "source": "beyan.docx", "institution": "", "project_number": ""},
                "early_publication": {"status": "Evet", "source": "beyan.docx"},
            },
        }, {"used": True, "available": True, "model": "test.gguf", "warning": ""}

    result = extract_application_information_hybrid(
        [("beyan.docx", source)],
        specification_text="TARİFNAME\nHologramlı görüşme yöntemi\nTEKNİK ALAN\n...",
        specification_filename="Tarifname_181140.docx",
        local_ai_runner=fake_runner,
    )

    assert result["reference"] == "181140"
    assert result["invention_title"] == "Hologramlı görüşme yöntemi"
    applicant = result["applicants"][0]
    assert applicant["name"] == "TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ"
    assert applicant["identity"] == "1234567890"
    assert applicant["email"] == "kubilay.aydin@turktelekom.com.tr"
    inventor = result["inventors"][0]
    assert inventor["name"] == "Gürkan Erkoç"
    assert inventor["identity"] == "54796123412"
    assert inventor["email"] == "gurkan.erkoc@turktelekom.com.tr"
    assert inventor["birth_date"] == "23.05.1992"
    assert result["filing_options"]["inventor_hidden"]["status"] == "Hayır"
    assert result["filing_options"]["public_project"]["status"] == "Hayır"
    assert result["filing_options"]["early_publication"]["status"] == "Evet"
    assert result["local_ai"]["used"] is True


def test_hybrid_falls_back_without_local_ai_and_sanitizes_email():
    text = """
BULUŞ SAHİBİ BİLGİLERİ
TCKN: 54796123412
Ad Soyad: Gürkan Erkoç
E-posta: İmza gurkan.erkoc@turktelekom.com.tr İmza
Adres: İstanbul Türkiye
Rüçhan: Yok
Başvuru Türü: Patent
"""

    def unavailable(prompt, schema):
        return {}, {"used": False, "available": False, "warning": "model yok"}

    result = extract_application_information_hybrid(
        [("mail.txt", text)],
        specification_text="TARİFNAME\nBaşlık\nTEKNİK ALAN\n...",
        specification_filename="Tarifname_123456.docx",
        local_ai_runner=unavailable,
    )
    assert result["inventors"][0]["email"] == "gurkan.erkoc@turktelekom.com.tr"
    assert result["local_ai"]["used"] is False
