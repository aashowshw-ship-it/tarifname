from pathlib import Path

from processes import merge_verified_cpu_ner_application_information
from rules import APP_VERSION

ROOT = Path(__file__).resolve().parent


def test_v5455_version_and_browser_component_is_cpu_wasm_only():
    assert APP_VERSION == "v5.4.55"
    js = (ROOT / "browser_ai_component" / "main.js").read_text(encoding="utf-8")
    assert 'device: "wasm"' in js
    assert 'dtype: "q8"' in js
    assert "distilbert-base-multilingual-cased-ner-hrl" in js
    assert 'device: "webgpu"' not in js
    assert "Qwen2.5" not in js


def test_cpu_ner_repairs_company_inventor_and_binds_labeled_contacts():
    source = """HAK SAHİBİ / BAŞVURU SAHİBİ
Unvanı: TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ
VKN: 1234567890
Adres: Gayrettepe Mah. Yıldız Posta Cad. Türk Telekom Genel Müdürlüğü 40 İstanbul Türkiye
E-posta: kurumsal@turktelekom.com.tr
Telefon: 02124606574
Dosya içi ilgisiz numara: 181140

BULUŞ SAHİBİ
Ad Soyad: Gürkan Erkoç
TCKN: 54796123412
Doğum Tarihi: 23.05.1992
Adres: Fatih Sultan Mehmet Mah, Balkan Cd. No:49, 34771 Ümraniye/İstanbul
E-posta: gurkan.erkoc@turktelekom.com.tr
Telefon: 5552550437
"""
    rule_data = {
        "applicants": [{
            "entity_type": "", "identity": "", "name": "(SAHİPLERİ): Not: Başvuru sahibinin birden fazla olması",
            "country": "", "city": "", "district": "", "address": "", "email": "", "phone": "",
            "birth_date": "", "source": "beyan.docx",
        }],
        "inventors": [{
            "identity": "54796123412", "name": "", "country": "Türkiye", "city": "İstanbul", "district": "Ümraniye",
            "address": "Fatih Sultan Mehmet Mah, Balkan Cd. No:49, 34771 Ümraniye/İstanbul",
            "email": "İmza gurkan.erkoc@turktelekom.com.tr İmza", "phone": "181140", "birth_date": "23.05.1992",
            "source": "beyan.docx",
        }],
        "filing_options": {}, "field_sources": {}, "other_fields": {}, "conflicts": [],
    }
    ner_data = {"entities": [
        {
            "source": "beyan.docx", "label": "ORG", "word": "TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ", "score": 0.99,
            "before": "HAK SAHİBİ / BAŞVURU SAHİBİ\nUnvanı: ",
            "after": "\nVKN: 1234567890\nAdres: Gayrettepe Mah. Yıldız Posta Cad. Türk Telekom Genel Müdürlüğü 40 İstanbul Türkiye\nE-posta: kurumsal@turktelekom.com.tr\nTelefon: 02124606574\nDosya içi ilgisiz numara: 181140",
        },
        {
            "source": "beyan.docx", "label": "PER", "word": "Gürkan Erkoç", "score": 0.99,
            "before": "BULUŞ SAHİBİ\nAd Soyad: ",
            "after": "\nTCKN: 54796123412\nDoğum Tarihi: 23.05.1992\nAdres: Fatih Sultan Mehmet Mah, Balkan Cd. No:49, 34771 Ümraniye/İstanbul\nE-posta: gurkan.erkoc@turktelekom.com.tr\nTelefon: 5552550437",
        },
    ]}

    out = merge_verified_cpu_ner_application_information(rule_data, ner_data, [("beyan.docx", source)])
    applicant = out["applicants"][0]
    inventor = out["inventors"][0]

    assert applicant["name"] == "TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ"
    assert applicant["identity"] == "1234567890"
    assert applicant["phone"] == "02124606574"
    assert applicant["email"] == "kurumsal@turktelekom.com.tr"
    assert applicant["address"].startswith("Gayrettepe Mah.")
    assert applicant["country"] == "Türkiye"

    assert inventor["name"] == "Gürkan Erkoç"
    assert inventor["identity"] == "54796123412"
    assert inventor["email"] == "gurkan.erkoc@turktelekom.com.tr"
    assert inventor["phone"] == "5552550437"
    assert inventor["phone"] != "181140"
    assert inventor["birth_date"] == "23.05.1992"
