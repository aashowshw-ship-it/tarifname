from pathlib import Path

from processes import (
    extract_application_information_rule_based,
    merge_verified_cpu_semantic_application_information,
)
from rules import APP_VERSION

ROOT = Path(__file__).resolve().parent


def test_v5457_version_and_semantic_cpu_model():
    assert APP_VERSION == "v5.4.58"
    js = (ROOT / "browser_ai_component" / "main.js").read_text(encoding="utf-8")
    assert "zero-shot-classification" in js
    assert "multilingual-MiniLMv2-L6-mnli-xnli-ONNX" in js
    assert 'device: "wasm"' in js
    assert "token-classification" not in js
    assert "webgpu" not in js.lower()


def test_v5457_semantic_roles_rebuild_wrong_rule_rows_from_different_form_layout():
    source = """FORM A
Şirket Bilgisi
Merkez Adresi: Gayrettepe Mah. İstanbul Türkiye
Telefon: 02164606574
Unvanı: TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ
E-posta: kurumsal@turktelekom.com.tr
Vergi No: 1234567890

Teknik ekip kişisi
Telefon: 5552550437
Ad Soyad: Gürkan Erkoç
Adres: Ümraniye İstanbul Türkiye
TCKN: 54796123412
E-posta: gurkan.erkoc@turktelekom.com.tr
"""
    # Eski parser'ın üretebildiği hatalı satırları özellikle taklit et.
    rule = extract_application_information_rule_based(
        [("farkli-form.docx", source)],
        specification_text="TARİFNAME\nHologramlı görüşme yöntemi",
        specification_filename="Tarifname_181140.docx",
    )
    rule["applicants"] = [{
        "entity_type": "", "identity": "54796123412", "name": "ronik", "country": "Türkiye",
        "city": "", "district": "", "address": "Gayrettepe Mah. İstanbul Türkiye",
        "email": "gurkan.erkoc@turktelekom.com.tr", "phone": "5552550437", "birth_date": "",
        "source": "farkli-form.docx",
    }]
    rule["inventors"] = [{
        "identity": "54796123412", "name": "Sultan", "country": "Türkiye", "city": "İstanbul",
        "district": "", "address": "Ümraniye İstanbul Türkiye", "email": "gurkan.erkoc@turktelekom.com.tr",
        "phone": "5552550437", "birth_date": "", "source": "farkli-form.docx",
    }]
    semantic = {"blocks": [
        {
            "source": "farkli-form.docx", "role": "applicant", "score": 0.84, "margin": 0.25,
            "text": "Şirket Bilgisi\nMerkez Adresi: Gayrettepe Mah. İstanbul Türkiye\nTelefon: 02164606574\nUnvanı: TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ\nE-posta: kurumsal@turktelekom.com.tr\nVergi No: 1234567890",
        },
        {
            "source": "farkli-form.docx", "role": "inventor", "score": 0.87, "margin": 0.31,
            "text": "Teknik ekip kişisi\nTelefon: 5552550437\nAd Soyad: Gürkan Erkoç\nAdres: Ümraniye İstanbul Türkiye\nTCKN: 54796123412\nE-posta: gurkan.erkoc@turktelekom.com.tr",
        },
    ]}
    out = merge_verified_cpu_semantic_application_information(rule, semantic, [("farkli-form.docx", source)])
    assert len(out["applicants"]) == 1
    assert len(out["inventors"]) == 1
    app = out["applicants"][0]
    inv = out["inventors"][0]
    assert app["name"] == "TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ"
    assert app["identity"] == "1234567890"
    assert app["phone"] == "02164606574"
    assert app["email"] == "kurumsal@turktelekom.com.tr"
    assert inv["name"] == "Gürkan Erkoç"
    assert inv["identity"] == "54796123412"
    assert inv["phone"] == "5552550437"
    assert inv["email"] == "gurkan.erkoc@turktelekom.com.tr"
    assert all(x["name"] not in {"ronik", "Sultan"} for x in out["applicants"] + out["inventors"])


def test_v5457_semantic_ai_only_assigns_role_values_are_source_verified():
    source = "Unvanı: GERÇEK ŞİRKET ANONİM ŞİRKETİ\nVKN: 1234567890"
    rule = extract_application_information_rule_based([("form.txt", source)])
    semantic = {"blocks": [{
        "source": "form.txt", "role": "applicant", "score": .9, "margin": .4,
        "text": "Unvanı: UYDURMA ŞİRKET ANONİM ŞİRKETİ\nVKN: 9999999999",
    }]}
    out = merge_verified_cpu_semantic_application_information(rule, semantic, [("form.txt", source)])
    assert not any(x.get("name") == "UYDURMA ŞİRKET ANONİM ŞİRKETİ" for x in out.get("applicants") or [])
    assert not any(x.get("identity") == "9999999999" for x in out.get("applicants") or [])


def test_v5457_mail_preferences_still_read_hayir_hayir_evet_directly():
    mail = """1. Başvuru esnasında buluşçu bilgileri gizlensin mi? (Cevap verilmemesi halinde HAYIR olarak işleme alınacaktır.) à HAYIR
2. Buluş, TÜBİTAK, KOSGEB vb. bir kamu kurum tarafından desteklenen bir proje kapsamında mı ortaya çıktı? (Cevap verilmemesi halinde destek alınmadığı varsayılarak başvuru yapılacaktır.) à HAYIR
3. Erken yayın talep ediliyor mu? (Cevap verilmemesi halinde HAYIR olarak işleme alınacaktır.) à EVET"""
    out = extract_application_information_rule_based([("mail.msg", mail)])
    assert out["filing_options"]["inventor_hidden"]["status"] == "Hayır"
    assert out["filing_options"]["inventor_hidden"]["explicit"] is True
    assert out["filing_options"]["public_project"]["status"] == "Hayır"
    assert out["filing_options"]["public_project"]["explicit"] is True
    assert out["filing_options"]["early_publication"]["status"] == "Evet"
    assert out["filing_options"]["early_publication"]["explicit"] is True
