from processes import (
    extract_application_information_rule_based,
    merge_verified_ai_application_information,
)
from rules import APP_VERSION


def test_browser_ai_merge_repairs_roles_only_with_source_evidence():
    assert APP_VERSION == "v5.4.55"
    blocks = [("beyan.docx", """
Hak sahibi / başvuru sahibi
Unvanı: TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ
VKN: 1234567890
Adres: Gayrettepe Mah. İstanbul Türkiye
Buluş sahibi
Ad Soyad: Gürkan Erkoç
TCKN: 54796123412
E-posta: gurkan.erkoc@turktelekom.com.tr
Telefon: 5552550437
""")]
    rule = extract_application_information_rule_based(blocks, specification_text="Hologramlı görüşme yöntemi", specification_filename="Tarifname_181140.docx")
    ai = {
        "applicants": [{"name": "TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ", "identity": "1234567890", "source": "beyan.docx"}],
        "inventors": [{"name": "Gürkan Erkoç", "identity": "54796123412", "email": "gurkan.erkoc@turktelekom.com.tr", "phone": "5552550437", "source": "beyan.docx"}],
        "filing_options": {},
    }
    merged = merge_verified_ai_application_information(rule, ai, blocks)
    assert merged["applicants"][0]["name"] == "TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ"
    assert merged["inventors"][0]["name"] == "Gürkan Erkoç"
    assert merged["inventors"][0]["email"] == "gurkan.erkoc@turktelekom.com.tr"


def test_browser_ai_cannot_invent_name_not_in_source():
    blocks = [("mail.txt", "Buluş sahibi TCKN: 54796123412")]
    rule = extract_application_information_rule_based(blocks)
    ai = {"applicants": [], "inventors": [{"name": "Uydurma Kişi", "identity": "54796123412", "source": "mail.txt"}], "filing_options": {}}
    merged = merge_verified_ai_application_information(rule, ai, blocks)
    assert not any(x.get("name") == "Uydurma Kişi" for x in merged.get("inventors") or [])
