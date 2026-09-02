from email.message import EmailMessage

from processes import _current_email_body, _eml_text, _local_ai_excerpt
from rules import APP_VERSION


def test_version():
    assert APP_VERSION == "v5.4.58"


def test_current_body_drops_old_outlook_thread_and_signature():
    raw = """Merhaba,\n\n1. Buluşçu bilgileri gizlensin mi? HAYIR\n2. Kamu destekli proje kapsamında mı? HAYIR\n3. Erken yayın talep ediliyor mu? EVET\n\nSaygılarımla\nGürkan Erkoç\ngurkan.erkoc@example.com\n\nFrom: Eski Kişi <old@example.com>\nSent: Monday\nTo: Someone\nSubject: Eski konu\nHak Sahibi: YANLIŞ ŞİRKET A.Ş.\n"""
    body = _current_email_body(raw)
    assert "Buluşçu bilgileri" in body
    assert "Erken yayın" in body
    assert "Gürkan Erkoç" not in body
    assert "YANLIŞ ŞİRKET" not in body
    assert "old@example.com" not in body


def test_eml_uses_only_current_body_not_headers_or_attachment_or_old_reply():
    msg = EmailMessage()
    msg["Subject"] = "Eski/önemsiz konu"
    msg["From"] = "sender@example.com"
    msg["To"] = "receiver@example.com"
    msg.set_content("""Hak Sahibi: TT MOBİL İLETİŞİM HİZMETLERİ ANONİM ŞİRKETİ\nErken yayın talep ediliyor mu? EVET\n\n-----Original Message-----\nFrom: old@example.com\nHak Sahibi: YANLIŞ ŞİRKET A.Ş.\n""")
    msg.add_attachment(b"Hak Sahibi: EKTEKI YANLIS SIRKET", maintype="text", subtype="plain", filename="eski.txt")
    text = _eml_text(msg.as_bytes())
    assert "TT MOBİL" in text
    assert "Erken yayın" in text
    assert "YANLIŞ ŞİRKET" not in text
    assert "EKTEKI" not in text
    assert "sender@example.com" not in text
    assert "Eski/önemsiz konu" not in text


def test_local_ai_excerpt_is_short():
    source = [("mail.eml", "Hak Sahibi: ABC A.Ş.\n" + ("gereksiz satır\n" * 500) + "Erken yayın talep ediliyor mu? EVET")]
    excerpt = _local_ai_excerpt(source)
    assert len(excerpt) <= 3200
    assert "Hak Sahibi" in excerpt
