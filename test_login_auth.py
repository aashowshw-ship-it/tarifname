from pathlib import Path

from auth import authenticate, load_users, make_password_hash, verify_password

ROOT = Path(__file__).resolve().parent


def test_password_hash_roundtrip():
    hashed = make_password_hash("Gizli-123")
    assert hashed.startswith("pbkdf2_sha256$")
    assert verify_password("Gizli-123", hashed)
    assert not verify_password("yanlis", hashed)


def test_multiple_users_plain_and_hashed():
    hashed = make_password_hash("ikinci-sifre")
    raw = (
        '{'
        '"samet":{"password":"ilk-sifre","display_name":"Samet","role":"admin"},'
        f'"musteri1":{{"password_hash":"{hashed}","display_name":"Müşteri 1","role":"user"}}'
        '}'
    )
    users = load_users(raw)
    assert authenticate(users, "samet", "ilk-sifre").role == "admin"
    assert authenticate(users, "musteri1", "ikinci-sifre").display_name == "Müşteri 1"
    assert authenticate(users, "musteri1", "yanlis") is None
    assert authenticate(users, "yok", "x") is None


def test_inactive_user_cannot_login():
    users = load_users('{"kapali":{"password":"123","active":false}}')
    assert authenticate(users, "kapali", "123") is None


def test_login_gate_is_before_main_work_type():
    text = (ROOT / "app.py").read_text(encoding="utf-8")
    gate = text.index('AUTH_SESSION_KEY = "pa_authenticated_user"')
    stop = text.index("st.stop()", gate)
    work_type = text.index("work_type = st.radio(")
    assert gate < stop < work_type
    assert 'PATENT_USERS_JSON' in text
    assert 'Çıkış yap' in text
