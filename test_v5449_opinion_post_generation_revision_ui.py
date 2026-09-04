from pathlib import Path

ROOT = Path(__file__).parent


def test_version_bumped():
    rules = (ROOT / "rules.py").read_text(encoding="utf-8")
    assert 'APP_VERSION = "v5.4.52"' in rules
    assert 'RULESET_VERSION = "2026-09-04.v42"' in rules


def test_chat_revision_ui_and_direct_editor_exist():
    app = (ROOT / "app.py").read_text(encoding="utf-8")
    assert "Görüşü revize et" in app
    assert "Revizyon talebiniz" in app
    assert "Talebi uygula ve Word'ü yeniden oluştur" in app
    assert "Metni doğrudan düzenle (isteğe bağlı)" in app
    assert "Elle düzenlenen metni uygula ve Word'ü yeniden oluştur" in app


def test_post_generation_revision_reuses_all_gates():
    app = (ROOT / "app.py").read_text(encoding="utf-8")
    assert "gorus_user_revision_prompt" in app
    assert "validate_opinion_against_raw_sources" in app
    assert "gorus_quality_audit_prompt" in app
    assert "build_and_gate_gorus_opinion" in app
    assert "gorus_examiner_persuasion_prompt" in app


def test_quotes_locked_in_direct_editor():
    app = (ROOT / "app.py").read_text(encoding="utf-8")
    assert "kaynak alıntısı {block_index+1} (kilitli)" in app
    assert "disabled=True" in app


def test_rules_document_minimal_targeted_revision():
    rules = (ROOT / "rules.py").read_text(encoding="utf-8")
    assert "mümkün olan en küçük kapsamda" in (ROOT / "app.py").read_text(encoding="utf-8")
    assert "Görüş revizyon asistanı" in rules
    assert "kullanıcının açık talebi dışında" in rules.lower()
