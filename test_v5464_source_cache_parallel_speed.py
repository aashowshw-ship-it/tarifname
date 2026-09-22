from pathlib import Path

from rules import APP_VERSION, RULESET_VERSION
from source_cache import source_artifact_key, bounded_cache_get, bounded_cache_set, ordered_parallel_results
import app_core

ROOT = Path(__file__).resolve().parent


def test_version_and_ruleset():
    assert APP_VERSION == "v5.4.75"
    assert RULESET_VERSION == "2026-09-22.v67"


def test_content_addressed_source_key_uses_exact_bytes():
    a = source_artifact_key(name="same.pdf", data=b"abc", mime="application/pdf", kind="text", extractor_version="v1")
    b = source_artifact_key(name="same.pdf", data=b"abd", mime="application/pdf", kind="text", extractor_version="v1")
    c = source_artifact_key(name="same.pdf", data=b"abc", mime="application/pdf", kind="text", extractor_version="v1")
    assert a != b
    assert a == c


def test_bounded_cache_is_exact_and_lru_like():
    cache = {}
    bounded_cache_set(cache, "a", {"x": [1, 2]}, max_entries=2)
    bounded_cache_set(cache, "b", "B", max_entries=2)
    assert bounded_cache_get(cache, "a") == {"x": [1, 2]}
    bounded_cache_set(cache, "c", "C", max_entries=2)
    assert "b" not in cache
    assert cache["a"] == {"x": [1, 2]}
    assert cache["c"] == "C"


def test_parallel_helper_preserves_input_order():
    vals = [3, 1, 2, 0]
    assert ordered_parallel_results(vals, lambda x: x * 10, max_workers=4) == [30, 10, 20, 0]


def test_app_core_text_cache_reuses_byte_identical_text_without_summarising():
    app_core._SOURCE_ARTIFACT_CACHE_FALLBACK.clear()
    a = app_core.UploadedAsset("a.txt", b"satir 1\nsatir 2", "text/plain")
    first = app_core.extract_text_from_asset(a)
    second = app_core.extract_text_from_asset(a)
    assert first == "satir 1\nsatir 2"
    assert second == first
    assert len(app_core._SOURCE_ARTIFACT_CACHE_FALLBACK) == 1


def test_parallel_combine_preserves_asset_order_and_content():
    app_core._SOURCE_ARTIFACT_CACHE_FALLBACK.clear()
    assets = [
        app_core.UploadedAsset("one.txt", b"ONE", "text/plain"),
        app_core.UploadedAsset("two.txt", b"TWO", "text/plain"),
        app_core.UploadedAsset("three.txt", b"THREE", "text/plain"),
    ]
    text, images = app_core.combine_asset_text("KAYNAK", assets)
    assert images == []
    assert text.index("one.txt") < text.index("two.txt") < text.index("three.txt")
    assert text.index("ONE") < text.index("TWO") < text.index("THREE")


def test_ui_uses_parallel_local_source_preprocessing_but_keeps_visual_fallback_serial():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert "extract_asset_texts_parallel" in src
    assert "Image-only PDF AI fallback remains serial" in src
    assert "_visual_pdf_text(asset, metric_context=metric_context)" in src


def test_opinion_final_word_gate_and_examiner_can_overlap_without_removing_gates():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert "ThreadPoolExecutor(max_workers=2" in src
    assert "build_and_gate_gorus_opinion," in src
    assert "gorus_examiner_persuasion_prompt(" in src
    assert "validate_opinion_against_raw_sources(" in src
    assert "validate_ai_quality_audit(" in src
    assert "validate_examiner_persuasion_assessment(" in src


def test_cache_is_session_scoped_in_ui_and_not_global_cross_user_state():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert 'state = getattr(st, "session_state")' in src
    assert '_SOURCE_ARTIFACT_CACHE_KEY = "_patent_atolyesi_source_artifacts_v1"' in src
