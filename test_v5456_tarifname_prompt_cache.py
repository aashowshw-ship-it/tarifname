from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from ai_cache import tarifname_prompt_cache_key, tarifname_prompt_cache_kwargs
from ai_metrics import estimate_token_cost_usd, metric_from_response, usage_counts
from rules import APP_VERSION, RULESET_VERSION, TARIFNAME_RULES

ROOT = Path(__file__).resolve().parent


def _response(*, input_tokens=100_000, cached=20_000, cache_write=10_000, output=10_000, reasoning=4_000):
    return SimpleNamespace(
        output_text='{"ok": true}',
        usage=SimpleNamespace(
            input_tokens=input_tokens,
            output_tokens=output,
            input_tokens_details=SimpleNamespace(cached_tokens=cached, cache_write_tokens=cache_write),
            output_tokens_details=SimpleNamespace(reasoning_tokens=reasoning),
        ),
    )


def test_release_version_and_manifest_include_cache_helper():
    assert APP_VERSION == "v5.4.59"
    assert RULESET_VERSION == "2026-09-11.v52"
    assert (ROOT / "README.md").read_text(encoding="utf-8").startswith("# Patent Atölyesi v5.4.59")
    manifest = (ROOT / "REPO_FILE_MANIFEST.txt").read_text(encoding="utf-8").splitlines()
    assert "ai_cache.py" in manifest
    assert "test_v5456_tarifname_prompt_cache.py" in manifest


def test_tarifname_cache_key_is_stable_non_sensitive_and_rule_bound():
    key1 = tarifname_prompt_cache_key()
    key2 = tarifname_prompt_cache_key()
    assert key1 == key2
    assert key1.startswith("pa-tarifname-rules-")
    assert len(key1) < 64
    assert "BBF" not in key1 and "müşteri" not in key1.lower()


def test_cache_controls_apply_only_to_gpt56_tarifname_rule_prefix():
    prompt = TARIFNAME_RULES + "\nDİNAMİK BBF TESTİ"
    kwargs = tarifname_prompt_cache_kwargs(prompt, "gpt-5.6")
    assert kwargs["prompt_cache_key"] == tarifname_prompt_cache_key()
    assert kwargs["prompt_cache_options"] == {"ttl": "30m"}
    assert tarifname_prompt_cache_kwargs("Görüş promptu", "gpt-5.6") == {}
    assert tarifname_prompt_cache_kwargs(prompt, "other-model") == {}


def test_cache_write_tokens_are_metered_without_double_counting():
    counts = usage_counts(_response())
    assert counts == {
        "input_tokens": 100_000,
        "cached_input_tokens": 20_000,
        "cache_write_tokens": 10_000,
        "output_tokens": 10_000,
        "reasoning_tokens": 4_000,
    }
    metric = metric_from_response(_response(), stage="test", model="gpt-5.6", duration_seconds=1.0)
    assert metric.estimated_token_cost_usd == pytest.approx(0.538, rel=1e-9)
    assert metric.cache_write_tokens == 10_000


def test_cache_write_is_clamped_and_long_context_rate_applies():
    counts = usage_counts(_response(input_tokens=100, cached=80, cache_write=50, output=0, reasoning=0))
    assert counts["cache_write_tokens"] == 20
    cost = estimate_token_cost_usd(
        "gpt-5.6",
        input_tokens=300_000,
        cached_input_tokens=100_000,
        cache_write_tokens=50_000,
        output_tokens=20_000,
    )
    assert cost == pytest.approx(2.38, rel=1e-9)


def test_current_ui_injects_cache_controls_without_prompt_rewrite():
    src = (ROOT / "app.py").read_text(encoding="utf-8")
    assert "kwargs.update(workflow_prompt_cache_kwargs(prompt, MODEL))" in src
    assert '"input": [{"role": "user", "content": content}]' in src
    assert '"Cache write"' in src
