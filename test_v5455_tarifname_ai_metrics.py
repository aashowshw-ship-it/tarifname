from __future__ import annotations

from types import SimpleNamespace
from pathlib import Path

import pytest

from ai_metrics import estimate_token_cost_usd, metric_from_response, usage_counts
from rules import APP_VERSION, RULESET_VERSION

ROOT = Path(__file__).resolve().parent


def _fake_response(*, input_tokens=100_000, cached_tokens=20_000, output_tokens=10_000, reasoning_tokens=4_000):
    return SimpleNamespace(
        output_text='{"ok": true}',
        usage=SimpleNamespace(
            input_tokens=input_tokens,
            output_tokens=output_tokens,
            input_tokens_details=SimpleNamespace(cached_tokens=cached_tokens, cache_write_tokens=0),
            output_tokens_details=SimpleNamespace(reasoning_tokens=reasoning_tokens),
        ),
    )


def test_release_version_and_manifest_include_metrics_helper():
    assert APP_VERSION == "v5.4.59"
    assert RULESET_VERSION == "2026-09-11.v52"
    assert (ROOT / "README.md").read_text(encoding="utf-8").startswith("# Patent Atölyesi v5.4.59")
    assert "ai_metrics.py" in (ROOT / "REPO_FILE_MANIFEST.txt").read_text(encoding="utf-8").splitlines()


def test_usage_reads_cached_and_reasoning_without_double_counting():
    counts = usage_counts(_fake_response())
    assert counts == {
        "input_tokens": 100_000,
        "cached_input_tokens": 20_000,
        "cache_write_tokens": 0,
        "output_tokens": 10_000,
        "reasoning_tokens": 4_000,
    }
    metric = metric_from_response(
        _fake_response(), stage="test", model="gpt-5.6", duration_seconds=12.5
    )
    # 80k uncached input * $4/M + 20k cached * $0.40/M + 10k output * $20/M = $0.528
    assert metric.estimated_token_cost_usd == pytest.approx(0.528, rel=1e-9)
    assert metric.reasoning_tokens == 4_000


def test_long_context_price_multiplier_is_applied_to_full_request():
    cost = estimate_token_cost_usd(
        "gpt-5.6",
        input_tokens=300_000,
        cached_input_tokens=100_000,
        output_tokens=20_000,
    )
    # >272k: input/cached rates x2 and output rate x1.5.
    assert cost == pytest.approx(2.28, rel=1e-9)


def test_app_ask_json_is_instrumented_without_changing_model_or_input_contract():
    source = (ROOT / "app.py").read_text(encoding="utf-8")
    assert "started = time.perf_counter()" in source
    assert "duration = time.perf_counter() - started" in source
    assert "metric = metric_from_response(" in source
    assert "_record_ai_metric(workflow, signature, metric.as_dict())" in source
    assert '"model": MODEL' in source
    assert '"input": [{"role": "user", "content": content}]' in source


def test_tarifname_core_calls_are_labeled_for_measurement():
    source = (ROOT / "app.py").read_text(encoding="utf-8")
    for label in (
        'metric_stage="1. BBF teknik envanteri"',
        'metric_stage="2. Envanter ikinci kontrolü"',
        'metric_stage="3. Patent literatürü"',
        'metric_stage="4. Tam tarifname taslağı"',
        'metric_stage="5. Tarifname kalite kontrolü"',
        'metric_stage=f"6. Final ham-kaynak ikinci okuma (tur {repair_round + 1})"',
        'metric_stage=f"7. Kalite düzeltmesi (tur {repair_round + 1})"',
        '_show_tarifname_ai_metrics(tariff_signature)',
    ):
        assert label in source
