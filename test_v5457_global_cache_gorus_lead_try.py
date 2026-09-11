from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

import gorus_audit
from ai_cache import workflow_prompt_cache_key, workflow_prompt_cache_kwargs
from ai_metrics import metric_from_response, usd_to_try, usd_try_rate
from rules import (
    APP_VERSION,
    RULESET_VERSION,
    ARASTIRMA_GUNCELLEME_RULES,
    ARASTIRMA_RULES,
    GORUS_RULES,
    TARIFNAME_DUZENLEME_RULES,
    TARIFNAME_RULES,
)

ROOT = Path(__file__).resolve().parent


def _response(*, input_tokens=10_000, cached=2_000, cache_write=1_000, output=1_000, reasoning=300):
    return SimpleNamespace(
        output_text='{"ok": true}',
        usage=SimpleNamespace(
            input_tokens=input_tokens,
            output_tokens=output,
            input_tokens_details=SimpleNamespace(cached_tokens=cached, cache_write_tokens=cache_write),
            output_tokens_details=SimpleNamespace(reasoning_tokens=reasoning),
        ),
    )


def test_release_version_and_manifest():
    assert APP_VERSION == "v5.4.59"
    assert RULESET_VERSION == "2026-09-11.v52"
    assert (ROOT / "README.md").read_text(encoding="utf-8").startswith("# Patent Atölyesi v5.4.59")
    manifest = (ROOT / "REPO_FILE_MANIFEST.txt").read_text(encoding="utf-8").splitlines()
    assert "test_v5457_global_cache_gorus_lead_try.py" in manifest


@pytest.mark.parametrize(
    "rules_text,expected_prefix",
    [
        (TARIFNAME_RULES, "pa-tarifname-rules-"),
        (TARIFNAME_DUZENLEME_RULES, "pa-tarifname-duzenleme-rules-"),
        (GORUS_RULES, "pa-gorus-rules-"),
        (ARASTIRMA_RULES, "pa-arastirma-rules-"),
        (ARASTIRMA_GUNCELLEME_RULES, "pa-arastirma-guncelleme-rules-"),
    ],
)
def test_global_prompt_cache_applies_to_all_core_rule_prefixes(rules_text: str, expected_prefix: str):
    prompt_a = rules_text + "\nKULLANICI/DİNAMİK A"
    prompt_b = rules_text + "\nKULLANICI/DİNAMİK B"
    key_a = workflow_prompt_cache_key(prompt_a)
    key_b = workflow_prompt_cache_key(prompt_b)
    assert key_a == key_b
    assert key_a is not None and key_a.startswith(expected_prefix)
    kwargs = workflow_prompt_cache_kwargs(prompt_a, "gpt-5.6")
    assert kwargs["prompt_cache_key"] == key_a
    assert kwargs["prompt_cache_options"] == {"ttl": "30m"}


def test_global_prompt_cache_does_not_touch_unknown_or_other_model():
    assert workflow_prompt_cache_kwargs("tek seferlik görsel doğrulama", "gpt-5.6") == {}
    assert workflow_prompt_cache_kwargs(GORUS_RULES + "\nX", "other-model") == {}


def test_try_cost_metric_uses_configurable_rate(monkeypatch):
    monkeypatch.setenv("USD_TRY_RATE", "50.25")
    assert usd_try_rate() == pytest.approx(50.25)
    assert usd_to_try(2.0) == pytest.approx(100.5)
    metric = metric_from_response(_response(), stage="Görüş taslağı", model="gpt-5.6", duration_seconds=3.2)
    assert metric.estimated_token_cost_usd is not None
    assert metric.estimated_token_cost_try == pytest.approx(metric.estimated_token_cost_usd * 50.25)
    assert metric.usd_try_rate == pytest.approx(50.25)


def test_gorus_rules_forbid_double_tarifname_source_wording():
    assert "Tarifnamedeki dayanak şöyledir: Tarifname sayfa" in GORUS_RULES
    assert "çift `Tarifname` tekrarı yasaktır" in GORUS_RULES


def test_contextual_quote_lead_omits_second_tarifname(monkeypatch):
    opinion = {
        "amendment_assessment": {},
        "sections": [
            {
                "blocks": [
                    {"type": "paragraph", "text": "Tarifnamedeki dayanak şöyledir:"},
                    {"type": "quote", "text": "Birebir teknik pasaj", "attach_to_previous": True},
                ]
            }
        ],
    }
    monkeypatch.setattr(gorus_audit, "locate_quote_page_line_span", lambda *a, **k: (6, 13, 6, 17))
    gorus_audit.annotate_quote_locations(
        opinion, "spec.docx", b"dummy", "Türkçe", page_line_index=[{"page": 1, "line": 1, "text": "x"}]
    )
    quote = opinion["sections"][0]["blocks"][1]
    assert quote["lead"] == "sayfa 6, satır 13-17’de bu durum şu şekilde belirtilmiştir:" or quote["lead"] == "Sayfa 6, satır 13-17’de bu durum şu şekilde belirtilmiştir:"
    assert "Tarifname sayfa" not in quote["lead"]
    gorus_audit.validate_quote_locations_against_spec(
        opinion, "spec.docx", b"dummy", "Türkçe", page_line_index=[{"page": 1, "line": 1, "text": "x"}]
    )


def test_contextual_quote_lead_keeps_tarifname_when_previous_sentence_does_not_name_source(monkeypatch):
    opinion = {
        "amendment_assessment": {},
        "sections": [
            {
                "blocks": [
                    {"type": "paragraph", "text": "Bu teknik fark D1'de açıklanmamaktadır."},
                    {"type": "quote", "text": "Birebir teknik pasaj", "attach_to_previous": True},
                ]
            }
        ],
    }
    monkeypatch.setattr(gorus_audit, "locate_quote_page_line_span", lambda *a, **k: (6, 13, 6, 17))
    gorus_audit.annotate_quote_locations(
        opinion, "spec.docx", b"dummy", "Türkçe", page_line_index=[{"page": 1, "line": 1, "text": "x"}]
    )
    quote = opinion["sections"][0]["blocks"][1]
    assert quote["lead"].startswith("Tarifname sayfa 6")
