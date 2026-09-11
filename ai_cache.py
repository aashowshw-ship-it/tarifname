from __future__ import annotations

import hashlib
from typing import Any

from rules import (
    ARASTIRMA_GUNCELLEME_RULES,
    ARASTIRMA_RULES,
    GORUS_RULES,
    TARIFNAME_DUZENLEME_RULES,
    TARIFNAME_RULES,
)


# Binding, reusable rule prefixes for the five core Patent Atölyesi workflows.
# Cache keys are derived only from static rule text; customer/source content never
# enters the key. This keeps caching quality-neutral and non-sensitive.
_RULE_PREFIXES: tuple[tuple[str, str], ...] = (
    ("tarifname-duzenleme", TARIFNAME_DUZENLEME_RULES),
    ("arastirma-guncelleme", ARASTIRMA_GUNCELLEME_RULES),
    ("tarifname", TARIFNAME_RULES),
    ("gorus", GORUS_RULES),
    ("arastirma", ARASTIRMA_RULES),
)


def _supports_prompt_cache_options(model: str) -> bool:
    """Enable the explicit 30m cache controls only for GPT-5.6 family calls."""
    normalized = str(model or "").strip().lower()
    return normalized == "gpt-5.6" or normalized.startswith("gpt-5.6-")


def _cache_key(label: str, rules_text: str) -> str:
    digest = hashlib.sha256(str(rules_text or "").encode("utf-8")).hexdigest()[:24]
    return f"pa-{label}-rules-{digest}"


def workflow_prompt_cache_key(prompt: str) -> str | None:
    """Return the rule-bound cache bucket for a core workflow prompt, if any."""
    value = str(prompt or "")
    for label, rules_text in _RULE_PREFIXES:
        if value.startswith(rules_text):
            return _cache_key(label, rules_text)
    return None


def workflow_prompt_cache_kwargs(prompt: str, model: str) -> dict[str, Any]:
    """Return cache controls for every rule-bound core Patent Atölyesi AI workflow.

    The prompt itself is never rewritten. Only prompts whose first bytes are one of
    the binding workflow rule blocks are grouped under a stable, non-sensitive key.
    This covers tarifname, tarifname düzenleme, görüş/görüş revizyonu, Tip 3 araştırma
    and araştırma güncelleme. One-off AI calls without a reusable binding-rule prefix
    are intentionally left unchanged because they have no stable prefix to reuse.
    """
    if not _supports_prompt_cache_options(model):
        return {}
    key = workflow_prompt_cache_key(prompt)
    if not key:
        return {}
    return {
        "prompt_cache_key": key,
        "prompt_cache_options": {"ttl": "30m"},
    }


# Backward-compatible v5.4.56 API kept for older tests/callers.
def tarifname_prompt_cache_key() -> str:
    return _cache_key("tarifname", TARIFNAME_RULES)


def tarifname_prompt_cache_kwargs(prompt: str, model: str) -> dict[str, Any]:
    value = str(prompt or "")
    if not value.startswith(TARIFNAME_RULES):
        return {}
    return workflow_prompt_cache_kwargs(value, model)
