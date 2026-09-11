from __future__ import annotations

from dataclasses import dataclass, asdict
import os
from typing import Any

# GPT-5.6 Sol promotional API pricing, effective in September 2026.
# Source of truth for this release: OpenAI model pricing page.
_GPT56_SOL_INPUT_PER_M = 4.00
_GPT56_SOL_CACHED_INPUT_PER_M = 0.40
_GPT56_SOL_CACHE_WRITE_MULTIPLIER = 1.25
_GPT56_SOL_OUTPUT_PER_M = 20.00
_GPT56_SOL_LONG_CONTEXT_THRESHOLD = 272_000
_GPT56_SOL_LONG_INPUT_MULTIPLIER = 2.0
_GPT56_SOL_LONG_OUTPUT_MULTIPLIER = 1.5

# Approximate TL display only. API billing remains USD internally.
# Operators can override this without code changes via USD_TRY_RATE.
_DEFAULT_USD_TRY_RATE = 48.6031  # 2026-09-11 reference rate


@dataclass(frozen=True)
class AiUsageMetric:
    stage: str
    model: str
    duration_seconds: float
    input_tokens: int
    cached_input_tokens: int
    cache_write_tokens: int
    output_tokens: int
    reasoning_tokens: int
    estimated_token_cost_usd: float | None
    estimated_token_cost_try: float | None
    usd_try_rate: float
    web_search: bool = False
    status: str = "OK"

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


def _field(obj: Any, name: str, default: Any = 0) -> Any:
    if obj is None:
        return default
    if isinstance(obj, dict):
        return obj.get(name, default)
    return getattr(obj, name, default)


def usage_counts(response: Any) -> dict[str, int]:
    """Read Responses API usage defensively across SDK object/dict shapes."""
    usage = _field(response, "usage", None)
    input_tokens = int(_field(usage, "input_tokens", 0) or 0)
    output_tokens = int(_field(usage, "output_tokens", 0) or 0)
    input_details = _field(usage, "input_tokens_details", None)
    output_details = _field(usage, "output_tokens_details", None)
    cached_input_tokens = int(_field(input_details, "cached_tokens", 0) or 0)
    cache_write_tokens = int(_field(input_details, "cache_write_tokens", 0) or 0)
    reasoning_tokens = int(_field(output_details, "reasoning_tokens", 0) or 0)
    cached_input_tokens = max(0, min(cached_input_tokens, input_tokens))
    cache_write_tokens = max(0, min(cache_write_tokens, input_tokens - cached_input_tokens))
    reasoning_tokens = max(0, min(reasoning_tokens, output_tokens))
    return {
        "input_tokens": max(0, input_tokens),
        "cached_input_tokens": cached_input_tokens,
        "cache_write_tokens": cache_write_tokens,
        "output_tokens": max(0, output_tokens),
        "reasoning_tokens": reasoning_tokens,
    }


def estimate_token_cost_usd(
    model: str,
    *,
    input_tokens: int,
    cached_input_tokens: int,
    output_tokens: int,
    cache_write_tokens: int = 0,
) -> float | None:
    """Estimate model-token cost only; tool-call/image-generation fees are intentionally excluded."""
    normalized = str(model or "").strip().lower()
    if not (normalized == "gpt-5.6" or normalized.startswith("gpt-5.6-sol")):
        return None

    input_tokens = max(0, int(input_tokens or 0))
    cached_input_tokens = max(0, min(int(cached_input_tokens or 0), input_tokens))
    cache_write_tokens = max(0, min(int(cache_write_tokens or 0), input_tokens - cached_input_tokens))
    output_tokens = max(0, int(output_tokens or 0))
    regular_uncached_input_tokens = input_tokens - cached_input_tokens - cache_write_tokens

    input_rate = _GPT56_SOL_INPUT_PER_M
    cached_rate = _GPT56_SOL_CACHED_INPUT_PER_M
    output_rate = _GPT56_SOL_OUTPUT_PER_M
    if input_tokens > _GPT56_SOL_LONG_CONTEXT_THRESHOLD:
        input_rate *= _GPT56_SOL_LONG_INPUT_MULTIPLIER
        cached_rate *= _GPT56_SOL_LONG_INPUT_MULTIPLIER
        output_rate *= _GPT56_SOL_LONG_OUTPUT_MULTIPLIER

    cache_write_rate = input_rate * _GPT56_SOL_CACHE_WRITE_MULTIPLIER
    return (
        regular_uncached_input_tokens * input_rate
        + cache_write_tokens * cache_write_rate
        + cached_input_tokens * cached_rate
        + output_tokens * output_rate
    ) / 1_000_000.0



def usd_try_rate() -> float:
    """Return the configurable approximate USD/TRY rate used only for UI estimates."""
    raw = str(os.getenv("USD_TRY_RATE", "") or "").strip().replace(",", ".")
    if raw:
        try:
            value = float(raw)
            if value > 0:
                return value
        except ValueError:
            pass
    return _DEFAULT_USD_TRY_RATE


def usd_to_try(value: float | None, *, rate: float | None = None) -> float | None:
    if value is None:
        return None
    effective = float(rate if rate is not None else usd_try_rate())
    return float(value) * effective

def metric_from_response(
    response: Any,
    *,
    stage: str,
    model: str,
    duration_seconds: float,
    web_search: bool = False,
) -> AiUsageMetric:
    counts = usage_counts(response)
    cost_usd = estimate_token_cost_usd(
        model,
        input_tokens=counts["input_tokens"],
        cached_input_tokens=counts["cached_input_tokens"],
        output_tokens=counts["output_tokens"],
        cache_write_tokens=counts["cache_write_tokens"],
    )
    rate = usd_try_rate()
    return AiUsageMetric(
        stage=str(stage or "AI çağrısı"),
        model=str(model or ""),
        duration_seconds=max(0.0, float(duration_seconds or 0.0)),
        input_tokens=counts["input_tokens"],
        cached_input_tokens=counts["cached_input_tokens"],
        cache_write_tokens=counts["cache_write_tokens"],
        output_tokens=counts["output_tokens"],
        reasoning_tokens=counts["reasoning_tokens"],
        estimated_token_cost_usd=cost_usd,
        estimated_token_cost_try=usd_to_try(cost_usd, rate=rate),
        usd_try_rate=rate,
        web_search=bool(web_search),
        status="OK",
    )
