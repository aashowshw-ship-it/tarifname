from __future__ import annotations

import json
import os
import re
from typing import Any

from openai import OpenAI


def _extract_json(text: str) -> dict[str, Any]:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start >= 0 and end > start:
            return json.loads(text[start : end + 1])
        raise


def client_from_env() -> OpenAI:
    key = os.getenv("OPENAI_API_KEY", "").strip()
    if not key:
        raise RuntimeError("OPENAI_API_KEY tanımlı değil.")
    return OpenAI(api_key=key)


def ask_json(prompt: str, *, use_web: bool = False) -> dict[str, Any]:
    client = client_from_env()
    model = os.getenv("OPENAI_MODEL", "gpt-5.6")
    kwargs: dict[str, Any] = {
        "model": model,
        "input": prompt,
    }
    if use_web:
        kwargs["tools"] = [{"type": "web_search"}]
    response = client.responses.create(**kwargs)
    return _extract_json(response.output_text)
