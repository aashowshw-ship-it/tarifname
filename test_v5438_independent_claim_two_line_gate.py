from __future__ import annotations

import pytest

import app_core
from rules import APP_VERSION, RULESET_VERSION, TARIFNAME_RULES


def test_versions_and_two_line_rule_are_binding():
    assert APP_VERSION == "v5.4.48"
    assert RULESET_VERSION == "2026-09-01.v34"
    low = TARIFNAME_RULES.casefold()
    assert "en az iki fiziksel satır" in low
    assert "yalnız buluş adından" in low
    assert "non-breaking kuyruk kontrolü" in low


def test_render_gate_rejects_one_line_independent_preamble():
    lines = [
        "İSTEMLER",
        "1. Hücrelerin haftalık davranışına göre anormal durum tespit sistemi olup, özelliği;",
        "bir trafik ve performans veri kaynağı (1),",
        "içermesidir.",
        "2. İstem 1’e uygun sistem olup, özelliği; ...",
        "ÖZET",
    ]
    with pytest.raises(ValueError, match="en az iki fiziksel satır"):
        app_core._validate_rendered_independent_claim_preamble_lines(lines)


def test_render_gate_accepts_two_physical_preamble_lines_and_ignores_dependent_claim():
    lines = [
        "İSTEMLER",
        "1. Mobil haberleşme şebekelerinde baz istasyonu hücrelerine ait trafik ve performans verilerini",
        "geçmiş haftalara ait davranış verileriyle karşılaştırarak anormal durumu tespit eden sistem olup, özelliği;",
        "bir trafik ve performans veri kaynağı (1),",
        "içermesidir.",
        "2. İstem 1’e uygun sistem olup, özelliği; ek bir sınırlama içermesidir.",
        "ÖZET",
    ]
    app_core._validate_rendered_independent_claim_preamble_lines(lines)
