from __future__ import annotations

from pathlib import Path
from docx import Document

from rules import APP_VERSION, RULESET_VERSION, TARIFNAME_RULES
from tarifname_figure_generation import protect_turkish_claim_transition

ROOT = Path(__file__).resolve().parent

def test_versions_and_natural_wrap_rule():
    assert APP_VERSION == "v5.4.71"
    assert RULESET_VERSION == "2026-09-18.v64"
    low = TARIFNAME_RULES.casefold()
    assert "doğal satır kaydırması" in low
    assert "non-breaking" in low and "yasak" in low

def test_protector_removes_nbsp_tail_instead_of_creating_it():
    raw = "ultra düşük gecikmeli telemetri iletimini sağlayan dinamik ağ\u00a0optimizasyon\u00a0ve\u00a0yönetim\u00a0sistemi\u00a0olup,\u00a0özelliği;"
    out = protect_turkish_claim_transition(raw)
    assert "\u00a0" not in out
    assert "ağ optimizasyon ve yönetim sistemi olup, özelliği;" in out
