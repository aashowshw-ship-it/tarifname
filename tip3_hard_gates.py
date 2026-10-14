"""v5.4.89: Tip-3 fail-closed research, problem-solution and language checks.

This file does not certify independent retrieval; it records and checks four
separately returned search results. Model self-declared completion booleans are
never a substitute for the phase outputs. Human patent judgment remains necessary.
"""
from __future__ import annotations

import hashlib
from copy import deepcopy
import json
import re
from typing import Any

ROUND_NAMES = ("GLOBAL_RECALL", "GAP_SEARCH", "FAMILY_NEIGHBOUR", "D1_CHALLENGE")
PS_FIELDS = (
    "d1_distinguishing_features", "technical_effect", "objective_technical_problem",
    "d2_teaching", "combination_motivation", "remaining_features", "inventive_step_conclusion",
)
SECOND_REVIEW_FIELDS = (
    "source_support", "d1_closest_prior_art", "distinguishing_features",
    "objective_problem_no_solution_hindsight", "d2_motivation", "remaining_features_and_effect",
    "result_consistency", "language_and_terminology",
)
# Deliberately narrow, high-certainty disallowed forms; the independent
# language audit is obligatory for errors outside these recognisable patterns.
_TURKISH_ERRORS = (
    (r"(?u)\btestinı\b", "testini"),
    (r"(?u)\btestinın\b", "testinin"),
    (r"(?u)\btestının\b", "testinin"),
    (r"(?u)\btestya\b", "teste"),
    (r"(?u)\bedilmesinı\b", "edilmesini"),
    (r"(?u)\bdoğrulamasınını\b", "doğrulamasını"),
    (r"(?u)\b(?:yapılmasınını|sağlanmasınını)\b", "ekleri kontrol edin"),
)

def _norm(value: str) -> str:
    return " ".join(str(value or "").split()).casefold()

def _sha_obj(payload: Any) -> str:
    return hashlib.sha256(json.dumps(payload, sort_keys=True, ensure_ascii=False, default=str).encode()).hexdigest()

def _check_phase_contents(recall: dict, gap: dict, neighbour: dict, final: dict) -> None:
    if len(recall.get("candidate_documents") or []) < 20:
        raise ValueError("GLOBAL_RECALL FAIL: yirmi gerçek aday bulunmadı.")
    axes = recall.get("search_axes") or []
    if len([a for a in axes if len(str(a).strip()) > 5]) < 8:
        raise ValueError("GLOBAL_RECALL FAIL: sekiz somut arama ekseni yok.")
    if not recall.get("technical_problem") or not recall.get("technical_features") or not recall.get("technical_effects"):
        raise ValueError("GLOBAL_RECALL FAIL: problem/unsur/teknik etki kaydı eksik.")
    for entry in (gap.get("gap_targets") or []):
        if not all(entry.get(key) for key in ("feature_or_relation", "why_uncovered", "search_queries")):
            raise ValueError("GAP_SEARCH FAIL: teknik boşluk, gerekçe veya sorgu eksik.")
    if not gap.get("gap_targets") or not str(gap.get("gap_search_note") or "").strip():
        raise ValueError("GAP_SEARCH FAIL: araştırma boşluğu gerekçesi eksik.")
    if len(neighbour.get("provisional_shortlist") or []) < 8 or not neighbour.get("provisional_d1"):
        raise ValueError("FAMILY_NEIGHBOUR FAIL: geçici liste/D1 eksik.")
    if not str(neighbour.get("family_neighbour_note") or "").strip():
        raise ValueError("FAMILY_NEIGHBOUR FAIL: aile/atıf/sınıflandırma değerlendirmesi boş.")
    if not str(final.get("research_quality_audit", {}).get("challenge_note") or "").strip():
        raise ValueError("D1_CHALLENGE FAIL: karşı aday ve kombinasyon gerekçesi eksik.")

def seal_four_rounds(recall: dict, gap: dict, neighbour: dict, final: dict) -> dict:
    """Generate evidence exclusively from the four actual returned payloads."""
    _check_phase_contents(recall, gap, neighbour, final)
    return {
        "schema": "tip3-four-round-evidence-1",
        # Full outputs, not only opaque completion flags. Inspectable in job state.
        "raw_rounds": {name: deepcopy(payload) for name, payload in zip(ROUND_NAMES, (recall, gap, neighbour, final))},
        "rounds": [
            {"name": name, "input_output_sha256": _sha_obj(payload), "status": "PASS"}
            for name, payload in zip(ROUND_NAMES, (recall, gap, neighbour, final))
        ],
        "final_list_sha256": _sha_obj(final.get("documents") or []),
        "d1": str(final.get("proposed_d1") or ""),
        "d2": str(final.get("proposed_d2") or ""),
        "provenance": "four-separate-model-search-results",
    }

def require_four_rounds(result: dict) -> None:
    proof = (result or {}).get("_research_round_evidence") or {}
    rounds = proof.get("rounds") or []
    if proof.get("schema") != "tip3-four-round-evidence-1" or len(rounds) != 4:
        raise ValueError("Tip 3 dört tur kapısı FAIL/NOT_RUN: gerçek dört tur kanıtı yok.")
    if [r.get("name") for r in rounds] != list(ROUND_NAMES) or any(r.get("status") != "PASS" or not re.fullmatch(r"[a-f0-9]{64}", str(r.get("input_output_sha256") or "")) for r in rounds):
        raise ValueError("Tip 3 dört tur kapısı FAIL: kayıtlı tur bütünlüğü bozuk.")
    raw = proof.get("raw_rounds") or {}
    if not isinstance(raw, dict) or any(not isinstance(raw.get(name), dict) for name in ROUND_NAMES):
        raise ValueError("Tip 3 dört tur kapısı FAIL: gerçek ham tur çıktıları yok.")
    for item in rounds:
        if _sha_obj(raw[item["name"]]) != item["input_output_sha256"]:
            raise ValueError("Tip 3 dört tur kapısı FAIL: kaynak tur içeriği değiştirilmiş.")
    _check_phase_contents(*(raw[name] for name in ROUND_NAMES))
    if (proof.get("final_list_sha256") != _sha_obj(result.get("documents") or []) or
        proof.get("d1") != str(result.get("proposed_d1") or "") or
        proof.get("d2") != str(result.get("proposed_d2") or "") or
        raw["D1_CHALLENGE"].get("documents") != result.get("documents") or
        raw["D1_CHALLENGE"].get("proposed_d1") != result.get("proposed_d1") or
        raw["D1_CHALLENGE"].get("proposed_d2") != result.get("proposed_d2")):
        raise ValueError("Tip 3 dört tur kapısı FAIL: son 10 veya D1/D2 tur sonrasında değişmiş.")

def _iter_report_turkish_text(report: dict):
    for k in ("evaluation_intro", "purpose", "title"):
        if str(report.get(k) or "").strip():
            yield k, str(report[k])
    for key in ("inventive_step_paragraphs", "conclusion_paragraphs", "warnings"):
        for i, value in enumerate(report.get(key) or []):
            if str(value or "").strip():
                yield f"{key}[{i}]", str(value)
    for j, doc in enumerate(report.get("documents") or []):
        for key in ("description", "novelty_assessment"):
            for i, value in enumerate(doc.get(key) or []):
                if str(value or "").strip():
                    yield f"documents[{j}].{key}[{i}]", str(value)
        for i, row in enumerate(doc.get("comparison_rows") or []):
            for key in ("feature", "status_evidence"):
                if str(row.get(key) or "").strip():
                    yield f"documents[{j}].comparison_rows[{i}].{key}", str(row[key])

def check_turkish_forms(report: dict) -> None:
    for location, string in _iter_report_turkish_text(report):
        for regex, replacement in _TURKISH_ERRORS:
            if re.search(regex, string):
                raise ValueError(f"Tip 3 dil kapısı FAIL: {location}, yazım/ek hatası: {re.search(regex,string).group(0)!r}; öneri: {replacement!r}.")
        # Mixed-case OCR and machine-generated malformed circumflex variants.
        if re.search(r"\b\w*(?:testinı|testinın|testının|testya|edilmesinı)\b", string.lower()):
            raise ValueError(f"Tip 3 dil kapısı FAIL: {location}, yanlış Türkçe çekim tespit edildi.")

def check_problem_solution(report: dict, *, conclusion_override: str | None = None) -> None:
    parts = [str(x).strip() for x in (report.get("conclusion_paragraphs") or []) if str(x).strip()]
    conclusion = conclusion_override if conclusion_override is not None else " ".join(parts)
    if not conclusion.strip():
        raise ValueError("Tip 3 problem-çözüm kapısı FAIL: SONUÇ boş.")
    ps = report.get("problem_solution")
    has_d2 = len(report.get("documents") or []) >= 2
    required_fields = PS_FIELDS if has_d2 else tuple(k for k in PS_FIELDS if k not in ("d2_teaching", "combination_motivation"))
    if not isinstance(ps, dict) or any(not str(ps.get(k) or "").strip() for k in required_fields):
        raise ValueError("Tip 3 problem-çözüm kapısı FAIL: gerekçe bileşenleri eksik.")
    # Structural anchor requirement prevents metadata-only or heading-only PASS.
    for key in required_fields:
        val = _norm(ps[key])
        if len(val) < 22 or val not in _norm(conclusion):
            raise ValueError(f"Tip 3 problem-çözüm kapısı FAIL: {key}, SONUÇ metnindeki gerçek ifadeyle eşleşmiyor.")
    folded = _norm(conclusion)
    for marker in (("d1", "d2", "objektif teknik problem", "teknik etki", "buluş basamağı") if has_d2 else ("d1", "objektif teknik problem", "teknik etki", "buluş basamağı")):
        if marker not in folded:
            raise ValueError("Tip 3 problem-çözüm kapısı FAIL: SONUÇ içinde " + marker + " bulunmuyor.")
    if len(conclusion) < 650:
        raise ValueError("Tip 3 problem-çözüm kapısı FAIL: gerekçe yüzeysel.")

def validate_second_review(audit: dict) -> None:
    if not isinstance(audit, dict):
        raise ValueError("Tip 3 bağımsız ikinci uzman kapısı NOT_RUN.")
    errors = []
    for key in SECOND_REVIEW_FIELDS:
        val = audit.get(key)
        if not isinstance(val, dict) or val.get("pass") is not True or len(str(val.get("note") or "").strip()) < 30:
            errors.append(key)
    if audit.get("overall_pass") is not True or errors:
        raise ValueError("Tip 3 bağımsız ikinci uzman kapısı FAIL/NOT_RUN: " + ", ".join(errors or ["overall_pass"]))
    if audit.get("unresolved_issues"):
        raise ValueError("Tip 3 bağımsız ikinci uzman kapısı FAIL: çözülmemiş meseleler: " + str(audit["unresolved_issues"]))


def check_docx_tip3(docx, *, ps_paragraph_index: int = 85) -> None:
    """Independently inspect the final physical document, not only JSON fields."""
    from docx import Document
    import io
    doc = Document(io.BytesIO(docx))
    conclusion = doc.paragraphs[ps_paragraph_index].text
    has_d2 = bool(doc.paragraphs[57].text.strip())
    markers = ("D1", "D2", "objektif teknik problem", "teknik etki", "buluş basamağı") if has_d2 else ("D1", "objektif teknik problem", "teknik etki", "buluş basamağı")
    for marker in markers:
        if marker.casefold() not in conclusion.casefold():
            raise ValueError("Tip 3 son Word problem-çözüm kapısı FAIL: " + marker)
    if len(conclusion.strip()) < 650:
        raise ValueError("Tip 3 son Word problem-çözüm kapısı FAIL: değerlendirme kısa.")
    for i in (42, 55, 59, 71, 75, 77, 79, 81, 85):
        string = doc.paragraphs[i].text
        for regex, _ in _TURKISH_ERRORS:
            if re.search(regex, string):
                raise ValueError(f"Tip 3 son Word Türkçe dil kapısı FAIL: paragraf {i}")
    for ti in (2,3):
        for row in doc.tables[ti].rows:
            for cell in row.cells:
                for regex, _ in _TURKISH_ERRORS:
                    if re.search(regex, cell.text):
                        raise ValueError(f"Tip 3 son Word Türkçe dil kapısı FAIL: tablo {ti}")
