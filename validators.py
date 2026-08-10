from __future__ import annotations

import re
from typing import Any


def validate_draft(draft: dict[str, Any]) -> list[dict[str, str]]:
    findings: list[dict[str, str]] = []
    elements = draft.get("elements") or []
    numbers = [str(x.get("number", "")).strip() for x in elements]
    names = [str(x.get("name", "")).strip() for x in elements]

    if len(numbers) != len(set(numbers)):
        findings.append({"level": "Hata", "message": "Tekrarlanan unsur numarası var."})
    if any(not n for n in numbers):
        findings.append({"level": "Hata", "message": "Numarasız unsur var."})

    combined = "\n".join([
        str(draft.get("technical_field", "")),
        *map(str, draft.get("prior_art_paragraphs") or []),
        str(draft.get("short_description_intro", "")),
        *map(str, draft.get("objectives") or []),
        str(draft.get("unumbered_system_definition", "")),
        *map(str, draft.get("unumbered_system_elements") or []),
    ])
    if re.search(r"\(\s*(?:\d{1,3}|10\d{2})\s*\)", combined):
        findings.append({"level": "Hata", "message": "Referans numaralarından önceki bölümlerde parantezli numara bulundu."})

    for obj in draft.get("objectives") or []:
        ot = str(obj).strip()
        if re.search(r"(?:mak|mek)\.?$", ot, re.I):
            findings.append({"level": "Hata", "message": f"Amaç cümlesi çıplak mastarla bitiyor; tam yüklem kullanılmalı: {ot}"})

    prose_for_semicolon = "\n".join([
        str(draft.get("technical_field", "")),
        *map(str, draft.get("prior_art_paragraphs") or []),
        str(draft.get("short_description_intro", "")),
        *map(str, draft.get("objectives") or []),
        *map(str, draft.get("detailed_paragraphs") or []),
        str(draft.get("abstract", "")),
    ])
    if ";" in prose_for_semicolon:
        findings.append({"level": "Uyarı", "message": "Açıklama metninde noktalı virgül bulundu; zorunlu değilse virgül veya nokta kullanılmalı."})

    for name in names:
        words = [w for w in re.split(r"\s+", name) if w]
        if len(words) >= 3 and sum(1 for w in words if w[:1].isupper()) == len(words):
            findings.append({"level": "Uyarı", "message": f"Unsur adı başlık biçiminde olabilir: {name}"})

    method = draft.get("method_claim")
    steps = draft.get("method_steps") or []
    if method:
        claim_text = " ".join(method.get("steps") or [])
        for step in steps:
            n = str(step.get("number", ""))
            if n and f"({n})" not in claim_text:
                findings.append({"level": "Hata", "message": f"Yöntem isteminde ({n}) işlem referansı yok."})

    claim_texts = []
    sc = draft.get("system_claim") or {}
    claim_texts.extend(sc.get("elements") or [])
    claim_texts.extend(draft.get("dependent_system_claims") or [])
    if method:
        claim_texts.extend(method.get("steps") or [])
    claim_texts.extend(draft.get("dependent_method_claims") or [])
    symbol_requirements = {
        "HPU_W": "hibrit güç ünitesi ağırl",
        "FW_min": "asgari görev yakıt",
        "UW_F": "ilave yakıt",
    }
    for ct in claim_texts:
        low = str(ct).lower()
        for sym, phrase in symbol_requirements.items():
            if sym.lower() in low and phrase not in low:
                findings.append({"level": "Uyarı", "message": f"İstemde {sym} sembolünün teknik açılımı görünmüyor; önce açılımı, sonra parantez içinde sembolü kullan."})

    for claim in draft.get("dependent_system_claims") or []:
        if re.search(r"(?:yapmasıdır|etmesidir|belirlemesidir)\.?$", claim.strip(), re.I):
            findings.append({"level": "Hata", "message": "Sistem alt istemi yanlış fiil sonuyla bitiyor."})

    if not findings:
        findings.append({"level": "Uygun", "message": "Otomatik kontrollerde belirgin hata bulunmadı."})
    return findings
