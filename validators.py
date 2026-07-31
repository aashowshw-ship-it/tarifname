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

    for claim in draft.get("dependent_system_claims") or []:
        if re.search(r"(?:yapmasıdır|etmesidir|belirlemesidir)\.?$", claim.strip(), re.I):
            findings.append({"level": "Hata", "message": "Sistem alt istemi yanlış fiil sonuyla bitiyor."})

    if not findings:
        findings.append({"level": "Uygun", "message": "Otomatik kontrollerde belirgin hata bulunmadı."})
    return findings
