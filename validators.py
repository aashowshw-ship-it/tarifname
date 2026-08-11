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

    tf_raw = str(draft.get("technical_field", "") or "").strip()
    tf_parts = [x.strip() for x in re.split(r"\n\s*\n", tf_raw) if x.strip()]
    if len(tf_parts) < 2:
        findings.append({"level": "Hata", "message": "TEKNİK ALAN iki paragraf olmalı: ilk paragraf yalnız ‘Buluş, ... ile ilgilidir.’, ikinci paragraf ‘Buluş, özellikle ...’ ile başlamalı."})
    else:
        if not re.fullmatch(r"Buluş,\s+.+?ile ilgilidir\.", tf_parts[0], re.I | re.S):
            findings.append({"level": "Hata", "message": "TEKNİK ALAN ilk paragrafı yalnız ‘Buluş, ... ile ilgilidir.’ giriş cümlesinden oluşmalı."})
        if not re.match(r"^Buluş,\s*özellikle\b", tf_parts[1], re.I):
            findings.append({"level": "Hata", "message": "TEKNİK ALAN ikinci paragrafı ‘Buluş, özellikle ...’ ile başlamalı."})

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
    step_numbers = [str(step.get("number", "") or "").strip() for step in steps]
    if any(not n for n in step_numbers):
        findings.append({"level": "Hata", "message": "Yöntem işlem adımlarında boş referans var. Kaynakta referans yoksa 1001... varsayılanı atanmalı; müşteri referansı varsa aynen korunmalı."})
    if len(step_numbers) != len(set(step_numbers)):
        findings.append({"level": "Hata", "message": "Tekrarlanan yöntem işlem adımı referansı var."})
    if method:
        claim_text = " ".join(method.get("steps") or [])
        for step in steps:
            n = str(step.get("number", ""))
            if n and f"({n})" not in claim_text:
                findings.append({"level": "Hata", "message": f"Yöntem isteminde ({n}) işlem referansı yok."})

        claim_steps = list(method.get("steps") or [])
        if len(claim_steps) == len(steps):
            for idx, step in enumerate(steps):
                n = str(step.get("number", ""))
                expected_text = str(step.get("text", "")).strip().rstrip(".,;:")
                actual = re.sub(rf"\s*\({re.escape(n)}\)\s*$", "", str(claim_steps[idx]).strip().rstrip(".,;:")).strip().rstrip(".,;:")
                if n and expected_text != actual:
                    findings.append({"level": "Hata", "message": f"({n}) yöntem adımı REFERANS NUMARALARI/detaylı açıklama ile bağımsız yöntem isteminde birebir aynı değil."})

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


    for table in draft.get("tables") or []:
        headers = [str(x or "").casefold() for x in (table.get("headers") or [])]
        if any("işlem adımı" in h for h in headers) and any("gerçekleştiren unsur" in h for h in headers):
            findings.append({"level": "Hata", "message": "Sistem-yöntem ilişki tablosu tablo olarak bırakılmamalı; doğal teknik paragrafa dönüştürülmeli."})

    for claim in draft.get("dependent_system_claims") or []:
        if re.search(r"(?:yapmasıdır|etmesidir|belirlemesidir)\.?$", claim.strip(), re.I):
            findings.append({"level": "Hata", "message": "Sistem alt istemi yanlış fiil sonuyla bitiyor."})

    for claim in [*(draft.get("dependent_system_claims") or []), *(draft.get("dependent_method_claims") or [])]:
        if re.search(r"önceki\s+istemlerden\s+herhangi\s+birine", str(claim), re.I):
            findings.append({"level": "Hata", "message": "Bağımlı istemde ‘Önceki istemlerden herhangi birine’ kullanılmış; ek özelliğin dayandığı doğrudan istem numarası seçilmeli."})

    hardware_anchor_re = re.compile(r"elektronik cihaz|elektronik işlem birimi|işlemci|donanım|bilgisayar|mikrodenetleyici|kontrol birimi", re.I)
    software_terms_re = re.compile(r"modül|birim|algoritma|yazılım|veri işleme|hesaplama", re.I)
    sc_text = " ".join([str(sc.get("preamble", "")), *map(str, sc.get("elements") or [])])
    if sc_text and len(software_terms_re.findall(sc_text)) >= 2 and not hardware_anchor_re.search(sc_text):
        findings.append({"level": "Hata", "message": "Yazılım/modül ağırlıklı bağımsız sistem istemi elektronik cihaz/işlemci gibi geniş bir donanımsal taşıyıcıya dayandırılmalı."})
    if method:
        mc_text = " ".join([str(method.get("preamble", "")), *map(str, method.get("steps") or [])])
        if len(software_terms_re.findall(mc_text)) >= 2 and not hardware_anchor_re.search(mc_text):
            findings.append({"level": "Hata", "message": "Yazılım/algoritma ağırlıklı bağımsız yöntem istemi elektronik cihaz/işlemci gibi geniş bir donanımsal taşıyıcıya dayandırılmalı."})

    if not findings:
        findings.append({"level": "Uygun", "message": "Otomatik kontrollerde belirgin hata bulunmadı."})
    return findings
