from __future__ import annotations

import re
from typing import Any


def _normalize_semantics(text: str) -> set[str]:
    txt = str(text or "").casefold()
    txt = re.sub(r"^\s*istem\s+\d+(?:\s*(?:veya|ve|,)\s*\d+)*['’]?e\s+uygun\s+[^;]+;", " ", txt)
    txt = re.sub(r"\(\s*[a-z0-9_\-]+\s*\)", " ", txt)
    txt = re.sub(r"[^a-zçğıöşü0-9]+", " ", txt)
    stop = {"istem", "uygun", "olup", "özelliği", "bir", "ve", "veya", "ile", "olan", "olarak", "söz", "konusu", "şekilde", "şeklinde"}
    return {w for w in txt.split() if len(w) > 2 and w not in stop}



def _system_claim_entry_texts(entry: Any, include_group_lead: bool = True) -> list[str]:
    if isinstance(entry, dict):
        out: list[str] = []
        lead = str(entry.get("lead", "") or "").strip()
        if include_group_lead and lead:
            out.append(lead)
        out.extend(str(x or "").strip() for x in (entry.get("subelements") or []) if str(x or "").strip())
        return out
    text = str(entry or "").strip()
    return [text] if text else []


def _system_claim_all_texts(system_claim: dict[str, Any] | None) -> list[str]:
    out: list[str] = []
    for entry in (system_claim or {}).get("elements") or []:
        out.extend(_system_claim_entry_texts(entry))
    return out


def _reference_name_pattern(name: str) -> re.Pattern:
    """Match the reference-list element name immediately before a (N), allowing Turkish inflection."""
    tokens = re.findall(r"[A-Za-zÇĞİÖŞÜçğıöşü0-9]+", str(name or ""))
    parts = []
    for token in tokens:
        stem = token if len(token) <= 4 else token[:max(4, len(token) - 2)]
        parts.append(re.escape(stem) + r"\w*")
    return re.compile(r"\s+".join(parts) + r"\s*$", re.I)


def _reference_identity_findings(draft: dict[str, Any]) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    element_map = {str(x.get("number", "")).strip(): str(x.get("name", "")).strip() for x in (draft.get("elements") or []) if str(x.get("number", "")).strip()}
    texts = [
        *map(str, draft.get("detailed_paragraphs") or []),
        *_system_claim_all_texts(draft.get("system_claim") or {}),
        *map(str, draft.get("dependent_system_claims") or []),
        *map(str, (draft.get("method_claim") or {}).get("steps") or []),
        *map(str, draft.get("dependent_method_claims") or []),
    ]
    for text in texts:
        for m in re.finditer(r"\(([^()]+)\)", text):
            n = m.group(1).strip()
            if n not in element_map:
                continue
            before = text[max(0, m.start() - 140):m.start()].rstrip()
            if not _reference_name_pattern(element_map[n]).search(before):
                out.append({"level": "Hata", "message": f"Referans ({n}) REFERANS NUMARALARI listesindeki '{element_map[n]}' unsur adı/çekimli biçimi yerine kısaltma veya farklı adla kullanılmış."})
    return out




def _reference_mention_pattern(name: str) -> re.Pattern:
    """Canonical element-name mention, allowing Turkish inflection on the final token."""
    tokens = re.findall(r"[A-Za-zÇĞİÖŞÜçğıöşü0-9]+", str(name or ""))
    if not tokens:
        return re.compile(r"a^")
    fixed = [re.escape(x) for x in tokens[:-1]]
    last = tokens[-1]
    stem = last if len(last) <= 4 else last[:max(4, len(last) - 2)]
    fixed.append(re.escape(stem) + r"\w*")
    return re.compile(r"\b" + r"\s+".join(fixed) + r"\b", re.I)


def _reference_presence_findings(draft: dict[str, Any]) -> list[dict[str, str]]:
    """From DETAILED DESCRIPTION onward, canonical referenced elements may not appear without (N)."""
    out: list[dict[str, str]] = []
    elements = [x for x in (draft.get("elements") or []) if str(x.get("number", "")).strip() and str(x.get("name", "")).strip()]
    texts: list[tuple[str, str]] = []
    for i, text in enumerate(draft.get("detailed_paragraphs") or [], start=1):
        texts.append((f"Detaylı açıklama paragrafı {i}", str(text or "")))
    for i, step in enumerate(draft.get("method_steps") or [], start=1):
        texts.append((f"Yöntem adımı {i}", str(step.get("text", "") or "")))
    if draft.get("working_principle"):
        texts.append(("Çalışma prensibi", str(draft.get("working_principle") or "")))
    sc = draft.get("system_claim") or {}
    texts.extend(("Ana sistem istemi", t) for t in _system_claim_all_texts(sc))
    texts.extend((f"Bağımlı sistem istemi {i}", str(t or "")) for i, t in enumerate(draft.get("dependent_system_claims") or [], start=1))
    mc = draft.get("method_claim") or {}
    texts.extend(("Ana yöntem istemi", str(t or "")) for t in (mc.get("steps") or []))
    texts.extend((f"Bağımlı yöntem istemi {i}", str(t or "")) for i, t in enumerate(draft.get("dependent_method_claims") or [], start=1))

    for element in elements:
        number = str(element.get("number", "")).strip()
        name = str(element.get("name", "")).strip()
        mention_re = _reference_mention_pattern(name)
        ref_re = re.compile(r"^\s*(?:\([^)]{1,40}\)\s*)?\(\s*" + re.escape(number) + r"\s*\)")
        for label, text in texts:
            for m in mention_re.finditer(text):
                after = text[m.end():m.end() + 70]
                if not ref_re.match(after):
                    out.append({
                        "level": "Hata",
                        "message": f"{label}: '{name}' unsurunun kullanımı ({number}) referansını taşımıyor. BULUŞUN DETAYLI AÇIKLAMASI ve istemlerde referans-listesi unsurları her kullanımda numaralandırılmalıdır.",
                    })
                    break
    return out


def _common_carrier_scope_findings(draft: dict[str, Any]) -> list[dict[str, str]]:
    """Do not group passive stores/databases under an executable-software carrier unless source description says executable."""
    out: list[dict[str, str]] = []
    element_map = {str(x.get("number", "")).strip(): x for x in (draft.get("elements") or [])}
    passive_re = re.compile(r"veritaban|bellek|hafıza|veri depos|data store|profil tablos|kayıt tablos|veri yapıs", re.I)
    executable_re = re.compile(r"yazılım|modül|kontrolör|arayüz|yığın|stack|algoritma|koştur|çalıştır|yürüt", re.I)
    for entry in (draft.get("system_claim") or {}).get("elements") or []:
        if not isinstance(entry, dict):
            continue
        lead = str(entry.get("lead", "") or "")
        if not re.search(r"koşturulan|çalışan|yürütülen|executed|running", lead, re.I):
            continue
        for sub in entry.get("subelements") or []:
            text = str(sub or "")
            refs = re.findall(r"\(([^()]+)\)", text)
            for ref in refs:
                info = element_map.get(str(ref).strip()) or {}
                name = str(info.get("name", "") or "")
                desc = str(info.get("description", "") or "")
                if passive_re.search(name) and not executable_re.search(desc):
                    out.append({
                        "level": "Hata",
                        "message": f"Ortak yazılım taşıyıcı grubunda pasif/veri taşıyan unsur kullanılmış: {name} ({ref}). Kaynak bu unsuru yürütülebilir yazılım/modül olarak açıkça tanımlamıyorsa ortak grubun dışına çıkarılmalıdır.",
                    })
    return out


def _main_claim_first_definition_findings(draft: dict[str, Any]) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    elements = draft.get("elements") or []
    order = {str(x.get("number", "")).strip(): i for i, x in enumerate(elements) if str(x.get("number", "")).strip()}
    seen: set[str] = set()
    last_order = -1

    def check(text: str, label: str) -> None:
        nonlocal last_order
        refs = [x.strip() for x in re.findall(r"\(([^()]+)\)", str(text)) if x.strip() in order]
        new_refs: list[str] = []
        for ref in refs:
            if ref not in seen and ref not in new_refs:
                new_refs.append(ref)
        if len(new_refs) > 1:
            out.append({"level": "Hata", "message": f"Ana sistem istemi {label} henüz tanımlanmamış birden fazla referanslı unsuru aynı anda kullanıyor ({', '.join(new_refs)}). Her düz/alt madde tek yeni referanslı unsur tanımlamalı."})
            return
        if new_refs:
            ref = new_refs[0]
            if order[ref] < last_order:
                out.append({"level": "Hata", "message": f"Ana sistem isteminde ({ref}) unsuru kaynak/teknik tanım sırasının gerisinde ilk kez tanımlanıyor."})
            else:
                last_order = order[ref]
                seen.add(ref)

    for idx, entry in enumerate((draft.get("system_claim") or {}).get("elements") or [], start=1):
        if isinstance(entry, dict):
            lead = str(entry.get("lead", "") or "")
            if any(x in order for x in re.findall(r"\(([^()]+)\)", lead)):
                out.append({"level":"Hata","message":f"Ana istem {idx}. ortak taşıyıcı üst maddesi referans numarası taşıyamaz; referanslı modüller alt maddelerde ayrı tanımlanmalıdır."})
            subs = [str(x or "") for x in (entry.get("subelements") or []) if str(x or "").strip()]
            if len(subs) < 2:
                out.append({"level":"Hata","message":"Ortak taşıyıcı istem grubu en az iki ayrı alt unsur içermelidir."})
            for j, sub in enumerate(subs, start=1):
                check(sub, f"{idx}.{j}. alt maddesi")
        else:
            check(str(entry), f"{idx}. maddesi")
    return out




def _method_step_language_findings(draft: dict[str, Any]) -> list[dict[str, str]]:
    """Türkçe yöntem adımı gerçek işlem fiilimsisiyle bitmeli; salt isim/adlandırma kabul edilmez."""
    out: list[dict[str, str]] = []
    action_end_re = re.compile(r"(?:ması|mesi)\s*$", re.IGNORECASE)
    for step in draft.get("method_steps") or []:
        number = str(step.get("number", "") or "").strip()
        text = str(step.get("text", "") or "").strip()
        clean = re.sub(r"\s*\(\s*" + re.escape(number) + r"\s*\)\s*$", "", text).strip().rstrip(".,;:") if number else text.rstrip(".,;:")
        if clean and not action_end_re.search(clean):
            out.append({"level":"Hata","message":f"Yöntem işlem adımı {number or '?'} gerçek bir işlem fiilimsisiyle bitmiyor: '{clean}'. Adım '... yapılması/edilmesi/aktarılması/belirlenmesi' gibi bir işlem sonucu ile bitmelidir."})
    method = draft.get("method_claim") or {}
    for item in method.get("steps") or []:
        text = str(item or "").strip().rstrip(".,;:")
        text = re.sub(r"\s*\(\s*[^()]+\s*\)\s*$", "", text).strip()
        if text and not action_end_re.search(text):
            out.append({"level":"Hata","message":f"Bağımsız yöntem istemindeki işlem adımı gerçek bir işlem fiilimsisiyle bitmiyor: '{text}'."})
    return out


def _generic_claim_term_findings(draft: dict[str, Any]) -> list[dict[str, str]]:
    """İstemde teknik eleman türü yerine belirsiz 'unsur' placeholder'ı kullanılmasını engeller."""
    out: list[dict[str, str]] = []
    texts = [*_system_claim_all_texts(draft.get("system_claim") or {}), *map(str, draft.get("dependent_system_claims") or []), *map(str, (draft.get("method_claim") or {}).get("steps") or []), *map(str, draft.get("dependent_method_claims") or [])]
    for text in texts:
        if re.search(r"\bbir\s+unsur\b|\bunsur\s+olmasıdır|\bunsur\s+içermesidir", str(text), re.IGNORECASE):
            out.append({"level":"Hata","message":"İstemde teknik eleman türü yerine belirsiz 'unsur' ifadesi kullanılmış. Kaynağa göre anten/modül/birim/eleman/sunucu/veritabanı gibi gerçek teknik tür yazılmalıdır."})
    return out


def _semantic_repeat_findings(base_text: str, dependents: list[str], label: str) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    previous = [_normalize_semantics(base_text)]
    for idx, claim in enumerate(dependents, start=1):
        words = _normalize_semantics(claim)
        if len(words) >= 4:
            for prior in previous:
                if len(words & prior) / max(1, len(words)) >= 0.92:
                    out.append({"level": "Hata", "message": f"{label} bağımlı istem {idx} ana/üst istemdeki teknik özelliği semantik olarak tekrar ediyor; gerçek ek sınırlama gerekir."})
                    break
        previous.append(words)
    return out


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

    generic_names = {"diğer parçalar", "diğer parça", "diğer elemanlar", "çeşitli parçalar", "çeşitli elemanlar"}
    for name in names:
        if name.casefold() in generic_names:
            findings.append({"level": "Hata", "message": "Belirsiz referans unsuru kullanılmış: ‘Diğer parçalar/Diğer elemanlar’. Teknik unsur net adlandırılmalı."})

    if re.search(r"müşteri tarafından iletilen(?: teknik)? (?:çizim|belge)|müşteri bilgilerine göre|ek teknik belgede|iletilen teknik çizimde", combined, re.I):
        findings.append({"level": "Hata", "message": "Kullanıcıya görünen tarifname metninde kaynak/iletilen belge atfı bulunmamalı."})

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
    claim_texts.extend(_system_claim_all_texts(sc))
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

    for entry in (sc.get("elements") or []):
        if isinstance(entry, dict):
            lead = str(entry.get("lead", "") or "").strip()
            if not re.search(r"\bve\s*;$", lead, re.I):
                findings.append({"level":"Hata","message":"Ortak taşıyıcı üst maddesi Türkçe istemde ‘... ve;’ biçiminde alt maddeleri başlatmalıdır."})
            if re.search(r"\([^()]+\)", lead):
                findings.append({"level":"Hata","message":"Ortak taşıyıcı üst maddesi referans numarası taşıyamaz."})
            for sub in entry.get("subelements") or []:
                if ";" in str(sub):
                    findings.append({"level":"Hata","message":"Ortak taşıyıcı alt maddelerinde noktalı virgül kullanılmamalıdır."})

    for claim in draft.get("dependent_system_claims") or []:
        if re.search(r"(?:yapmasıdır|etmesidir|belirlemesidir|bulunmasıdır|oluşturulmasıdır|bağlanmasıdır|sağlanmasıdır|gerçekleştirilmesidir|yapılmasıdır|edilmesidir)\.?$", claim.strip(), re.I):
            findings.append({"level": "Hata", "message": "Yöntem dışındaki alt istem yanlış eylem/işlem sonuyla bitiyor; ‘olmasıdır.’ veya ‘içermesidir.’ kullanılmalı."})
        elif not re.search(r"(?:olmasıdır|içermesidir)\.?$", claim.strip(), re.I):
            findings.append({"level": "Hata", "message": "Yöntem dışındaki alt istem ‘olmasıdır.’ veya ‘içermesidir.’ ile bitmeli."})

    for dep_index, claim in enumerate(draft.get("dependent_method_claims") or [], 1):
        text = str(claim or "").strip()
        if not re.search(r"işlem adım(?:ını|larını)\s+içermesidir\.?$", text, re.I):
            findings.append({"level": "Hata", "message": f"Bağımlı yöntem istemi {dep_index}, `işlem adımını içermesidir.` veya `işlem adımlarını içermesidir.` ile bitmelidir."})

    for claim in [*(draft.get("dependent_system_claims") or []), *(draft.get("dependent_method_claims") or [])]:
        if re.search(r"önceki\s+istemlerden\s+herhangi\s+birine", str(claim), re.I):
            findings.append({"level": "Hata", "message": "Bağımlı istemde ‘Önceki istemlerden herhangi birine’ kullanılmış; ek özelliğin dayandığı doğrudan istem numarası seçilmeli."})

    findings.extend(_reference_identity_findings(draft))
    findings.extend(_reference_presence_findings(draft))
    findings.extend(_main_claim_first_definition_findings(draft))
    findings.extend(_common_carrier_scope_findings(draft))
    findings.extend(_generic_claim_term_findings(draft))
    findings.extend(_method_step_language_findings(draft))

    for claim in draft.get("dependent_system_claims") or []:
        if re.search(r"sistemin[,\s].*?(?:çalışmaya|kullanılmaya) uygun bir sistem olmasıdır\.?$", str(claim), re.I):
            findings.append({"level": "Hata", "message": "Bağımlı sistem istemi yalnız çalışma/kullanım ortamına uygunlukla tanımlanmış; ortam özelliği somut teknik unsurun niteliği/bağlantısı olarak yazılmalı."})
    for claim in draft.get("dependent_method_claims") or []:
        if re.search(r"yöntemin .*?(?:ortamda|şebekede|yapıda).*?gerçekleştirilmesidir\.?$", str(claim), re.I):
            findings.append({"level": "Hata", "message": "Bağımlı yöntem istemi yalnız gerçekleştirme ortamını söylüyor; ortam gerçek bir işlem adımı, girdi veya teknik taşıyıcı ile ilişkilendirilmeli."})

    hardware_anchor_re = re.compile(r"elektronik cihaz|elektronik işlem birimi|işlemci|donanım|bilgisayar|mikrodenetleyici|kontrol birimi", re.I)
    software_terms_re = re.compile(r"modül|birim|algoritma|yazılım|veri işleme|hesaplama", re.I)
    sc_text = " ".join([str(sc.get("preamble", "")), *_system_claim_all_texts(sc)])
    if sc_text and len(software_terms_re.findall(sc_text)) >= 2 and not hardware_anchor_re.search(sc_text):
        findings.append({"level": "Hata", "message": "Yazılım/modül ağırlıklı bağımsız sistem istemi elektronik cihaz/işlemci gibi geniş bir donanımsal taşıyıcıya dayandırılmalı."})
    if method:
        mc_text = " ".join([str(method.get("preamble", "")), *map(str, method.get("steps") or [])])
        if len(software_terms_re.findall(mc_text)) >= 2 and not hardware_anchor_re.search(mc_text):
            findings.append({"level": "Hata", "message": "Yazılım/algoritma ağırlıklı bağımsız yöntem istemi elektronik cihaz/işlemci gibi geniş bir donanımsal taşıyıcıya dayandırılmalı."})

    execution_relation_re = re.compile(r"üzerinde\s+(?:çalışan|koşturulan|yürütülen)|içerisinde\s+(?:çalışan|koşturulan|yürütülen)|vasıtasıyla|tarafından\s+(?:çalıştırılan|yürütülen)", re.I)
    if sc_text and len(software_terms_re.findall(sc_text)) >= 2 and not execution_relation_re.search(sc_text):
        findings.append({"level": "Hata", "message": "Yazılım/modül ağırlıklı bağımsız sistem isteminde teknik taşıyıcı ile modül/yazılım arasında açık çalışma/koşturma ilişkisi kurulmalı; yalnız işlemci/donanım kelimesi yeterli değildir."})

    system_base = " ".join([str(sc.get("preamble", "")), *_system_claim_all_texts(sc)])
    findings.extend(_semantic_repeat_findings(system_base, [str(x or "") for x in (draft.get("dependent_system_claims") or [])], "Sistem"))
    if method:
        method_base = " ".join([str(method.get("preamble", "")), *map(str, method.get("steps") or [])])
        findings.extend(_semantic_repeat_findings(method_base, [str(x or "") for x in (draft.get("dependent_method_claims") or [])], "Yöntem"))

    coverage_map = draft.get("source_coverage_map")
    if coverage_map is not None:
        for row in coverage_map:
            if row.get("covered") is not True or not (row.get("sections") or []) or not str(row.get("evidence", "") or "").strip():
                findings.append({"level": "Hata", "message": f"BBF teknik bilgi kapsam kaydı eksik/kanıtsız: {row.get('fact_id','?')}"})

    if not findings:
        findings.append({"level": "Uygun", "message": "Otomatik kontrollerde belirgin hata bulunmadı."})
    return findings
