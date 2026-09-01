from __future__ import annotations

import re
import hashlib
from typing import Any

TECHNICAL_PASSAGE_HINT_RE = re.compile(
    r"\b(?:modül|module|sistem|system|yöntem|method|algoritma|algorithm|simül|monte carlo|MRI|ADI|EHDS|CEWSS|ESI|jamming|karıştırma|spoofing|aldatma|rüzgar|wind|enerji|energy|güdüm|navigation|hedef|target|otonomi|autonomy|görev|mission|platform|eşik|threshold|hesap|formula|formül|oran|ratio|katsayı|coefficient|sinyal|signal|koordinat|coordinate)\b",
    re.IGNORECASE,
)
ADMIN_PASSAGE_HINT_RE = re.compile(
    r"(?:adı\s*soyadı|t\.c\.?\s*kimlik|telefon|e-?posta|ikamet|imza|tebliğ|izleç|yayım\s*trh|revizyon|buluşçu|katkı\s*yüzdesi|doğum\s*tarihi|aselsan\s*özel|elektronik ortamdan alınan kopyadır)",
    re.IGNORECASE,
)


def build_source_passage_registry(source_text: str, technical_supplement_text: str) -> list[dict[str, str]]:
    """Ham kaynaklardan deterministik pasaj kimliği üretir; fact envanterinin kendi eksikliğini gizlemesini engeller."""
    registry: list[dict[str, str]] = []
    for prefix, label, text in (("B", "BBF", source_text), ("E", "EK_TEKNIK", technical_supplement_text)):
        counter = 0
        for raw in str(text or "").splitlines():
            passage = re.sub(r"\s+", " ", raw).strip()
            if not passage or passage.startswith("---") or passage.startswith("[TEKNİK GÖRSEL DOSYASI:"):
                continue
            if len(passage) < 3:
                continue
            counter += 1
            registry.append({"passage_id": f"{prefix}{counter:04d}", "source": label, "text": passage})
    if not registry:
        raise ValueError("Ham kaynak pasaj envanteri oluşturulamadı.")
    return registry


def validate_source_passage_audit(extracted: dict[str, Any], registry: list[dict[str, str]]) -> None:
    """Her ham pasajın tam bir kez fact veya gerekçeli teknik-dışı sınıfına bağlandığını doğrular."""
    expected = {x["passage_id"]: x for x in registry}
    rows = extracted.get("source_passage_audit") or []
    seen: dict[str, dict[str, Any]] = {}
    duplicates: list[str] = []
    for row in rows:
        pid = str(row.get("passage_id", "") or "").strip()
        if not pid:
            continue
        if pid in seen:
            duplicates.append(pid)
        seen[pid] = row
    missing = sorted(set(expected) - set(seen))
    extra = sorted(set(seen) - set(expected))
    if duplicates or missing or extra:
        details: list[str] = []
        if missing:
            details.append("eksik=" + ", ".join(missing[:30]))
        if extra:
            details.append("tanımsız=" + ", ".join(extra[:20]))
        if duplicates:
            details.append("tekrar=" + ", ".join(sorted(set(duplicates))[:20]))
        raise ValueError("Ham kaynak pasaj tamlık kapısı başarısız: " + "; ".join(details))

    fact_ids = {str(f.get("id", "") or "").strip() for f in (extracted.get("technical_facts") or [])}
    bad: list[str] = []
    suspicious_exclusions: list[str] = []
    for pid, rec in expected.items():
        row = seen[pid]
        cls = str(row.get("classification", "") or "").strip().casefold()
        mapped = [str(x or "").strip() for x in (row.get("fact_ids") or []) if str(x or "").strip()]
        reason = str(row.get("reason", "") or "").strip()
        if cls == "technical":
            if not mapped or any(fid not in fact_ids for fid in mapped):
                bad.append(pid)
        elif cls == "nontechnical":
            if not reason:
                bad.append(pid)
            text = rec["text"]
            if TECHNICAL_PASSAGE_HINT_RE.search(text) and not ADMIN_PASSAGE_HINT_RE.search(text):
                suspicious_exclusions.append(pid)
        else:
            bad.append(pid)
    if bad:
        raise ValueError("Ham kaynak pasaj auditinde teknik fact eşleşmesi/sınıflandırması geçersiz: " + ", ".join(bad[:40]))
    if suspicious_exclusions:
        raise ValueError("Teknik içerik işareti taşıdığı halde teknik-dışı sınıflandırılan kaynak pasajları bulundu: " + ", ".join(suspicious_exclusions[:40]))


def resolve_tarifname_claim_mode(extracted: dict[str, Any], requested_mode: str) -> str:
    """Otomatik modda açık sistem unsurları + yöntem adımları varsa model önerisi bunlardan birini düşüremez."""
    if requested_mode != "BBF'ye göre otomatik belirle":
        return requested_mode
    elements = [x for x in (extracted.get("elements") or []) if str(x.get("name", "") or "").strip()]
    method_steps = [x for x in (extracted.get("method_steps") or []) if str(x.get("text", "") or "").strip()]
    system_name_hits = sum(
        bool(re.search(r"modül|birim|sistem|arayüz|motor|yönetici|alt sistem|cihaz|platform", str(x.get("name", "")), re.I))
        for x in elements
    )
    has_system = bool(extracted.get("has_system_basis")) or len(elements) >= 2 or system_name_hits >= 2
    has_method = bool(extracted.get("has_method_basis")) or len(method_steps) >= 2
    if has_system and has_method:
        return "Sistem ve yöntem"
    if has_method:
        return "Yalnızca yöntem"
    if has_system:
        return "Yalnızca sistem"
    recommended = str(extracted.get("recommended_claim_mode", "") or "").strip()
    if recommended in {"Yalnızca sistem", "Yalnızca yöntem", "Sistem ve yöntem"}:
        return recommended
    raise ValueError("BBF'ye göre otomatik istem yapısı belirlenemedi; açık sistem/yöntem teknik dayanağı bulunamadı.")


_FILENAME_FORBIDDEN_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]+')


def derive_tarifname_output_names(reference: str) -> tuple[str, str]:
    """DP referansını tek kaynak kabul ederek Tarifname/Şekiller çıktı adlarını otomatik üretir."""
    token = re.sub(r"\s+", "_", str(reference or "").strip())
    token = _FILENAME_FORBIDDEN_RE.sub("_", token).strip(" ._")
    if not token:
        raise ValueError("DP referans numarası boş bırakılamaz; çıktı dosya adları bu referanstan otomatik oluşturulur.")
    return f"Tarifname_{token}.docx", f"Şekiller_{token}.docx"


def _coverage_row_map(coverage_rows: list[dict[str, Any]] | None) -> dict[str, dict[str, Any]]:
    rows: dict[str, dict[str, Any]] = {}
    duplicates: list[str] = []
    for row in coverage_rows or []:
        fid = str(row.get("fact_id", "") or "").strip()
        if not fid:
            continue
        if fid in rows:
            duplicates.append(fid)
        rows[fid] = row
    if duplicates:
        raise ValueError("Nihai kaynak kapsam tablosunda tekrarlanan fact_id bulundu: " + ", ".join(sorted(set(duplicates))))
    return rows


def validate_final_source_coverage_chain(
    extracted: dict[str, Any],
    registry: list[dict[str, str]],
    coverage_rows: list[dict[str, Any]] | None,
    final_text: str,
) -> dict[str, int]:
    """Ham pasaj -> technical_fact -> source_coverage_map -> nihai Word zincirini deterministik olarak doğrular."""
    validate_source_passage_audit(extracted, registry)

    facts = {
        str(f.get("id", "") or "").strip(): f
        for f in (extracted.get("technical_facts") or [])
        if str(f.get("id", "") or "").strip()
    }
    if not facts:
        raise ValueError("Nihai ham kaynak zinciri kontrolü için technical_facts envanteri bulunamadı.")

    rows = _coverage_row_map(coverage_rows)
    missing_facts: list[str] = []
    for fid in facts:
        row = rows.get(fid) or {}
        evidence = str(row.get("evidence", "") or "").strip()
        if row.get("covered") is not True or not (row.get("sections") or []) or len(evidence) < 20 or evidence not in final_text:
            missing_facts.append(fid)
    if missing_facts:
        raise ValueError(
            "NİHAİ HAM VERİ ZİNCİRİ başarısız: technical_fact -> nihai Word kanıtı eksik/geçersiz: "
            + ", ".join(sorted(missing_facts))
        )

    source_rows = {
        str(r.get("passage_id", "") or "").strip(): r
        for r in (extracted.get("source_passage_audit") or [])
        if str(r.get("passage_id", "") or "").strip()
    }
    technical_passages = 0
    broken_passages: list[str] = []
    for rec in registry:
        pid = rec["passage_id"]
        row = source_rows.get(pid) or {}
        if str(row.get("classification", "") or "").strip().casefold() != "technical":
            continue
        technical_passages += 1
        mapped = [str(x or "").strip() for x in (row.get("fact_ids") or []) if str(x or "").strip()]
        if not mapped or any(fid not in facts for fid in mapped):
            broken_passages.append(pid)
            continue
        for fid in mapped:
            cov = rows.get(fid) or {}
            evidence = str(cov.get("evidence", "") or "").strip()
            if cov.get("covered") is not True or not (cov.get("sections") or []) or len(evidence) < 20 or evidence not in final_text:
                broken_passages.append(pid)
                break
    if broken_passages:
        raise ValueError(
            "NİHAİ HAM VERİ ZİNCİRİ başarısız: teknik ham pasaj -> fact -> nihai Word zinciri kopuk: "
            + ", ".join(sorted(set(broken_passages)))
        )

    return {
        "raw_passages_total": len(registry),
        "technical_passages": technical_passages,
        "nontechnical_passages": len(registry) - technical_passages,
        "technical_facts": len(facts),
        "covered_facts": len(facts),
    }



_DETAIL_PRIOR_ART_CATEGORIES = {
    "önceki_teknik", "onceki_teknik", "prior_art", "prior art",
    "literatür", "literatur", "literature", "patent_literature",
    "problem", "teknik_problem", "technical_problem", "technical problem",
}
_DETAIL_SECTION_MARKERS = ("buluşun detaylı açıklaması", "buluşun detayli açiklamasi", "detailed description")
_GENERIC_REF_PLACEHOLDER_RE = re.compile(
    r"\b(?:(?:birinci|ikinci|üçüncü|dördüncü|beşinci|altıncı|yedinci|sekizinci|dokuzuncu|onuncu)\s+)?(?:eleman|unsur|parça)\s*\(\s*([A-Za-z0-9_.-]+)\s*\)",
    re.IGNORECASE,
)
_TECH_ACRONYM_RE = re.compile(
    r"(?:\b[A-ZÇĞİÖŞÜ]{3,7}\b|\b[A-Z]{1,6}\d(?:[A-Z0-9]*)(?:\.\d+[A-Z0-9]*)?\b)"
)
_TECH_VALUE_RE = re.compile(
    r"\b\d+(?:[.,]\d+)?(?:\s*[–—-]\s*\d+(?:[.,]\d+)?)?\s*(?:nm|µm|um|mm|cm|km|Hz|kHz|MHz|GHz|W|mW|kW|V|mV|A|mA|°C|%|dB|rpm)\b",
    re.IGNORECASE,
)


def _detail_fact_required(fact: dict[str, Any]) -> bool:
    category = re.sub(r"\s+", "_", str(fact.get("category", "") or "").strip().casefold())
    return category not in _DETAIL_PRIOR_ART_CATEGORIES


def _norm_compare_text(value: str) -> str:
    value = str(value or "").replace("–", "-").replace("—", "-").replace("−", "-")
    value = re.sub(r"\s+", " ", value).strip().casefold()
    value = re.sub(r"\s*-\s*", "-", value)
    return value


def _extract_protected_technical_literals(text: str) -> set[str]:
    literals: set[str] = set()
    raw = str(text or "")
    for rx in (_TECH_ACRONYM_RE, _TECH_VALUE_RE):
        for m in rx.finditer(raw):
            lit = re.sub(r"\s+", " ", m.group(0)).strip()
            if lit:
                literals.add(lit)
    return literals



def validate_detailed_description_fact_coverage(
    extracted: dict[str, Any],
    coverage_rows: list[dict[str, Any]] | None,
    detail_text: str,
    elements: list[dict[str, Any]] | None = None,
) -> dict[str, int]:
    """Taslak aşamasında fact -> Detaylı Açıklama evidence ve fact-literal kapısını çalıştırır."""
    detail_norm = _norm_compare_text(detail_text)
    if len(detail_norm) < 40:
        raise ValueError("DETAYLI AÇIKLAMA fact kapısı başarısız: detaylı açıklama metni boş veya yetersiz.")
    facts = {
        str(f.get("id", "") or "").strip(): f
        for f in (extracted.get("technical_facts") or [])
        if str(f.get("id", "") or "").strip()
    }
    rows = _coverage_row_map(coverage_rows)
    required_ids = {fid for fid, fact in facts.items() if _detail_fact_required(fact)}
    missing: list[str] = []
    checked_literals: set[str] = set()
    literal_missing: list[str] = []
    for fid in sorted(required_ids):
        fact = facts[fid]
        row = rows.get(fid) or {}
        sections = [str(x or "").strip().casefold() for x in (row.get("sections") or [])]
        evidence = str(row.get("evidence", "") or "").strip()
        detail_section_named = any(any(marker in sec for marker in _DETAIL_SECTION_MARKERS) for sec in sections)
        if row.get("covered") is not True or not detail_section_named or len(evidence) < 20 or _norm_compare_text(evidence) not in detail_norm:
            missing.append(fid)
        for literal in _extract_protected_technical_literals(str(fact.get("statement", "") or "")):
            norm_lit = _norm_compare_text(literal)
            if not norm_lit or norm_lit in checked_literals:
                continue
            checked_literals.add(norm_lit)
            if norm_lit not in detail_norm:
                literal_missing.append(literal)
    if missing:
        raise ValueError(
            "DETAYLI AÇIKLAMA fact kapısı başarısız: buluş-teknik technical_fact için Detaylı Açıklama evidence eksik/geçersiz: "
            + ", ".join(missing[:60])
        )
    if literal_missing:
        raise ValueError(
            "DETAYLI AÇIKLAMA fact-literal kapısı başarısız; technical_fact içinde olup detaylı açıklamada bulunmayan değer/kısaltma: "
            + ", ".join(sorted(set(literal_missing), key=str.casefold)[:60])
        )
    valid_numbers = {str(e.get("number", "") or "").strip() for e in (elements or []) if str(e.get("number", "") or "").strip()}
    placeholders = sorted({m.group(0) for m in _GENERIC_REF_PLACEHOLDER_RE.finditer(str(detail_text or "")) if m.group(1) in valid_numbers})
    if placeholders:
        raise ValueError(
            "DETAYLI AÇIKLAMA referans adı normalizasyon kapısı başarısız; gerçek unsur adı yerine geçici/genel ifade kullanılmış: "
            + ", ".join(placeholders[:30])
        )
    return {
        "detail_required_facts": len(required_ids),
        "detail_covered_facts": len(required_ids),
        "detail_fact_literals": len(checked_literals),
    }


def validate_detailed_description_source_transfer(
    extracted: dict[str, Any],
    registry: list[dict[str, str]],
    coverage_rows: list[dict[str, Any]] | None,
    detail_text: str,
    elements: list[dict[str, Any]] | None = None,
) -> dict[str, int]:
    """Buluş-teknik kaynak bilgisinin özellikle DETAYLI AÇIKLAMA içinde tam aktarıldığını doğrular."""
    validate_source_passage_audit(extracted, registry)
    fact_stats = validate_detailed_description_fact_coverage(extracted, coverage_rows, detail_text, elements)
    detail_norm = _norm_compare_text(detail_text)
    if len(detail_norm) < 40:
        raise ValueError("DETAYLI AÇIKLAMA kaynak-transfer kapısı başarısız: detaylı açıklama metni boş veya yetersiz.")

    facts = {
        str(f.get("id", "") or "").strip(): f
        for f in (extracted.get("technical_facts") or [])
        if str(f.get("id", "") or "").strip()
    }
    rows = _coverage_row_map(coverage_rows)
    required_ids = {fid for fid, fact in facts.items() if _detail_fact_required(fact)}
    missing: list[str] = []
    for fid in sorted(required_ids):
        row = rows.get(fid) or {}
        sections = [str(x or "").strip().casefold() for x in (row.get("sections") or [])]
        evidence = str(row.get("evidence", "") or "").strip()
        detail_section_named = any(any(marker in sec for marker in _DETAIL_SECTION_MARKERS) for sec in sections)
        if row.get("covered") is not True or not detail_section_named or len(evidence) < 20 or _norm_compare_text(evidence) not in detail_norm:
            missing.append(fid)
    if missing:
        raise ValueError(
            "DETAYLI AÇIKLAMA kaynak-transfer kapısı başarısız: buluş-teknik technical_fact için Detaylı Açıklama evidence eksik/geçersiz: "
            + ", ".join(missing[:60])
        )

    # Ham pasajlardaki ayırt edici teknik literal/değerler, ilgili fact Detaylı Açıklamaya zorunluysa kaybolamaz.
    registry_by_id = {str(r.get("passage_id", "") or "").strip(): r for r in registry}
    literal_missing: list[str] = []
    checked_literals: set[str] = set()
    for row in (extracted.get("source_passage_audit") or []):
        if str(row.get("classification", "") or "").strip().casefold() != "technical":
            continue
        mapped = [str(x or "").strip() for x in (row.get("fact_ids") or []) if str(x or "").strip()]
        if not any(fid in required_ids for fid in mapped):
            continue
        pid = str(row.get("passage_id", "") or "").strip()
        raw_text = str((registry_by_id.get(pid) or {}).get("text", "") or "")
        for literal in _extract_protected_technical_literals(raw_text):
            norm_lit = _norm_compare_text(literal)
            if not norm_lit or norm_lit in checked_literals:
                continue
            checked_literals.add(norm_lit)
            if norm_lit not in detail_norm:
                literal_missing.append(literal)
    if literal_missing:
        raise ValueError(
            "DETAYLI AÇIKLAMA teknik literal/değer kapısı başarısız; kaynakta olup detaylı açıklamada bulunmayan değer/kısaltma: "
            + ", ".join(sorted(set(literal_missing), key=str.casefold)[:60])
        )

    # Kaynaktaki geçici/genel referans adları nihai gerçek unsur adı belirlenmişse gövdede kalamaz.
    valid_numbers = {str(e.get("number", "") or "").strip() for e in (elements or []) if str(e.get("number", "") or "").strip()}
    placeholders = sorted({m.group(0) for m in _GENERIC_REF_PLACEHOLDER_RE.finditer(str(detail_text or "")) if m.group(1) in valid_numbers})
    if placeholders:
        raise ValueError(
            "DETAYLI AÇIKLAMA referans adı normalizasyon kapısı başarısız; gerçek unsur adı yerine geçici/genel ifade kullanılmış: "
            + ", ".join(placeholders[:30])
        )

    return {
        **fact_stats,
        "detail_protected_literals": len(checked_literals),
    }


def _registry_fingerprint(registry: list[dict[str, str]]) -> str:
    payload = "\n".join(
        f"{str(r.get('passage_id','')).strip()}|{str(r.get('source','')).strip()}|{re.sub(r'\s+', ' ', str(r.get('text','') or '')).strip()}"
        for r in registry
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _draft_fingerprint(final_draft_text: str) -> str:
    normalized = re.sub(r"\s+", " ", str(final_draft_text or "")).strip()
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


def validate_final_raw_source_audit(
    audit: dict[str, Any],
    extracted: dict[str, Any],
    registry: list[dict[str, str]],
    final_draft_text: str,
    *,
    detail_text: str = "",
    expected_audit_nonce: str = "",
) -> dict[str, int]:
    """Gerçek bağımsız ikinci okumayı doğrular; önceki teknik sınıflandırmaya güvenmez."""
    validate_source_passage_audit(extracted, registry)  # birinci kaynak envanteri kapısı ayrı olarak hâlâ zorunlu
    meta = audit.get("audit_meta") or {}
    expected_source_fp = _registry_fingerprint(registry)
    expected_draft_fp = _draft_fingerprint(final_draft_text)
    meta_errors=[]
    if meta.get("audit_mode") != "independent_raw_source_second_read_v2": meta_errors.append("audit_mode")
    if meta.get("independent_second_read") is not True: meta_errors.append("independent_second_read")
    if meta.get("prior_classification_used") is not False: meta_errors.append("prior_classification_used")
    if meta.get("source_coverage_map_used") is not False: meta_errors.append("source_coverage_map_used")
    if expected_audit_nonce and str(meta.get("audit_nonce", "")) != expected_audit_nonce: meta_errors.append("audit_nonce")
    if str(meta.get("source_fingerprint", "")) != expected_source_fp: meta_errors.append("source_fingerprint")
    if str(meta.get("draft_fingerprint", "")) != expected_draft_fp: meta_errors.append("draft_fingerprint")
    if meta_errors:
        raise ValueError("SON HAM KAYNAK İKİNCİ OKUMA bağımsızlık/provenance kapısı başarısız: " + ", ".join(meta_errors))

    expected={str(r.get("passage_id", "") or "").strip(): r for r in registry if str(r.get("passage_id", "") or "").strip()}
    checks={}
    dups=[]
    for row in audit.get("passage_checks") or []:
        rid=str(row.get("passage_id", "") or "").strip()
        if not rid: continue
        if rid in checks: dups.append(rid)
        checks[rid]=row
    missing=sorted(set(expected)-set(checks)); extra=sorted(set(checks)-set(expected))
    if missing or extra or dups:
        parts=[]
        if missing: parts.append("eksik pasaj="+", ".join(missing[:30]))
        if extra: parts.append("tanımsız pasaj="+", ".join(extra[:20]))
        if dups: parts.append("tekrar pasaj="+", ".join(sorted(set(dups))[:20]))
        raise ValueError("SON HAM KAYNAK İKİNCİ OKUMA kapısı eksik: " + "; ".join(parts))

    failed=[]; technical_count=0; nontechnical_count=0; detail_transfer_count=0
    detail_norm = _norm_compare_text(detail_text)
    first_rows = {str(r.get("passage_id", "") or "").strip(): r for r in (extracted.get("source_passage_audit") or []) if str(r.get("passage_id", "") or "").strip()}
    facts = {str(f.get("id", "") or "").strip(): f for f in (extracted.get("technical_facts") or []) if str(f.get("id", "") or "").strip()}
    for rid,row in checks.items():
        classification=str(row.get("classification", "") or "").strip().casefold()
        reason=str(row.get("classification_reason", "") or "").strip()
        source_quote=re.sub(r"\s+", " ", str(row.get("source_quote", "") or "")).strip()
        source_text=re.sub(r"\s+", " ", str(expected[rid].get("text", "") or "")).strip()
        evidence=row.get("evidence") or []
        if isinstance(evidence,str): evidence=[evidence]
        evidence=[str(x or "").strip() for x in evidence if str(x or "").strip()]
        if classification not in {"technical","nontechnical"} or len(reason) < 10:
            failed.append(f"{rid}:classification") ; continue
        first_row = first_rows.get(rid) or {}
        first_classification = str(first_row.get("classification", "") or "").strip().casefold()
        if first_classification not in {"technical", "nontechnical"} or first_classification != classification:
            failed.append(f"{rid}:classification_mismatch")
        mapped_first = [str(x or "").strip() for x in (first_row.get("fact_ids") or []) if str(x or "").strip()]
        expected_detail_required = classification == "technical" and any(_detail_fact_required(facts.get(fid) or {}) for fid in mapped_first)
        min_quote = min(20, len(source_text))
        if len(source_quote) < min_quote or source_quote not in source_text:
            failed.append(f"{rid}:source_quote") ; continue
        if classification == "technical":
            technical_count += 1
            if row.get("covered") is not True:
                failed.append(f"{rid}:covered") ; continue
            if not evidence or any(len(ev) < 20 or ev not in final_draft_text for ev in evidence):
                failed.append(f"{rid}:evidence")
            detail_required = row.get("detail_transfer_required")
            if detail_required not in {True, False}:
                failed.append(f"{rid}:detail_transfer_required")
            elif detail_required is not expected_detail_required:
                failed.append(f"{rid}:detail_transfer_mismatch")
            elif detail_required is True:
                detail_transfer_count += 1
                detail_evidence = str(row.get("detail_evidence", "") or "").strip()
                if len(detail_evidence) < 20 or _norm_compare_text(detail_evidence) not in detail_norm:
                    failed.append(f"{rid}:detail_evidence")
        else:
            nontechnical_count += 1
            if row.get("covered") is not True:
                failed.append(f"{rid}:nontechnical_covered")
            if row.get("detail_transfer_required") not in {False, None}:
                failed.append(f"{rid}:nontechnical_detail_flag")
    if audit.get("all_pass") is not True: failed.append("all_pass")
    if failed:
        raise ValueError("SON HAM KAYNAK İKİNCİ OKUMA başarısız: " + ", ".join(failed[:50]))
    if technical_count == 0:
        raise ValueError("SON HAM KAYNAK İKİNCİ OKUMA başarısız: hiçbir ham pasaj teknik olarak sınıflandırılmadı; kaynak yeniden okunmalıdır.")
    return {"audited_raw_passages":len(expected),"audited_technical_passages":technical_count,"audited_nontechnical_passages":nontechnical_count,"audited_detail_transfer_passages":detail_transfer_count,"independent_raw_second_read":1}
