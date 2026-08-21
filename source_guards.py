from __future__ import annotations

import re
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


def validate_final_raw_source_audit(
    audit: dict[str, Any],
    extracted: dict[str, Any],
    registry: list[dict[str, str]],
    final_draft_text: str,
) -> dict[str, int]:
    """Taslak üretildikten sonra yapılan bağımsız ham-BBF ikinci okumasının eksiksiz olduğunu doğrular."""
    validate_source_passage_audit(extracted, registry)
    source_rows = {
        str(r.get("passage_id", "") or "").strip(): r
        for r in (extracted.get("source_passage_audit") or [])
        if str(r.get("passage_id", "") or "").strip()
    }
    expected_passages = {
        rec["passage_id"]
        for rec in registry
        if str((source_rows.get(rec["passage_id"]) or {}).get("classification", "") or "").strip().casefold() == "technical"
    }
    expected_facts = {
        str(f.get("id", "") or "").strip()
        for f in (extracted.get("technical_facts") or [])
        if str(f.get("id", "") or "").strip()
    }

    def index_rows(rows: list[dict[str, Any]], key: str) -> tuple[dict[str, dict[str, Any]], list[str]]:
        out: dict[str, dict[str, Any]] = {}
        dups: list[str] = []
        for row in rows:
            rid = str(row.get(key, "") or "").strip()
            if not rid:
                continue
            if rid in out:
                dups.append(rid)
            out[rid] = row
        return out, dups

    passage_checks, pdups = index_rows(audit.get("passage_checks") or [], "passage_id")
    fact_checks, fdups = index_rows(audit.get("fact_checks") or [], "fact_id")
    missing_p = sorted(expected_passages - set(passage_checks))
    extra_p = sorted(set(passage_checks) - expected_passages)
    missing_f = sorted(expected_facts - set(fact_checks))
    extra_f = sorted(set(fact_checks) - expected_facts)
    if pdups or fdups or missing_p or extra_p or missing_f or extra_f:
        details: list[str] = []
        if missing_p: details.append("eksik teknik pasaj=" + ", ".join(missing_p[:30]))
        if extra_p: details.append("tanımsız pasaj=" + ", ".join(extra_p[:20]))
        if pdups: details.append("tekrar pasaj=" + ", ".join(sorted(set(pdups))[:20]))
        if missing_f: details.append("eksik fact=" + ", ".join(missing_f[:30]))
        if extra_f: details.append("tanımsız fact=" + ", ".join(extra_f[:20]))
        if fdups: details.append("tekrar fact=" + ", ".join(sorted(set(fdups))[:20]))
        raise ValueError("SON HAM KAYNAK İKİNCİ OKUMA kapısı eksik: " + "; ".join(details))

    failed: list[str] = []
    for label, rows in (("P", passage_checks), ("T", fact_checks)):
        for rid, row in rows.items():
            evidence_items = row.get("evidence") or []
            if isinstance(evidence_items, str):
                evidence_items = [evidence_items]
            evidence_items = [str(x or "").strip() for x in evidence_items if str(x or "").strip()]
            if row.get("covered") is not True:
                failed.append(f"{label}:{rid}")
                continue
            if not evidence_items or any(len(ev) < 20 or ev not in final_draft_text for ev in evidence_items):
                failed.append(f"{label}:{rid}")
    if audit.get("all_pass") is not True:
        failed.append("all_pass")
    if failed:
        missing_notes = []
        for row in [*(audit.get("passage_checks") or []), *(audit.get("fact_checks") or [])]:
            if row.get("covered") is not True and str(row.get("missing_detail", "") or "").strip():
                rid = row.get("passage_id") or row.get("fact_id") or "?"
                missing_notes.append(f"{rid}: {str(row.get('missing_detail')).strip()}")
        suffix = (" | " + " ; ".join(missing_notes[:12])) if missing_notes else ""
        raise ValueError("SON HAM KAYNAK İKİNCİ OKUMA başarısız: " + ", ".join(failed[:50]) + suffix)

    return {"audited_technical_passages": len(expected_passages), "audited_technical_facts": len(expected_facts)}
