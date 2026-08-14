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
