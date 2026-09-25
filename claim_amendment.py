from __future__ import annotations

import re
from collections import Counter
from copy import deepcopy
from difflib import SequenceMatcher
from typing import Any, Iterable


def _norm(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _tokens(text: str) -> list[str]:
    """Lexical + punctuation tokens used only for amendment minimality checks."""
    return re.findall(r"\w+|[^\w\s]", str(text or ""), flags=re.UNICODE)


def _is_token_subsequence(old_text: str, new_text: str) -> bool:
    old = _tokens(old_text)
    new = _tokens(new_text)
    if not old:
        return False
    j = 0
    for token in new:
        if j < len(old) and token == old[j]:
            j += 1
    return j == len(old)


def _old_token_multiset_preserved(old_text: str, new_text: str) -> bool:
    """True when every original token still exists in the proposal, even if reordered.

    Reordering existing claim words is a common hidden full-phrase rewrite. If all old tokens
    survive but the edit plan contains deletion/reinsertion, insertion at the proper anchor is
    structurally preferable and the rewrite must be blocked.
    """
    old_counts = Counter(_tokens(old_text))
    new_counts = Counter(_tokens(new_text))
    return bool(old_counts) and all(new_counts[token] >= count for token, count in old_counts.items())


def _changed_ratio(old_text: str, new_text: str) -> float:
    old = str(old_text or "")
    new = str(new_text or "")
    if not old:
        return 0.0
    matcher = SequenceMatcher(a=old, b=new, autojunk=False)
    preserved = sum(block.size for block in matcher.get_matching_blocks())
    return max(0.0, min(1.0, 1.0 - (preserved / max(1, len(old)))))


def build_atomic_operations(old_text: str, new_text: str) -> list[dict[str, Any]]:
    """Return deterministic character-level atomic edit operations.

    AI may propose the legal/technical old_text→new_text target. The OOXML layer never
    trusts a model-produced redline. This function derives the actual INSERT/DELETE/REPLACE
    plan deterministically from the approved pair.
    """
    old = str(old_text or "")
    new = str(new_text or "")
    matcher = SequenceMatcher(a=old, b=new, autojunk=False)
    out: list[dict[str, Any]] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        row: dict[str, Any] = {
            "op": tag,
            "old_start": i1,
            "old_end": i2,
            "old_text": old[i1:i2],
            "new_text": new[j1:j2],
        }
        if tag == "insert":
            row["op"] = "insert"
        elif tag == "delete":
            row["op"] = "delete"
        else:
            row["op"] = "replace"
        out.append(row)
    return out


def _op_has_deletion(op: dict[str, Any]) -> bool:
    return str(op.get("op", "")).lower() in {"delete", "replace"} and bool(str(op.get("old_text", "")))


def _op_has_insertion(op: dict[str, Any]) -> bool:
    return str(op.get("op", "")).lower() in {"insert", "replace"} and bool(str(op.get("new_text", "")))


def _deleted_token_count(ops: Iterable[dict[str, Any]]) -> int:
    return sum(len(_tokens(str(op.get("old_text", "")))) for op in ops if _op_has_deletion(op))


def _largest_deleted_span_tokens(ops: Iterable[dict[str, Any]]) -> int:
    return max((len(_tokens(str(op.get("old_text", "")))) for op in ops if _op_has_deletion(op)), default=0)


def enrich_amendment_atomic_plan(amendment: dict[str, Any]) -> dict[str, Any]:
    """Attach the deterministic atomic plan and structural receipts in-place."""
    old_text = str(amendment.get("old_text", ""))
    new_text = str(amendment.get("new_text", ""))
    ops = build_atomic_operations(old_text, new_text)
    amendment["atomic_operations"] = ops
    has_delete = any(_op_has_deletion(x) for x in ops)
    has_insert = any(_op_has_insertion(x) for x in ops)
    if not amendment.get("edit_mode"):
        if has_delete:
            amendment["edit_mode"] = "replace_minimal" if has_insert else "delete_minimal"
        else:
            amendment["edit_mode"] = "insert_only"
    amendment["structural_change_ratio"] = round(_changed_ratio(old_text, new_text), 4)
    amendment["original_tokens_preserved_in_order"] = _is_token_subsequence(old_text, new_text)
    return amendment


def enrich_analysis_atomic_plans(analysis: dict[str, Any]) -> dict[str, Any]:
    for item in analysis.get("amendments") or []:
        enrich_amendment_atomic_plan(item)
    for cand in analysis.get("amendment_candidates") or []:
        for item in cand.get("amendments") or []:
            enrich_amendment_atomic_plan(item)
    return analysis


def _direct_support_valid(item: dict[str, Any], normalized_spec: str) -> bool:
    rows = item.get("direct_support") or []
    if str(item.get("change_type", "technical") or "technical").lower() == "technical" and not rows:
        return False
    new_text = _norm(item.get("new_text", "")).casefold()
    for row in rows:
        feature = _norm(row.get("added_feature", ""))
        basis = _norm(row.get("basis_quote", ""))
        if not feature or not basis or basis not in normalized_spec:
            return False
        # Added feature may be a grammaticalized form of the basis, but it must be visibly
        # present in the proposed claim wording; this prevents free-standing support rows.
        if feature.casefold() not in new_text:
            return False
    return True


def validate_atomic_amendment(
    item: dict[str, Any],
    spec_text: str,
    *,
    known_objection_ids: Iterable[str] = (),
    require_receipts: bool = False,
) -> dict[str, bool]:
    """Hard deterministic amendment gate.

    Core invariant: for technical amendments the default operation is INSERT. Existing
    claim wording is not deleted/rephrased merely for elegance. DELETE/REPLACE requires
    an explicit necessity flag and concrete reason tied to an examiner objection.
    """
    normalized_spec = _norm(spec_text)
    old_text = _norm(item.get("old_text", ""))
    new_text = _norm(item.get("new_text", ""))
    if not old_text or not new_text:
        raise ValueError("İstem revizyonu atomik kapısı: old_text/new_text boş bırakılamaz.")
    if old_text not in normalized_spec:
        raise ValueError("İstem revizyonu atomik kapısı: old_text as-filed tarifnamede/istemde birebir bulunmalıdır.")

    enrich_amendment_atomic_plan(item)
    ops = item.get("atomic_operations") or []
    edit_mode = str(item.get("edit_mode", "")).strip().lower()
    if edit_mode not in {"insert_only", "replace_minimal", "delete_minimal"}:
        raise ValueError("İstem revizyonu atomik kapısı: edit_mode insert_only/replace_minimal/delete_minimal olmalıdır.")

    change_type = str(item.get("change_type", "technical") or "technical").strip().lower()
    has_delete = any(_op_has_deletion(x) for x in ops)
    has_insert = any(_op_has_insertion(x) for x in ops)
    original_preserved = bool(item.get("original_tokens_preserved_in_order"))

    if change_type == "technical":
        if not _direct_support_valid(item, normalized_spec):
            raise ValueError("İstem revizyonu atomik kapısı: her teknik ek özellik için doğrudan as-filed tarifname dayanağı zorunludur.")
        if not has_delete and edit_mode != "insert_only":
            raise ValueError(
                "İstem revizyonu insertion-first kapısı: mevcut kelime silinmiyorsa edit_mode yalnız insert_only olabilir. "
                "Ekleme yeterliyken replace/delete modu kullanılamaz."
            )
        if has_delete:
            if not bool(item.get("deletion_required")):
                raise ValueError(
                    "İstem revizyonu insertion-first kapısı: teknik revizyonda mevcut kelime/ibare silinmiş. "
                    "Silme zorunlu değilse yalnız ekleme yapılmalıdır."
                )
            deletion_reason = _norm(item.get("deletion_reason", ""))
            if len(deletion_reason) < 20:
                raise ValueError("İstem revizyonu insertion-first kapısı: zorunlu silme/değiştirme için somut deletion_reason gerekir.")
            low_reason = deletion_reason.casefold()
            if any(x in low_reason for x in ("daha iyi ifade", "daha akıcı", "daha güzel", "üslup", "stil amacı")):
                raise ValueError("İstem revizyonu insertion-first kapısı: üslup/akıcılık mevcut istem kelimesini silme gerekçesi olamaz.")
            deleted_tokens = _deleted_token_count(ops)
            largest_deleted = _largest_deleted_span_tokens(ops)
            old_token_count = max(1, len(_tokens(old_text)))
            # Technical deletion must stay at word/short-phrase granularity. This is an
            # anti-rewrite guard, not a semantic substitute for the independent auditor.
            if largest_deleted > 12 or (deleted_tokens > 8 and deleted_tokens / old_token_count > 0.30):
                raise ValueError(
                    "İstem revizyonu kelime/ibare kapısı: zorunlu olduğu ileri sürülen silme yerel kelime/ibare sınırını aşıyor. "
                    "Tüm cümleyi/parçayı yeniden kurmak yerine yalnız sorunlu kelime veya kısa ibare değiştirilmelidir."
                )
        if edit_mode == "insert_only" and has_delete:
            raise ValueError("İstem revizyonu insertion-first kapısı: insert_only planında silme/değiştirme bulunamaz.")
        # Strongest deterministic signal of an unnecessary rewrite: all original tokens
        # can remain in the same order in the proposed wording, yet the plan deletes some.
        if original_preserved and has_delete:
            raise ValueError(
                "İstem revizyonu insertion-first kapısı: orijinal istem kelimeleri aynı sırada korunabildiği halde "
                "silme/replace üretilmiş. Mevcut metni koruyup yalnız gerekli kelime/ibareyi ekleyin."
            )
        if has_delete and _old_token_multiset_preserved(old_text, new_text):
            raise ValueError(
                "İstem revizyonu yeniden-sıralama kapısı: mevcut istem kelimelerinin tamamı öneride hâlâ bulunduğu halde "
                "kelimeler silinip başka sırada yeniden eklenmiş. Mevcut sıralamayı koruyup yalnız gerekli eklemeyi yapın."
            )
        if float(item.get("structural_change_ratio", 0.0) or 0.0) > 0.45 and has_delete:
            raise ValueError(
                "İstem revizyonu yeniden-yazım kapısı: mevcut ifade gereğinden geniş ölçüde yeniden kurulmuş. "
                "Yalnız zorunlu kelime/ibare düzeyinde değişiklik yapılabilir."
            )

    objection_ids = {str(x).strip() for x in known_objection_ids if str(x).strip()}
    linked = {str(x).strip() for x in (item.get("addresses_objection_ids") or []) if str(x).strip()}
    if objection_ids:
        if not linked:
            raise ValueError("İstem revizyonu itiraz-eşleme kapısı: teknik değişiklik en az bir uzman itiraz kimliğine bağlanmalıdır.")
        unknown = linked - objection_ids
        if unknown:
            raise ValueError("İstem revizyonu itiraz-eşleme kapısı: tanımsız itiraz kimliği kullanılmış: " + ", ".join(sorted(unknown)))

    if require_receipts:
        for key, label in (
            ("clarity_effect", "açıklık etkisi"),
            ("scope_effect", "kapsam etkisi"),
            ("prior_art_effect", "önceki teknik karşısındaki etkisi"),
            ("remaining_risk", "kalan risk"),
        ):
            if not _norm(item.get(key, "")):
                raise ValueError(f"İstem revizyonu analiz kapısı: {label} boş bırakılamaz.")

    return {
        "direct_basis": _direct_support_valid(item, normalized_spec),
        "insertion_first": not has_delete or bool(item.get("deletion_required")),
        "minimality": not (original_preserved and has_delete) and not (float(item.get("structural_change_ratio", 0.0) or 0.0) > 0.45 and has_delete),
        "objection_linked": (not objection_ids) or bool(linked),
        "has_deletion": has_delete,
        "has_insertion": has_insert,
    }


def validate_amendment_candidates(analysis: dict[str, Any], spec_text: str) -> None:
    """Validate background candidates and selected-candidate consistency when present."""
    candidates = analysis.get("amendment_candidates") or []
    if not candidates:
        return
    selected_id = str(analysis.get("selected_candidate_id", "")).strip()
    if not selected_id:
        raise ValueError("İstem revizyonu aday kapısı: selected_candidate_id zorunludur.")
    ids: list[str] = []
    objection_ids = [str(x.get("id", "")).strip() for x in (analysis.get("examiner_objections") or []) if str(x.get("id", "")).strip()]
    for candidate in candidates:
        cid = str(candidate.get("candidate_id", "")).strip()
        if not cid or cid in ids:
            raise ValueError("İstem revizyonu aday kapısı: her adayın benzersiz candidate_id değeri olmalıdır.")
        ids.append(cid)
        if not candidate.get("amendments"):
            raise ValueError(f"İstem revizyonu aday kapısı: {cid} değişiklik içermiyor.")
        for amendment in candidate.get("amendments") or []:
            validate_atomic_amendment(
                amendment, spec_text, known_objection_ids=objection_ids, require_receipts=True
            )
    if selected_id not in ids:
        raise ValueError("İstem revizyonu aday kapısı: seçilen aday aday listesinde bulunamadı.")
    selected = next(x for x in candidates if str(x.get("candidate_id", "")).strip() == selected_id)
    selected_amendments = selected.get("amendments") or []
    top = analysis.get("amendments") or []
    # Compare the operational target, not commentary fields.
    def signature(rows: list[dict[str, Any]]) -> list[tuple[str, str, str]]:
        return [
            (str(x.get("claim_number", "")).strip(), _norm(x.get("old_text", "")), _norm(x.get("new_text", "")))
            for x in rows
        ]
    if signature(selected_amendments) != signature(top):
        raise ValueError("İstem revizyonu aday kapısı: top-level amendments seçilmiş adayla birebir aynı değildir.")


def structural_candidate_score(candidate: dict[str, Any], spec_text: str, known_objection_ids: Iterable[str]) -> float:
    """Deterministic structural score; never substitutes for the semantic auditor."""
    score = 0.0
    amendments = deepcopy(candidate.get("amendments") or [])
    if not amendments:
        return -999.0
    try:
        for item in amendments:
            receipt = validate_atomic_amendment(item, spec_text, known_objection_ids=known_objection_ids)
            score += 25.0 if receipt["direct_basis"] else 0.0
            score += 15.0 if receipt["objection_linked"] else 0.0
            score += 20.0 if not receipt["has_deletion"] else 0.0
            score += 10.0 * (1.0 - float(item.get("structural_change_ratio", 0.0) or 0.0))
        score /= max(1, len(amendments))
    except Exception:
        return -999.0
    return round(score, 3)


def validate_selected_candidate_is_structurally_best(analysis: dict[str, Any], spec_text: str) -> None:
    """Reject only a clearly structurally dominated selected candidate.

    Semantic priority (does it actually cure the examiner objection?) belongs to the independent
    Amendment Auditor. Deterministic code blocks a selected candidate only when another candidate
    targets the same objection set and preserves strictly more original wording with no greater
    structural intervention.
    """
    candidates = analysis.get("amendment_candidates") or []
    if len(candidates) < 2:
        return
    selected_id = str(analysis.get("selected_candidate_id", "")).strip()
    selected = next((c for c in candidates if str(c.get("candidate_id", "")).strip() == selected_id), None)
    if not selected:
        return

    def metrics(candidate: dict[str, Any]) -> tuple[frozenset[str], int, float]:
        targets = frozenset(str(x).strip() for x in (candidate.get("addresses_objection_ids") or []) if str(x).strip())
        rows = deepcopy(candidate.get("amendments") or [])
        deletions = 0
        ratio = 0.0
        for row in rows:
            enrich_amendment_atomic_plan(row)
            deletions += _deleted_token_count(row.get("atomic_operations") or [])
            ratio += float(row.get("structural_change_ratio", 0.0) or 0.0)
        return targets, deletions, ratio

    sel_targets, sel_del, sel_ratio = metrics(selected)
    for other in candidates:
        oid = str(other.get("candidate_id", "")).strip()
        if oid == selected_id:
            continue
        targets, deletions, ratio = metrics(other)
        if targets == sel_targets and deletions < sel_del and ratio <= sel_ratio + 1e-9:
            raise ValueError(
                "İstem revizyonu aday kapısı: seçilen aday aynı uzman itirazlarını hedefleyen başka bir aday tarafından "
                f"minimum müdahale bakımından açıkça domine ediliyor ({oid}). Seçim yeniden değerlendirilmelidir."
            )


def validate_amendment_audit(audit: dict[str, Any], analysis: dict[str, Any]) -> None:
    """Fail-closed semantic auditor receipt validation."""
    if not isinstance(audit, dict):
        raise ValueError("İstem revizyonu bağımsız denetimi JSON nesnesi olmalıdır.")
    items = audit.get("items") or []
    amendments = analysis.get("amendments") or []
    if len(items) != len(amendments):
        raise ValueError("İstem revizyonu bağımsız denetimi: denetlenen değişiklik sayısı öneriyle uyuşmuyor.")
    blocking = [str(x).strip() for x in (audit.get("blocking_reasons") or []) if str(x).strip()]
    candidates = analysis.get("amendment_candidates") or []
    if len(candidates) > 1:
        if not bool(audit.get("selected_candidate_confirmed")):
            raise ValueError("İstem revizyonu bağımsız denetimi: seçilen aday diğer güvenli adaylara karşı doğrulanmadı.")
        comparisons = audit.get("candidate_comparison") or []
        compared_ids = {str(x.get("candidate_id", "")).strip() for x in comparisons if isinstance(x, dict)}
        expected_ids = {str(x.get("candidate_id", "")).strip() for x in candidates}
        if not expected_ids.issubset(compared_ids):
            raise ValueError("İstem revizyonu bağımsız denetimi: aday karşılaştırması tüm adayları kapsamıyor.")
        if not _norm(audit.get("selection_basis", "")):
            raise ValueError("İstem revizyonu bağımsız denetimi: aday seçim gerekçesi boş bırakılamaz.")
    if not bool(audit.get("overall_pass")) or blocking:
        raise ValueError("İstem revizyonu bağımsız denetimi başarısız: " + ("; ".join(blocking) or "bloklayıcı bulgu var"))
    for idx, row in enumerate(items, 1):
        if not bool(row.get("direct_basis_pass")):
            raise ValueError(f"İstem revizyonu bağımsız denetimi: {idx}. değişiklikte doğrudan dayanak PASS değil.")
        if str(row.get("addresses_examiner_issue", "")).strip().lower() != "yes":
            raise ValueError(f"İstem revizyonu bağımsız denetimi: {idx}. değişiklik uzman itirazını yeterli biçimde karşılamıyor.")
        if not bool(row.get("minimal_edit_pass")) or bool(row.get("unnecessary_deletion")):
            raise ValueError(f"İstem revizyonu bağımsız denetimi: {idx}. değişiklik minimum-fark/insertion-first kuralını geçmiyor.")
        if bool(row.get("new_ambiguity")):
            raise ValueError(f"İstem revizyonu bağımsız denetimi: {idx}. değişiklik yeni açıklık/belirsizlik sorunu yaratıyor.")
        if str(row.get("scope_effect", "")).strip().lower() == "broadens":
            raise ValueError(f"İstem revizyonu bağımsız denetimi: {idx}. değişiklik kapsamı genişletiyor.")
        for key in ("clarity_effect", "scope_effect", "prior_art_effect", "remaining_risk"):
            if not _norm(row.get(key, "")):
                raise ValueError(f"İstem revizyonu bağımsız denetimi: {idx}. değişiklikte {key} boş bırakılamaz.")


def amendment_ui_rows(analysis: dict[str, Any]) -> list[dict[str, Any]]:
    """Compact normalized rows for the Streamlit review card."""
    rows: list[dict[str, Any]] = []
    for item in analysis.get("amendments") or []:
        rows.append({
            "claim_number": str(item.get("claim_number", "")),
            "reason": _norm(item.get("reason", "")),
            "old_text": str(item.get("old_text", "")),
            "new_text": str(item.get("new_text", "")),
            "edit_mode": str(item.get("edit_mode", "")),
            "addresses": list(item.get("addresses_objection_ids") or []),
            "clarity_effect": _norm(item.get("clarity_effect", "")),
            "scope_effect": _norm(item.get("scope_effect", "")),
            "prior_art_effect": _norm(item.get("prior_art_effect", "")),
            "remaining_risk": _norm(item.get("remaining_risk", "")),
            "direct_support": list(item.get("direct_support") or []),
            "atomic_operations": list(item.get("atomic_operations") or []),
        })
    return rows
