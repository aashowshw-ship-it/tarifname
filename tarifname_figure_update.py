from __future__ import annotations

import base64
import io
import json
import os
import re
from pathlib import Path
from typing import Any

from PIL import Image


def parse_figure_number(label: str) -> int | None:
    """Return an explicit Şekil/Figure number; generic actions are not auto-editable."""
    text = str(label or "").strip()
    m = re.search(r"(?:şekil|figure)\s*[-:#.]?\s*(\d+)", text, flags=re.I)
    return int(m.group(1)) if m else None


def _extract_json(text: str) -> dict[str, Any]:
    text = str(text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start, end = text.find("{"), text.rfind("}")
        if start >= 0 and end > start:
            return json.loads(text[start : end + 1])
        raise ValueError("Şekil doğrulama yanıtı geçerli JSON olarak okunamadı.")


def _image_content(data: bytes, mime: str) -> dict[str, Any]:
    b64 = base64.b64encode(data).decode("ascii")
    return {"type": "input_image", "image_url": f"data:{mime or 'image/png'};base64,{b64}", "detail": "high"}


def _extract_image_generation_result(response: Any) -> bytes:
    for item in getattr(response, "output", []) or []:
        item_type = item.get("type") if isinstance(item, dict) else getattr(item, "type", "")
        if item_type != "image_generation_call":
            continue
        result = item.get("result") if isinstance(item, dict) else getattr(item, "result", None)
        if isinstance(result, list) and result:
            result = result[-1]
        if isinstance(result, str) and result.strip():
            try:
                return base64.b64decode(result)
            except Exception as exc:  # pragma: no cover - API response shape guard
                raise ValueError("Şekil revizyon çıktısı base64 görsel olarak çözülemedi.") from exc
    raise ValueError("Şekil revizyon çağrısı görsel çıktı üretmedi.")


def _valid_image(data: bytes) -> bool:
    try:
        with Image.open(io.BytesIO(data)) as im:
            return im.width > 0 and im.height > 0
    except Exception:
        return False


def edit_figure_for_customer_actions(
    image_data: bytes,
    mime: str,
    figure_number: int,
    actions: list[dict[str, Any]],
    *,
    model: str,
    client: Any | None = None,
) -> bytes:
    """Apply only source-backed, explicitly requested small figure changes."""
    if client is None:
        from openai import OpenAI
        client = OpenAI(api_key=os.getenv("OPENAI_API_KEY", "").strip())
    prompt = f"""Bu görsel bir patent şeklidir. Özgün müşteri çizimi teknik geometri bakımından bağlayıcıdır.
ŞEKİL {figure_number} üzerinde YALNIZ aşağıdaki kaynak-destekli müşteri revizyonlarını uygula. Bunların dışında hiçbir teknik değişiklik yapma.

REVİZYONLAR:
{json.dumps(actions, ensure_ascii=False, indent=2)}

KESİN KURALLAR:
- Yalnız `edit_instructions` alanında açıkça istenen değişiklikleri yap.
- Mevcut parça geometrisini, perspektifi, kesitleri, boyut ilişkilerini, mevcut referans işaretlerini ve teknik bağlantıları değiştirme.
- İstenen bir kablo/bağlantı çizgisi ise yalnız belirtilen iki mevcut unsur arasında sade siyah bağlantı çizgisi ekle; yeni cihaz/parça uydurma.
- İstenen düğme/sekme/etiket ise yalnız mevcut monitör/ekran üzerinde sade ve okunabilir bir arayüz işareti ekle; yeni işlev uydurma.
- İstenen referans/ok düzeltmesi ise yalnız ilgili numara ve kılavuz çizgisi/ok katmanını değiştir.
- Mevcut yazı ve formülleri mümkün olduğunca aynen koru.
- Siyah-beyaz patent çizim tarzını ve özgün en-boy oranını koru.
- Çıktıda sadece revize patent şekli bulunsun; açıklama, lejant, başlık veya yorum ekleme.
"""
    content = [{"type": "input_text", "text": prompt}, _image_content(image_data, mime)]
    common = {"model": model, "input": [{"role": "user", "content": content}]}
    try:
        response = client.responses.create(
            **common,
            tools=[{"type": "image_generation", "quality": "high", "input_fidelity": "high"}],
            tool_choice="required",
        )
    except Exception:
        response = client.responses.create(
            **common,
            tools=[{"type": "image_generation", "quality": "high"}],
            tool_choice="required",
        )
    data = _extract_image_generation_result(response)
    if not _valid_image(data):
        raise ValueError(f"ŞEKİL {figure_number}: otomatik revizyon geçerli bir görsel üretmedi.")
    return data


def verify_customer_figure_edit(
    original_data: bytes,
    edited_data: bytes,
    mime: str,
    figure_number: int,
    actions: list[dict[str, Any]],
    *,
    model: str,
    client: Any | None = None,
) -> dict[str, Any]:
    """Reject edits that alter anything beyond the requested, source-backed changes."""
    if client is None:
        from openai import OpenAI
        client = OpenAI(api_key=os.getenv("OPENAI_API_KEY", "").strip())
    prompt = f"""İki patent şekli veriliyor. Birinci görsel özgün ŞEKİL {figure_number}, ikinci görsel revize adaydır.
Adayı sıkı biçimde karşılaştır. Görsel üretme; yalnız JSON döndür.

İZİN VERİLEN REVİZYONLAR:
{json.dumps(actions, ensure_ascii=False, indent=2)}

KABUL KRİTERLERİ:
1. İstenen değişikliklerin tamamı görünür ve doğru uygulanmış olmalı.
2. İstenen değişiklikler dışında mekanik/elektronik geometri, parça biçimi, perspektif, kesit/tarama, formül, mevcut referans veya teknik kurgu değişmemiş olmalı.
3. Yeni, istenmemiş teknik unsur, bağlantı, metin veya referans eklenmemiş olmalı.
4. Eklenen kablo/bağlantı veya düğme/etiket, yalnız izin verilen mevcut unsurlara bağlanmalı ve teknik anlamı bozmamalı.
5. Patent çizimi siyah-beyaz ve okunabilir kalmalı.

JSON ŞEMASI:
{{
  "requested_changes_complete": true,
  "geometry_preserved_except_requested": true,
  "technical_consistency": true,
  "unexpected_changes": [],
  "missing_or_wrong_changes": [],
  "readable": true,
  "confidence": 0.95,
  "notes": ""
}}
"""
    response = client.responses.create(
        model=model,
        input=[{
            "role": "user",
            "content": [
                {"type": "input_text", "text": prompt},
                _image_content(original_data, mime),
                _image_content(edited_data, "image/png"),
            ],
        }],
    )
    return _extract_json(response.output_text)


def prepare_customer_figure_edits(
    figures: list[dict[str, Any]],
    figure_actions: list[dict[str, Any]],
    *,
    model: str,
    confidence_threshold: float = 0.86,
    client: Any | None = None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[str]]:
    """Edit only explicitly numbered figures whose every action passed the source/new-matter gates."""
    prepared = [dict(x) for x in figures]
    reports: list[dict[str, Any]] = []
    unresolved: list[str] = []

    grouped: dict[int, list[dict[str, Any]]] = {}
    for action in figure_actions or []:
        if not bool(action.get("safe_auto_edit")):
            unresolved.append(
                f"{action.get('figure', 'Şekil')}: otomatik şekil revizyonu güvenli işaretlenmedi; yalnız aksiyon önerisi olarak bırakıldı."
            )
            continue
        figure_number = parse_figure_number(str(action.get("figure", "")))
        if figure_number is None:
            unresolved.append(f"{action.get('figure', 'Genel')}: hedef şekil numarası açık olmadığı için otomatik revizyon yapılmadı.")
            continue
        grouped.setdefault(figure_number, []).append(action)

    for figure_number, actions in sorted(grouped.items()):
        if figure_number < 1 or figure_number > len(prepared):
            unresolved.append(f"ŞEKİL {figure_number}: yüklenen şekil setinde bu numaraya karşılık gelen görsel bulunamadı.")
            continue
        item = prepared[figure_number - 1]
        original = bytes(item["data"])
        mime = str(item.get("mime") or "image/png")
        try:
            edited = edit_figure_for_customer_actions(
                original, mime, figure_number, actions, model=model, client=client
            )
            verification = verify_customer_figure_edit(
                original, edited, mime, figure_number, actions, model=model, client=client
            )
            try:
                confidence = float(verification.get("confidence", 0.0) or 0.0)
            except (TypeError, ValueError):
                confidence = 0.0
            accepted = (
                bool(verification.get("requested_changes_complete"))
                and bool(verification.get("geometry_preserved_except_requested"))
                and bool(verification.get("technical_consistency"))
                and bool(verification.get("readable"))
                and not any(str(x).strip() for x in (verification.get("unexpected_changes") or []))
                and not any(str(x).strip() for x in (verification.get("missing_or_wrong_changes") or []))
                and confidence >= confidence_threshold
            )
            report = {"figure_number": figure_number, "actions": actions, "verification": verification}
            if not accepted:
                report["status"] = "unresolved"
                reports.append(report)
                unresolved.append(f"ŞEKİL {figure_number}: revize görsel ikinci doğrulamayı geçemedi; özgün şekil korunuyor.")
                continue
            item["data"] = edited
            item["mime"] = "image/png"
            item["name"] = f"{Path(str(item.get('name') or f'figure_{figure_number}')).stem}_revize.png"
            report["status"] = "corrected"
            reports.append(report)
        except Exception as exc:
            reports.append({"figure_number": figure_number, "actions": actions, "status": "unresolved", "error": str(exc)})
            unresolved.append(f"ŞEKİL {figure_number}: otomatik şekil revizyonu uygulanamadı ({exc}).")

    return prepared, reports, unresolved
