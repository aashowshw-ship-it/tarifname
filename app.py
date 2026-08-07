from __future__ import annotations

import base64
import io
import json
import os
import re
import subprocess
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from copy import deepcopy
from dataclasses import dataclass
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import unquote

import streamlit as st
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from openai import OpenAI
from PIL import Image

try:
    import fitz  # PyMuPDF: PDF içindeki şekilleri çıkarmak için
except ImportError:  # pragma: no cover - bağımlılık Render üzerinde requirements ile kurulur
    fitz = None
from pypdf import PdfReader

from rules import APP_VERSION, RULESET_VERSION, ARASTIRMA_RULES, GORUS_RULES, TARIFNAME_RULES

BASE_DIR = Path(__file__).resolve().parent
TARIFNAME_TEMPLATE = BASE_DIR / "Tarifname_181176_template.docx"
GORUS_TEMPLATE = BASE_DIR / "Gorus_metni_696809_template.docx"
ARASTIRMA_TEMPLATE = BASE_DIR / "On_Arastirma_Raporu_181612_template.docx"
MODEL = os.getenv("OPENAI_MODEL", "gpt-5.6")
MAX_TEXT_PER_FILE = int(os.getenv("MAX_TEXT_PER_FILE", "180000"))
MAX_TOTAL_TEXT = int(os.getenv("MAX_TOTAL_TEXT", "700000"))


# -----------------------------------------------------------------------------
# GENEL KURALLAR
# -----------------------------------------------------------------------------
# Kurallar tek bir kaynaktan (rules.py) yüklenir. Böylece arayüz ve üretim akışları
# arasında kural farkı oluşmaz.


@dataclass
class UploadedAsset:
    name: str
    data: bytes
    mime: str = "application/octet-stream"


# -----------------------------------------------------------------------------
# API / DOSYA YARDIMCILARI
# -----------------------------------------------------------------------------
def get_client() -> OpenAI:
    key = os.getenv("OPENAI_API_KEY", "").strip()
    if not key:
        raise RuntimeError(
            "OPENAI_API_KEY tanımlı değil. Render > servis > Environment bölümünde OPENAI_API_KEY ekleyin."
        )
    return OpenAI(api_key=key)


def extract_json(text: str) -> dict[str, Any]:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start, end = text.find("{"), text.rfind("}")
        if start >= 0 and end > start:
            return json.loads(text[start : end + 1])
        raise ValueError("Yapay zekâ yanıtı geçerli JSON olarak okunamadı.")


def image_content(asset: UploadedAsset) -> dict[str, Any]:
    b64 = base64.b64encode(asset.data).decode("ascii")
    mime = asset.mime or "image/png"
    return {"type": "input_image", "image_url": f"data:{mime};base64,{b64}", "detail": "high"}


def ask_json(prompt: str, *, web_search: bool = False, images: Iterable[UploadedAsset] | None = None) -> dict[str, Any]:
    client = get_client()
    content: list[dict[str, Any]] = [{"type": "input_text", "text": prompt}]
    for asset in images or []:
        content.append(image_content(asset))
    kwargs: dict[str, Any] = {
        "model": MODEL,
        "input": [{"role": "user", "content": content}],
    }
    if web_search:
        kwargs["tools"] = [{"type": "web_search"}]
        kwargs["tool_choice"] = "required"
    response = client.responses.create(**kwargs)
    return extract_json(response.output_text)


def docx_text(data: bytes) -> str:
    """DOCX metnini, Word denklem düğümleri (m:t) dahil olacak şekilde çıkarır."""
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            xml = zf.read("word/document.xml")
        root = ET.fromstring(xml)
        paragraphs: list[str] = []
        for node in root.iter():
            if node.tag.rsplit("}", 1)[-1] != "p":
                continue
            pieces: list[str] = []
            for child in node.iter():
                local = child.tag.rsplit("}", 1)[-1]
                if local == "t" and child.text:
                    pieces.append(child.text)
                elif local in {"tab"}:
                    pieces.append("\t")
                elif local in {"br", "cr"}:
                    pieces.append("\n")
            text = "".join(pieces).strip()
            if text:
                paragraphs.append(text)
        if paragraphs:
            return "\n".join(paragraphs)
    except Exception:
        pass

    # Bozuk/alışılmadık DOCX için güvenli geri dönüş.
    doc = Document(io.BytesIO(data))
    parts: list[str] = []
    for p in doc.paragraphs:
        if p.text.strip():
            parts.append(p.text.strip())
    for table in doc.tables:
        for row in table.rows:
            vals = [cell.text.strip().replace("\n", " | ") for cell in row.cells]
            if any(vals):
                parts.append("\t".join(vals))
    return "\n".join(parts)


def pdf_text(data: bytes) -> str:
    reader = PdfReader(io.BytesIO(data))
    return "\n".join((page.extract_text() or "") for page in reader.pages)


def legacy_doc_text(data: bytes, filename: str) -> str:
    with tempfile.TemporaryDirectory() as td:
        source = Path(td) / Path(filename).name
        source.write_bytes(data)
        try:
            result = subprocess.run(
                ["antiword", str(source)], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60
            )
            text = result.stdout.decode("utf-8", errors="replace")
            if text.strip():
                return text
        except Exception:
            pass
        outdir = Path(td) / "converted"
        outdir.mkdir()
        try:
            subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "docx", "--outdir", str(outdir), str(source)],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=120,
            )
            files = list(outdir.glob("*.docx"))
            if files:
                return docx_text(files[0].read_bytes())
        except Exception as exc:
            raise ValueError("Eski .doc dosyası okunamadı; .docx olarak kaydedip yükleyin.") from exc
    raise ValueError("Eski .doc dosyası okunamadı.")


def extract_text_from_asset(asset: UploadedAsset) -> str:
    suffix = Path(asset.name).suffix.lower()
    if suffix == ".docx":
        text = docx_text(asset.data)
    elif suffix == ".doc":
        text = legacy_doc_text(asset.data, asset.name)
    elif suffix == ".pdf":
        text = pdf_text(asset.data)
    elif suffix in {".txt", ".md"}:
        text = asset.data.decode("utf-8", errors="replace")
    elif suffix in {".png", ".jpg", ".jpeg", ".webp"}:
        text = f"[GÖRSEL DOSYASI: {asset.name}]"
    else:
        text = ""
    return text.replace("\x00", " ").strip()[:MAX_TEXT_PER_FILE]


def assets_from_uploads(files: Iterable[Any] | None) -> list[UploadedAsset]:
    out: list[UploadedAsset] = []
    for f in files or []:
        data = f.getvalue()
        suffix = Path(f.name).suffix.lower()
        if suffix == ".zip":
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                for info in zf.infolist():
                    if info.is_dir() or info.file_size > 30 * 1024 * 1024:
                        continue
                    inner_name = Path(info.filename).name
                    inner_suffix = Path(inner_name).suffix.lower()
                    if inner_suffix not in {".pdf", ".docx", ".doc", ".txt", ".md", ".png", ".jpg", ".jpeg", ".webp"}:
                        continue
                    out.append(UploadedAsset(inner_name, zf.read(info)))
        else:
            out.append(UploadedAsset(f.name, data, getattr(f, "type", "application/octet-stream")))
    return out


def combine_asset_text(label: str, assets: list[UploadedAsset]) -> tuple[str, list[UploadedAsset]]:
    blocks: list[str] = []
    images: list[UploadedAsset] = []
    total = 0
    for asset in assets:
        suffix = Path(asset.name).suffix.lower()
        if suffix in {".png", ".jpg", ".jpeg", ".webp"}:
            images.append(asset)
            blocks.append(f"\n--- {label}: {asset.name} (görsel ayrıca eklenmiştir) ---\n")
            continue
        text = extract_text_from_asset(asset)
        if not text:
            continue
        remain = MAX_TOTAL_TEXT - total
        if remain <= 0:
            break
        text = text[:remain]
        total += len(text)
        blocks.append(f"\n--- {label}: {asset.name} ---\n{text}\n")
    return "".join(blocks), images


def _valid_figure_image(data: bytes, *, min_width: int = 260, min_height: int = 160) -> bool:
    try:
        with Image.open(io.BytesIO(data)) as im:
            width, height = im.size
            return width >= min_width and height >= min_height
    except Exception:
        return False


def extract_embedded_images(asset: UploadedAsset) -> list[UploadedAsset]:
    """DOCX/PDF içindeki büyük görselleri özgün baytlarıyla çıkarır."""
    suffix = Path(asset.name).suffix.lower()
    images: list[UploadedAsset] = []

    if suffix in {".png", ".jpg", ".jpeg", ".webp"}:
        return [asset] if _valid_figure_image(asset.data, min_width=1, min_height=1) else []

    if suffix == ".docx":
        try:
            with zipfile.ZipFile(io.BytesIO(asset.data)) as zf:
                media = sorted(
                    name for name in zf.namelist()
                    if name.startswith("word/media/") and not name.endswith("/")
                )
                for index, name in enumerate(media, 1):
                    data = zf.read(name)
                    if not _valid_figure_image(data):
                        continue
                    ext = Path(name).suffix.lower() or ".png"
                    mime = "image/jpeg" if ext in {".jpg", ".jpeg"} else f"image/{ext.lstrip('.')}"
                    images.append(UploadedAsset(f"{Path(asset.name).stem}_sekil_{index}{ext}", data, mime))
        except Exception:
            return []
        return images

    if suffix == ".pdf" and fitz is not None:
        try:
            pdf = fitz.open(stream=asset.data, filetype="pdf")
            seen: set[int] = set()
            counter = 0
            for page in pdf:
                for info in page.get_images(full=True):
                    xref = int(info[0])
                    if xref in seen:
                        continue
                    seen.add(xref)
                    extracted = pdf.extract_image(xref)
                    data = extracted.get("image", b"")
                    if not data or not _valid_figure_image(data):
                        continue
                    counter += 1
                    ext = "." + str(extracted.get("ext", "png")).lower()
                    mime = "image/jpeg" if ext in {".jpg", ".jpeg"} else f"image/{ext.lstrip('.')}"
                    images.append(UploadedAsset(f"{Path(asset.name).stem}_sekil_{counter}{ext}", data, mime))
            pdf.close()
        except Exception:
            return []
    return images


def build_figures_docx(images: list[UploadedAsset]) -> bytes:
    if not images:
        raise ValueError("Şekiller Word dosyası için BBF içinde veya ayrıca yüklenen dosyalarda kullanılabilir görsel bulunamadı.")
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)
    for index, asset in enumerate(images, 1):
        add_text(doc, f"ŞEKİL {index}", bold=True, center=True)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            p.add_run().add_picture(io.BytesIO(asset.data), width=Cm(16.0))
        except Exception as exc:
            raise ValueError(f"{asset.name} şekiller dosyasına eklenemedi.") from exc
        if index < len(images):
            doc.add_page_break()
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# -----------------------------------------------------------------------------
# DOCX YARDIMCILARI
# -----------------------------------------------------------------------------
def clear_body(doc: Document) -> None:
    body = doc._element.body
    for child in list(body):
        if child.tag.endswith("sectPr"):
            continue
        body.remove(child)


def set_cell_text(cell, text: str, bold: bool = False, size: int = 10):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(text)
    run.bold = bold
    run.font.name = "Arial"
    run.font.size = Pt(size)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def format_paragraph(p, *, bold: bool = False, center: bool = False, italic: bool = False, size: int = 11):
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(size)
        # Set explicit direct formatting. The 696809 template's Normal style is
        # bold by default, so leaving run.bold=None would accidentally render
        # ordinary opinion paragraphs in bold.
        run.bold = bool(bold)
        run.italic = bool(italic)
    return p


def add_text(doc: Document, text: str, *, bold: bool = False, center: bool = False, italic: bool = False, size: int = 11):
    p = doc.add_paragraph()
    p.add_run(text)
    return format_paragraph(p, bold=bold, center=center, italic=italic, size=size)


def add_heading(doc: Document, text: str, *, center: bool = False):
    return add_text(doc, text, bold=True, center=center)


def add_bullet(doc: Document, text: str, *, symbol: str = "•"):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(1.25)
    p.paragraph_format.first_line_indent = Cm(-0.5)
    p.add_run(f"{symbol}\t{text}")
    return format_paragraph(p)


def add_quote(doc: Document, text: str):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(f'“{text}”')
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(11)
    return p


def add_page_number(section):
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = "PAGE"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instruction, end])


def copy_template_paragraph(doc: Document, template: Document, index: int):
    if 0 <= index < len(template.paragraphs):
        doc._element.body.insert(-1, deepcopy(template.paragraphs[index]._p))


def safe_output_name(name: str, default: str) -> str:
    """Return a human-readable, filesystem-safe DOCX download name.

    Browser/URL encoded names such as ``Görüş%20Metni_698891.docx`` or
    ``G%C3%B6r%C3%BC%C5%9F%20Metni_698891.docx`` are decoded before they
    reach Streamlit's ``file_name`` parameter. Turkish characters and normal
    spaces are deliberately preserved; the downloaded file must not contain
    literal ``%20``/``%C3...`` fragments.
    """
    raw = str(name or default).strip()

    # Decode once or twice so a previously double-encoded filename is also
    # repaired, while avoiding an unbounded decode loop.
    for _ in range(2):
        decoded = unquote(raw)
        if decoded == raw:
            break
        raw = decoded

    raw = raw.replace("\u00a0", " ")
    raw = re.sub(r"[\r\n\t]+", " ", raw)
    raw = re.sub(r"\s+", " ", raw).strip()

    # Strip any user-supplied path and keep only the filename. Handle both
    # Windows and POSIX separators regardless of the server OS.
    raw = raw.replace("\\", "/").split("/")[-1].strip()
    if not raw:
        raw = default
    if not raw.lower().endswith(".docx"):
        raw += ".docx"
    return raw


# -----------------------------------------------------------------------------
# TARİFNAME MODÜLÜ
# -----------------------------------------------------------------------------
TARIFNAME_DRAFT_SCHEMA = r"""
{
  "title":"",
  "technical_field":"",
  "prior_art_general_paragraphs":[""],
  "literature_paragraphs":[""],
  "short_description_intro":"",
  "objectives":[""],
  "unumbered_invention_definition":"",
  "unumbered_invention_features":[""],
  "figure_descriptions":[""],
  "elements":[{"number":"10","name":"","description":""}],
  "method_steps":[{"number":"1001","text":""}],
  "detailed_paragraphs":[""],
  "formulas":[{"label":"","expression":"","explanation":""}],
  "tables":[{"caption":"","headers":[""],"rows":[[""]]}],
  "experimental_results":[""],
  "alternatives":[""],
  "working_principle":"",
  "system_claim":null,
  "dependent_system_claims":[""],
  "method_claim":null,
  "dependent_method_claims":[""],
  "abstract":"",
  "coverage_audit":{
    "prior_art_complete":true,
    "reference_table_complete":true,
    "formulas_complete":true,
    "tables_complete":true,
    "experimental_results_complete":true,
    "alternatives_complete":true,
    "claims_consistent":true,
    "notes":[""]
  }
}
"""


def tarifname_extraction_prompt(
    source_text: str,
    current_draft_text: str = "",
    technical_supplement_text: str = "",
    example_structure_text: str = "",
) -> str:
    return f"""{TARIFNAME_RULES}
Aşağıdaki kaynakları yalnızca yapı ve kapsam envanteri çıkarmak için incele. Teknik metni yeniden icat etme,
kısaltma nedeniyle önemli bilgi kaybettirme ve örnek tarifnamelerin teknik içeriğini kullanma.

KAYNAK HİYERARŞİSİ:
1. BBF: temel teknik kaynak.
2. Mevcut revize tarifname: kullanıcının aktardığı/düzelttiği teknik metin korunacak kaynak.
3. Ek teknik müşteri belgeleri: yalnızca açık teknik dayanak olarak kullanılabilir.
4. Örnek tarifnameler: yalnızca unsur, yöntem adımı ve istem kurgusunu görmek içindir; teknik içerikleri kullanılamaz.

JSON dışında hiçbir şey yazma.
ŞEMA:
{{
 "title":"",
 "technical_field":"",
 "prior_art_inventory":["BBF'deki her ayrı önceki teknik konu ve kısıt"],
 "technical_problems":[""],
 "technical_solution":[""],
 "technical_effects":[""],
 "elements":[{{"number":"10","name":"","function":"","source":"BBF/mevcut tarifname"}}],
 "method_steps":[{{"number":"1001","text":"","stage":"","essential":true}}],
 "formulas":[{{"label":"","expression":"","variables":[""],"role":"zorunlu/tercihli"}}],
 "tables":[{{"caption":"","headers":[""],"rows":[[""]]}}],
 "experimental_results":[""],
 "alternatives":[""],
 "use_cases":[""],
 "figures":[""],
 "claim_core":[""],
 "parallel_step_groups":[{{"summary":"","step_numbers":["1007","1008"],"recommended_claim_location":"ana istem/tek bağımlı istem"}}],
 "stage_distinctions":[{{"step_numbers":["1001","1006"],"difference":""}}],
 "has_system_basis":true,
 "has_method_basis":true,
 "recommended_claim_mode":"Yalnızca yöntem/Sistem ve yöntem/Yalnızca sistem",
 "recommended_claim_mode_reason":"",
 "coverage_checklist":["Kaynakta bulunan ve taslakta mutlaka korunması gereken her içerik grubu"]
}}

BBF:
---
{source_text}
---

MEVCUT REVİZE TARİFNAME:
---
{current_draft_text}
---

EK TEKNİK BELGELER/NOTLAR:
---
{technical_supplement_text}
---

ÖRNEK TARİFNAMELER - YALNIZCA KURGU:
---
{example_structure_text}
---
"""


def tarifname_literature_prompt(extracted: dict[str, Any], count: int, jurisdiction: str) -> str:
    return f"""Aşağıdaki buluş için tam olarak {count} teknik olarak yakın patent dokümanı araştır. Web araması kullan.
Doküman uydurma; yayın/başvuru numarasını, İngilizce başlığını, yayın tarihini ve kaynak bağlantısını doğrula.
Tercih edilen ülke/veri tabanı: {jurisdiction or 'global'}.
Her doküman için teknik konusu ile buluşta bulunmayan temel teknik farkı açıkla. Bu aşamada tarifname metni yazma.
JSON dışında yazma.
ŞEMA:
{{"documents":[{{"application_number":"","title_en":"","title_tr":"","publication_date":"","jurisdiction":"","summary":"","difference":"","source_url":""}}]}}
BULUŞ ENVANTERİ:
{json.dumps(extracted, ensure_ascii=False, indent=2)}"""


def tarifname_drafting_prompt(
    extracted: dict[str, Any],
    claim_mode: str,
    literature: list[dict[str, Any]],
    source_text: str,
    current_draft_text: str,
    technical_supplement_text: str,
    example_structure_text: str,
) -> str:
    return f"""{TARIFNAME_RULES}
Aşağıdaki kaynaklara dayanarak Türk patent tarifnamesinin TAM metnini oluştur.
İstem yapısı: {claim_mode}

KRİTİK TALİMATLAR:
- BBF'nin ve mevcut revize tarifnamenin bütün teknik bilgilerini kullan. Uzun önceki teknik, formüller, tablolar,
  deneysel sonuçlar, alternatifler ve referans tablosu atlanamaz.
- Yapılandırılmış envanter yalnızca yardımcıdır. Çelişki halinde ham BBF ve mevcut revize tarifname esas alınır.
- Mevcut tarifnamedeki kullanıcı düzenlemelerini koru; teknik içerikte gereksiz kısaltma yapma.
- Örnek tarifnamelerden yalnızca kurguyu öğren; teknik bilgi aktarma.
- Ana istemde zorunlu teknik çekirdeği kapsayıcı biçimde ver. Aynı işlemin birinci/ikinci/k'ıncı tekrarlarını
  ana istemde gereksiz yere ayrı satırlara bölme. Bu ayrıntıları, aynı alt akışa aitse tek bağımlı istemde topla.
- Eğitim/genel aşama ile test aşamasındaki paralel akışları aynı mantıkla fakat ayrı teknik aşamalar olarak kur.
- REFERANS NUMARALARI bölümündeki yöntem adımları tam liste olarak korunur. Ana istemde numarasız kapsayıcı
  ifade kullanılabilir; ayrıntılı numaralı adımlar bağımlı istemde verilebilir.
- "Yöntemin gerçekleştirdiği işlem adımları aşağıdaki gibidir:" bölümü için method_steps tam ve tutarlı olsun.
- Yalnızca yöntem modunda system_claim null olmalıdır. Yalnızca sistem modunda method_claim null olmalıdır.
- Her bağımlı istem ana isteme göre gerçek bir daraltma sağlamalıdır.
- Patent literatürü yalnızca ÖNCEKİ TEKNİK bölümünde, her doküman ayrı paragraf olacak biçimde kullanılsın.
- Kullanıcıya sunulan tarifname metninde “BBF”, “buluş bildirim formu” veya kaynak dokümana atıf yapan benzer ifadeler kesinlikle bulunmasın; teknik bilgi doğrudan buluş anlatımı olarak yazılsın.
- Sistem ve yöntem istemleri birlikte oluşturuluyorsa başlık da buna uygun biçimde “... Sistemi ve Yöntemi” olsun; yalnızca “... Sistemi” olarak bırakılmasın.
- REFERANS NUMARALARI bölümünde unsur adlarında yalnızca ilk kelimenin ilk harfi büyük olsun; standart teknik kısaltmaları koru. Unsur adları cümle içinde geçtiğinde küçük harfle başlat.
- “Buluşun bir gerçekleştirilmesinde” ifadesini kullanma; gerekli yerde “Buluşun bir yapılanmasında” yaz.
- Önceki teknik bölümüne ham kaynakta verilen bütün teknik arka plan, eksiklik ve problem anlatımını aktar; literatür paragrafları bunların yerine geçmez.

JSON dışında hiçbir şey yazma.
ÇIKTI ŞEMASI:
{TARIFNAME_DRAFT_SCHEMA}

SİSTEM İSTEMİ ŞEMASI (varsa):
{{"preamble":"","elements":[""],"closing":"içermesidir."}}
YÖNTEM İSTEMİ ŞEMASI (varsa):
{{"preamble":"","steps":[""],"closing":"işlem adımlarını içermesidir."}}

YAPILANDIRILMIŞ ENVANTER:
{json.dumps(extracted, ensure_ascii=False, indent=2)}

ONAYLANAN PATENT LİTERATÜRÜ:
{json.dumps(literature, ensure_ascii=False, indent=2)}

HAM BBF:
---
{source_text}
---

MEVCUT REVİZE TARİFNAME:
---
{current_draft_text}
---

EK TEKNİK BELGELER/NOTLAR:
---
{technical_supplement_text}
---

ÖRNEK TARİFNAMELER - YALNIZCA KURGU:
---
{example_structure_text}
---
"""


def tarifname_quality_prompt(
    source_text: str,
    current_draft_text: str,
    technical_supplement_text: str,
    extracted: dict[str, Any],
    draft: dict[str, Any],
    claim_mode: str,
) -> str:
    return f"""{TARIFNAME_RULES}
Aşağıdaki tarifname taslağını kaynaklarla SATIR SATIR karşılaştır ve eksik/yanlış hususları düzelterek tam JSON'u yeniden üret.
Bu bir özetleme görevi değildir. Kaynakta olup taslakta bulunmayan her teknik bilgi geri eklenmelidir.

ZORUNLU KONTROL LİSTESİ:
1. BBF'deki önceki teknik anlatımının tamamı korunmuş mu?
2. Referans tablosundaki bütün unsurlar ve yöntem adımları var mı?
3. Farklı numaralı adımlar yanlışlıkla aynı metinle mi yazılmış? Aynı veri farklı aşamada kullanılıyorsa aşama farkı açık mı?
4. "Yöntemin gerçekleştirdiği işlem adımları" tam liste ve referans tablosuyla uyumlu mu?
5. Ana istem zorunlu çekirdeği kapsıyor mu; paralel tekrarları gereksiz yere tek tek sayıyor mu?
6. Aynı alt akışa ait paralel analizler ve çıktılar gerekiyorsa tek bağımlı istemde mi?
7. Eğitim/genel ve test aşamaları paralel fakat ayrı olarak mı kurulmuş?
8. Formüller, değişken açıklamaları, tablolar, deneysel sonuçlar, alternatifler ve teknik etkiler eksiksiz mi?
9. Seçilen istem modu ({claim_mode}) ile başlık, açıklama ve istemler tutarlı mı? Sistem ve yöntem ise başlık “Sistemi ve Yöntemi” yapısını taşıyor mu?
10. Bağımlı istemler gerçek daraltma sağlıyor mu?
11. Kullanıcıya sunulan metinde “BBF” veya “buluş bildirim formu” gibi kaynak atfı kalmış mı? Kalmışsa doğrudan buluş anlatımına dönüştür.
12. “Buluşun bir gerçekleştirilmesinde” kalıbı var mı? Varsa “Buluşun bir yapılanmasında” olarak düzelt.
13. REFERANS NUMARALARI unsur adları yalnızca ilk kelime büyük olacak biçimde mi? Cümle içindeki unsur adları küçük harfle mi başlıyor?
14. Önceki teknik kaynakta verilen bütün müşteri teknik arka planını ve eksikliklerini içeriyor mu?

JSON dışında hiçbir şey yazma. Çıktı, aşağıdaki şemaya tam uymalıdır:
{TARIFNAME_DRAFT_SCHEMA}

YAPILANDIRILMIŞ ENVANTER:
{json.dumps(extracted, ensure_ascii=False, indent=2)}

HAM BBF:
---
{source_text}
---

MEVCUT REVİZE TARİFNAME:
---
{current_draft_text}
---

EK TEKNİK BELGELER:
---
{technical_supplement_text}
---

KONTROL EDİLECEK TASLAK:
{json.dumps(draft, ensure_ascii=False, indent=2)}
"""


def _strip_claim_number(text: str) -> str:
    return re.sub(r"^\s*\d+\s*[.)-]\s*", "", str(text or "")).strip()


def add_numbered_claim(doc: Document, number: int, text: str):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    r_num = p.add_run(f"{number}. ")
    r_num.bold = True
    r_num.font.name = "Arial"
    r_num.font.size = Pt(11)
    r_text = p.add_run(_strip_claim_number(text))
    r_text.font.name = "Arial"
    r_text.font.size = Pt(11)
    return p


def _replace_in_nested(value: Any, old: str, new: str) -> Any:
    if isinstance(value, str):
        return value.replace(old, new)
    if isinstance(value, list):
        return [_replace_in_nested(x, old, new) for x in value]
    if isinstance(value, dict):
        return {k: _replace_in_nested(v, old, new) for k, v in value.items()}
    return value


def _ensure_title_for_claim_mode(title: str, claim_mode: str) -> str:
    text = str(title or "").strip()
    if claim_mode != "Sistem ve yöntem" or not text:
        return text
    if "yöntem" in text.lower():
        return text
    matches = list(re.finditer(r"\bsistemi\b", text, flags=re.IGNORECASE))
    if matches:
        last = matches[-1]
        return text[: last.start()] + "Sistemi ve Yöntemi" + text[last.end() :]
    return text + " Sistemi ve Yöntemi"


def apply_tarifname_house_style(draft: dict[str, Any], claim_mode: str) -> dict[str, Any]:
    """Kullanıcının sabit terminoloji ve başlık tercihlerini çıktıdan önce uygula."""
    draft = deepcopy(draft)
    for old, new in [
        ("Buluşun bir gerçekleştirilmesinde", "Buluşun bir yapılanmasında"),
        ("buluşun bir gerçekleştirilmesinde", "buluşun bir yapılanmasında"),
    ]:
        draft = _replace_in_nested(draft, old, new)
    draft["title"] = _ensure_title_for_claim_mode(draft.get("title", ""), claim_mode)
    return draft


def validate_tarifname_draft(draft: dict[str, Any], claim_mode: str) -> list[str]:
    warnings: list[str] = []
    steps = draft.get("method_steps") or []
    numbers = [str(x.get("number", "")).strip() for x in steps]
    if len(numbers) != len(set(numbers)):
        raise ValueError("REFERANS NUMARALARI bölümünde yinelenen yöntem adımı numarası bulundu.")
    if any(not n for n in numbers):
        raise ValueError("Numarası boş yöntem işlem adımı bulundu.")

    normalized_to_numbers: dict[str, list[str]] = {}
    for step in steps:
        text = re.sub(r"\s+", " ", str(step.get("text", ""))).strip().lower()
        if not text:
            raise ValueError(f"{step.get('number','?')} numaralı yöntem adımının metni boş.")
        normalized_to_numbers.setdefault(text, []).append(str(step.get("number", "")))
    duplicate_groups = [nums for nums in normalized_to_numbers.values() if len(nums) > 1]
    if duplicate_groups:
        warnings.append(
            "Aynı metinle yazılmış farklı yöntem adımları bulundu: "
            + "; ".join(", ".join(group) for group in duplicate_groups)
            + ". Bu adımların farklı aşamaları temsil edip etmediğini kontrol edin."
        )

    all_claim_text = json.dumps(
        {
            "system_claim": draft.get("system_claim"),
            "dependent_system_claims": draft.get("dependent_system_claims"),
            "method_claim": draft.get("method_claim"),
            "dependent_method_claims": draft.get("dependent_method_claims"),
        },
        ensure_ascii=False,
    )
    referenced = set(re.findall(r"\((1\d{3})\)", all_claim_text))
    missing = sorted(referenced - set(numbers))
    if missing:
        raise ValueError("İstemlerde bulunup referans listesinde bulunmayan yöntem adımları: " + ", ".join(missing))

    if claim_mode == "Yalnızca yöntem" and draft.get("system_claim"):
        raise ValueError("Yalnızca yöntem seçildiği halde sistem istemi üretildi.")
    if claim_mode == "Yalnızca sistem" and draft.get("method_claim"):
        raise ValueError("Yalnızca sistem seçildiği halde yöntem istemi üretildi.")
    if claim_mode in {"Yalnızca yöntem", "Sistem ve yöntem"} and not draft.get("method_claim"):
        raise ValueError("Seçilen istem yapısına rağmen bağımsız yöntem istemi üretilemedi.")
    if claim_mode in {"Yalnızca sistem", "Sistem ve yöntem"} and not draft.get("system_claim"):
        raise ValueError("Seçilen istem yapısına rağmen bağımsız sistem istemi üretilemedi.")

    user_facing_text = json.dumps(draft, ensure_ascii=False)
    if re.search(r"\bBBF\b|buluş bildirim formu", user_facing_text, flags=re.IGNORECASE):
        raise ValueError("Tarifname taslağında kullanıcıya görünmemesi gereken BBF/kaynak form atfı bulundu.")
    if claim_mode == "Sistem ve yöntem" and "yöntem" not in str(draft.get("title", "")).lower():
        raise ValueError("Sistem ve yöntem istem yapısında başlık yöntem ifadesini içermiyor.")
    return warnings


def _reference_sentence_case(name: str) -> str:
    """Referans unsurunu başlık biçiminden çıkar; teknik kısaltmaları koru."""
    text = str(name or "").strip()
    first_seen = False

    def repl(match: re.Match[str]) -> str:
        nonlocal first_seen
        word = match.group(0)
        is_acronym = len(word) > 1 and word.isupper()
        if not first_seen:
            first_seen = True
            if is_acronym:
                return word
            return word[:1].upper() + word[1:].lower()
        if is_acronym:
            return word
        return word.lower()

    return re.sub(r"[A-Za-zÇĞİÖŞÜçğıöşü]+", repl, text)


def build_tarifname_docx(draft: dict[str, Any]) -> bytes:
    template = Document(str(TARIFNAME_TEMPLATE))
    doc = Document(str(TARIFNAME_TEMPLATE))
    clear_body(doc)
    for section in doc.sections:
        section.top_margin = Cm(3)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(3)
        section.right_margin = Cm(2)

    add_heading(doc, "TARİFNAME", center=True)
    doc.add_paragraph()
    add_text(doc, draft.get("title", ""), bold=True, center=True)
    doc.add_paragraph()
    copy_template_paragraph(doc, template, 4)
    doc.add_paragraph()

    add_heading(doc, "TEKNİK ALAN")
    add_text(doc, draft.get("technical_field", ""))

    add_heading(doc, "ÖNCEKİ TEKNİK")
    prior = draft.get("prior_art_general_paragraphs") or []
    for index, paragraph in enumerate(prior):
        add_text(doc, paragraph)
        if index < len(prior) - 1:
            doc.add_paragraph()
    literature = draft.get("literature_paragraphs") or []
    if prior and literature:
        doc.add_paragraph()
    for index, paragraph in enumerate(literature):
        add_text(doc, paragraph)
        if index < len(literature) - 1:
            doc.add_paragraph()
    add_text(doc, "Sonuçta yukarıda bahsedilen ve mevcut teknik ışığında çözülemeyen sorunlar, ilgili teknik alanda bir yenilik yapmayı zorunlu kılmıştır.")

    add_heading(doc, "BULUŞUN KISA AÇIKLAMASI")
    add_text(doc, draft.get("short_description_intro", ""))
    for index, objective in enumerate(draft.get("objectives") or []):
        prefix = "Buluşun ana amacı, " if index == 0 else "Buluşun diğer bir amacı, "
        objective = str(objective).strip()
        add_text(doc, prefix + (objective[:1].lower() + objective[1:] if objective else ""))
    invention_definition = draft.get("unumbered_invention_definition") or draft.get("unumbered_system_definition")
    if invention_definition:
        add_text(doc, invention_definition)
    invention_features = draft.get("unumbered_invention_features") or draft.get("unumbered_system_elements") or []
    for item in invention_features:
        add_bullet(doc, item)
    if invention_features:
        add_text(doc, "işlem adımlarını içermesidir." if draft.get("method_claim") and not draft.get("system_claim") else "içermesidir.")

    add_heading(doc, "ŞEKİLLERİN KISA AÇIKLAMASI")
    for figure in draft.get("figure_descriptions") or ["Şekil 1, buluşa konu yapılanmanın temsili gösterimidir."]:
        add_text(doc, figure)

    add_heading(doc, "REFERANS NUMARALARI")
    for element in draft.get("elements") or []:
        add_text(doc, f"{element.get('number','')}. {_reference_sentence_case(element.get('name',''))}")
    method_steps = draft.get("method_steps") or []
    if draft.get("elements") and method_steps:
        doc.add_paragraph()
    for step in method_steps:
        text = re.sub(r"\s*\(\s*\d+\s*\)\s*", "", str(step.get("text", ""))).strip()
        add_text(doc, f"{step.get('number','')}. {text}")

    add_heading(doc, "BULUŞUN DETAYLI AÇIKLAMASI")
    title = str(draft.get("title", "buluş")).lower()
    add_text(doc, f"Bu detaylı açıklamada, buluş konusu olan {title} sadece konunun daha iyi anlaşılmasına yönelik hiçbir sınırlayıcı etki oluşturmayacak örneklerle açıklanmaktadır.")
    for paragraph in draft.get("detailed_paragraphs") or []:
        add_text(doc, paragraph)

    for formula in draft.get("formulas") or []:
        if formula.get("label"):
            add_text(doc, formula.get("label", ""), bold=True)
        if formula.get("expression"):
            add_text(doc, formula.get("expression", ""), center=True)
        if formula.get("explanation"):
            add_text(doc, formula.get("explanation", ""))

    for table_data in draft.get("tables") or []:
        caption = table_data.get("caption", "")
        if caption:
            add_text(doc, caption, bold=True)
        headers = [str(x) for x in table_data.get("headers") or []]
        rows = table_data.get("rows") or []
        column_count = max(len(headers), max((len(row) for row in rows), default=0))
        if column_count:
            table = doc.add_table(rows=1 if headers else 0, cols=column_count)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            if headers:
                for idx in range(column_count):
                    set_cell_text(table.rows[0].cells[idx], headers[idx] if idx < len(headers) else "", bold=True)
            for row_data in rows:
                cells = table.add_row().cells
                for idx in range(column_count):
                    set_cell_text(cells[idx], str(row_data[idx]) if idx < len(row_data) else "")

    for paragraph in draft.get("experimental_results") or []:
        add_text(doc, paragraph)
    for paragraph in draft.get("alternatives") or []:
        add_text(doc, paragraph)

    if method_steps:
        add_text(doc, "Yöntemin gerçekleştirdiği işlem adımları aşağıdaki gibidir:")
        for step in method_steps:
            text = re.sub(r"\s*\(\s*\d+\s*\)\s*", "", str(step.get("text", ""))).strip()
            add_bullet(doc, f"{text} ({step.get('number','')})", symbol="-")
    if draft.get("working_principle"):
        add_text(doc, draft.get("working_principle", ""))

    doc.add_page_break()
    add_heading(doc, "İSTEMLER", center=True)
    for index in (79, 81, 83):
        copy_template_paragraph(doc, template, index)

    claim_no = 1
    system_claim = draft.get("system_claim")
    if system_claim:
        add_numbered_claim(doc, claim_no, f"{system_claim.get('preamble','')} olup, özelliği;")
        for item in system_claim.get("elements") or []:
            add_bullet(doc, item, symbol="-")
        add_text(doc, system_claim.get("closing", "içermesidir."))
        claim_no += 1
        for dependent in draft.get("dependent_system_claims") or []:
            add_numbered_claim(doc, claim_no, dependent)
            claim_no += 1

    method_claim = draft.get("method_claim")
    if method_claim:
        add_numbered_claim(doc, claim_no, f"{method_claim.get('preamble','')} olup, özelliği;")
        for item in method_claim.get("steps") or []:
            add_bullet(doc, item, symbol="-")
        add_text(doc, method_claim.get("closing", "işlem adımlarını içermesidir."))
        claim_no += 1
        for dependent in draft.get("dependent_method_claims") or []:
            add_numbered_claim(doc, claim_no, dependent)
            claim_no += 1

    doc.add_page_break()
    add_heading(doc, "ÖZET", center=True)
    add_text(doc, draft.get("title", ""), bold=True, center=True)
    add_text(doc, draft.get("abstract", ""))
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# -----------------------------------------------------------------------------
# GÖRÜŞ MODÜLÜ
# -----------------------------------------------------------------------------
def gorus_analysis_prompt(
    report_type: str,
    reference: str,
    report_text: str,
    spec_text: str,
    prior_opinion_text: str,
    similar_text: str,
    customer_text: str,
) -> str:
    return f"""{GORUS_RULES}
Aşağıdaki dosyaları önce YALNIZCA analiz et. Bu aşamada görüş Word metni yazma.
Görüş türü: {report_type}
Ana dosya referansı: {reference}

Önce rapordaki itirazları, X/Y dokümanlarını, mevcut istemleri, tarifname dayanaklarını, varsa önceki görüşü ve müşteri bilgisini birlikte değerlendir.
İstem değişikliği sırf daha iyi yazılabilir diye önerilmez. Yalnızca itirazı gidermek için gerçekten zorunluysa amendment_required=true yap.
Revizyon gerekiyorsa EN AZ DEĞİŞİKLİK ilkesini uygula. Her old_text, TARİFNAME içindeki tek bir paragrafta birebir bulunabilen mümkün olan en kısa ifade olsun; tüm istemi old_text olarak verme.
Her basis_quote tarifnamede birebir bulunan dayanak pasajı olsun. Kapsam aşımı/yeni konu yaratma.

JSON dışında yazma.
ŞEMA:
{{
  "analysis_summary":"",
  "examiner_issues":[""],
  "defense_direction":[""],
  "amendment_required":false,
  "amendment_reason":"",
  "no_amendment_reason":"",
  "amendments":[
    {{
      "claim_number":"1",
      "reason":"",
      "basis_quote":"",
      "old_text":"",
      "new_text":""
    }}
  ]
}}

RAPOR:\n{report_text}\n
TARİFNAME:\n{spec_text}\n
ÖNCEKİ GÖRÜŞ:\n{prior_opinion_text}\n
BENZER DOKÜMANLAR:\n{similar_text}\n
MÜŞTERİ BİLGİLERİ:\n{customer_text}\n"""


def gorus_revision_refine_prompt(
    report_type: str,
    report_text: str,
    spec_text: str,
    prior_opinion_text: str,
    similar_text: str,
    customer_text: str,
    current_analysis: dict[str, Any],
    user_instruction: str,
) -> str:
    return f"""{GORUS_RULES}
Aşağıdaki mevcut istem-revizyon analizini kullanıcının talimatına göre yeniden değerlendir.
Yalnızca gerçekten zorunlu değişiklikleri bırak. En az değişiklik, açık tarifname dayanağı ve kapsam aşımı yapmama kuralları bağlayıcıdır.
old_text, kaynak TARİFNAME içindeki tek bir paragrafta birebir bulunabilen mümkün olan en kısa ifade olmalıdır.
Tüm istemi silip yeniden yazma. Kullanıcının talimatı teknik kaynaklarla çelişiyorsa uydurma yapma; güvenli olan minimum revizyonu seç.

JSON dışında yazma ve ilk analizle AYNI şemayı kullan.
Görüş türü: {report_type}
KULLANICI TALİMATI:\n{user_instruction}\n
MEVCUT ANALİZ:\n{json.dumps(current_analysis, ensure_ascii=False, indent=2)}\n
RAPOR:\n{report_text}\n
TARİFNAME:\n{spec_text}\n
ÖNCEKİ GÖRÜŞ:\n{prior_opinion_text}\n
BENZER DOKÜMANLAR:\n{similar_text}\n
MÜŞTERİ BİLGİLERİ:\n{customer_text}\n"""


def validate_gorus_analysis(analysis: dict[str, Any], spec_text: str) -> None:
    required = bool(analysis.get("amendment_required"))
    amendments = analysis.get("amendments") or []
    if required and not amendments:
        raise ValueError("Analiz istem revizyonu gerekli dedi ancak revizyon önerisi üretmedi.")
    normalized_spec = re.sub(r"\s+", " ", spec_text).strip()
    for item in amendments:
        old_text = re.sub(r"\s+", " ", str(item.get("old_text", ""))).strip()
        new_text = str(item.get("new_text", "")).strip()
        basis = re.sub(r"\s+", " ", str(item.get("basis_quote", ""))).strip()
        if not old_text or not new_text:
            raise ValueError("Revizyon önerisinde old_text/new_text boş bırakılamaz.")
        if required and not basis:
            raise ValueError("Her istem revizyonu için tarifnameden birebir dayanak gösterilmelidir.")
        if re.match(r"^\s*\d+\s*[.)-]", old_text) or len(old_text) > 600:
            raise ValueError("Revizyon old_text alanı tüm istemi kapsamamalı; mümkün olan en küçük ifade seçilmelidir.")
        if old_text not in normalized_spec:
            raise ValueError(f"Revize edilecek eski ifade tarifnamede birebir doğrulanamadı: {old_text[:120]}...")
        if basis and basis not in normalized_spec:
            raise ValueError(f"Revizyon dayanağı tarifnamede birebir doğrulanamadı: {basis[:120]}...")


def gorus_prompt(
    report_type: str,
    reference: str,
    report_text: str,
    spec_text: str,
    prior_opinion_text: str,
    similar_text: str,
    customer_text: str,
    preanalysis: dict[str, Any] | None = None,
    revision_status: str = "Mevcut istemlerle devam",
) -> str:
    return f"""{GORUS_RULES}
Aşağıdaki dosyalara dayanarak Türk Patent ve Marka Kurumu için ayrıntılı görüş metni hazırla.
Görüş türü: {report_type}
Ana dosya referansı: {reference}
Nihai istem durumu: {revision_status}

Bu aşama ilk teknik analizden SONRA çalışmaktadır. İstemlerde kendiliğinden yeni revizyon önerme veya onaylanmış istem setini değiştirme.
Aşağıdaki ÖN ANALİZ yalnızca savunma yönünü ve kullanıcının onayladığı revizyon durumunu anlamak içindir.

JSON dışında yazma.
ŞEMA:
{{
 "application_no":"", "applicant":"", "reference":"{reference}", "report_date":"", "intro":"",
 "cited_documents":[{{"label":"D1","number":"","title":"","category":"X/Y","summary":""}}],
 "sections":[{{"heading":"D1 dokümanı:","paragraphs":[""],"quotes":[{{"lead":"Tarifnamede bu durum şu şekilde belirtilmektedir:","text":"","following":""}}]}}],
 "combined_assessment":{{"heading":"Dokümanların birlikte değerlendirilmesi","paragraphs":[""]}},
 "conclusion":[""], "signoff":"Saygılarımızla,\nDESTEK PATENT A.Ş."
}}
ÖZEL:
- Raporda X/Y olmayan dokümanı savunma bölümüne alma.
- Tarifname alıntıları spec metninde birebir geçen tam cümle/pasaj olsun.
- Müşteri bilgisinin dayanağı yoksa doğrudan kullanma.
- İnceleme raporuysa önceki görüşteki savunmaların neden ikna etmemiş olabileceğini değerlendir ve farklı teknik hat geliştir.
- Onaylı istem setine yeni değişiklik ekleme.
- Metni ikinci kez kalite kontrolünden geçir.

ÖN ANALİZ:\n{json.dumps(preanalysis or {}, ensure_ascii=False, indent=2)}\n
RAPOR:\n{report_text}\n
TARİFNAME / ONAYLI NİHAİ İSTEM SETİ:\n{spec_text}\n
ÖNCEKİ GÖRÜŞ:\n{prior_opinion_text}\n
BENZER DOKÜMANLAR:\n{similar_text}\n
MÜŞTERİ BİLGİLERİ:\n{customer_text}\n"""


def validate_quotes(opinion: dict[str, Any], spec_text: str) -> None:
    normalized_spec = re.sub(r"\s+", " ", spec_text).strip()
    for section in opinion.get("sections") or []:
        for quote in section.get("quotes") or []:
            text = re.sub(r"\s+", " ", str(quote.get("text", ""))).strip()
            if text and text not in normalized_spec:
                raise ValueError(f"Tarifname alıntısı birebir doğrulanamadı: {text[:120]}...")


def build_gorus_docx(opinion: dict[str, Any]) -> bytes:
    # 696809 is the binding opinion template. Keep its section geometry,
    # header/footer, logo, margins and page setup exactly as stored in the
    # template. The first-page title, metadata table and salutation are cloned
    # from the template so their positions/column widths do not drift.
    doc = Document(str(GORUS_TEMPLATE))

    title_1 = deepcopy(doc.paragraphs[0]._p)
    title_2 = deepcopy(doc.paragraphs[1]._p)
    blank_after_meta = deepcopy(doc.paragraphs[2]._p)
    salutation = deepcopy(doc.paragraphs[3]._p)
    metadata_table = deepcopy(doc.tables[0]._tbl)

    clear_body(doc)
    body = doc._element.body
    body.insert(-1, title_1)
    body.insert(-1, title_2)
    body.insert(-1, metadata_table)
    body.insert(-1, blank_after_meta)
    body.insert(-1, salutation)

    # Update the cloned 3-column metadata table while preserving the exact
    # 696809 column widths and layout. Labels remain bold; values remain normal.
    table = doc.tables[0]
    values = [
        opinion.get("application_no", ""),
        opinion.get("applicant", ""),
        opinion.get("reference", ""),
    ]
    for row, value in zip(table.rows, values):
        # Col 0 and col 1 already contain the template label and colon.
        set_cell_text(row.cells[2], str(value), bold=False, size=11)
        row.cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        # Preserve/force the label semantics requested for the template.
        for run in row.cells[0].paragraphs[0].runs:
            run.bold = True
            run.font.name = "Arial"
            run.font.size = Pt(11)
        for run in row.cells[1].paragraphs[0].runs:
            run.bold = False
            run.font.name = "Arial"
            run.font.size = Pt(11)

    add_text(doc, opinion.get("intro", ""))

    docs = opinion.get("cited_documents") or []
    if docs:
        doc.add_paragraph()
        for d in docs:
            add_text(doc, f"{d.get('label','')}: {d.get('number','')} {d.get('title','')}")
        doc.add_paragraph()

    for section in opinion.get("sections") or []:
        add_heading(doc, section.get("heading", ""))
        paras = list(section.get("paragraphs") or [])
        quotes = list(section.get("quotes") or [])
        for p in paras:
            add_text(doc, p)
        for q in quotes:
            add_text(doc, q.get("lead", "Tarifnamede bu durum şu şekilde belirtilmektedir:"))
            add_quote(doc, q.get("text", ""))
            if q.get("following"):
                add_text(doc, q["following"])

    combined = opinion.get("combined_assessment") or {}
    if combined:
        add_heading(doc, combined.get("heading", "Dokümanların birlikte değerlendirilmesi"))
        for p in combined.get("paragraphs") or []:
            add_text(doc, p)

    for p in opinion.get("conclusion") or []:
        add_text(doc, p)

    doc.add_paragraph()
    for line in str(opinion.get("signoff", "Saygılarımızla,\nDESTEK PATENT A.Ş.")).splitlines():
        add_text(doc, line, bold=True)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()




def _find_relaxed_phrase_span(text: str, phrase: str) -> tuple[int, int] | None:
    tokens = re.findall(r"\S+", str(phrase or "").strip())
    if not tokens:
        return None
    pattern = r"\s+".join(re.escape(token) for token in tokens)
    match = re.search(pattern, text)
    return (match.start(), match.end()) if match else None


def _claim_paragraph_map(doc: Document) -> dict[int, str | None]:
    """Her paragrafı İSTEMLER bölümü içinde en son görülen istem numarasıyla eşleştirir."""
    mapping: dict[int, str | None] = {}
    in_claims = False
    current_claim: str | None = None
    for idx, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        upper = text.upper()
        if upper == "İSTEMLER":
            in_claims = True
            current_claim = None
            mapping[idx] = None
            continue
        if in_claims and upper == "ÖZET":
            in_claims = False
            current_claim = None
        if in_claims:
            match = re.match(r"^\s*(\d+)\s*[.)-]\s+", text)
            if match:
                current_claim = match.group(1)
            mapping[idx] = current_claim
        else:
            mapping[idx] = None
    return mapping


def _run_style_spans(paragraph) -> list[tuple[int, int, Any]]:
    spans: list[tuple[int, int, Any]] = []
    cursor = 0
    for run in paragraph.runs:
        text = run.text or ""
        if not text:
            continue
        rpr = run._r.find(qn("w:rPr"))
        spans.append((cursor, cursor + len(text), deepcopy(rpr) if rpr is not None else None))
        cursor += len(text)
    return spans


def _style_at(spans: list[tuple[int, int, Any]], position: int) -> Any:
    for start, end, rpr in spans:
        if start <= position < end:
            return deepcopy(rpr) if rpr is not None else None
    if spans:
        rpr = spans[-1][2]
        return deepcopy(rpr) if rpr is not None else None
    return None


def _append_unchanged_with_styles(parent, original: str, start: int, end: int, spans: list[tuple[int, int, Any]]) -> None:
    if end <= start:
        return
    cursor = start
    for span_start, span_end, rpr in spans:
        overlap_start = max(cursor, span_start)
        overlap_end = min(end, span_end)
        if overlap_end <= overlap_start:
            continue
        if overlap_start > cursor:
            _append_plain_run(parent, original[cursor:overlap_start], None)
        _append_plain_run(parent, original[overlap_start:overlap_end], rpr)
        cursor = overlap_end
        if cursor >= end:
            break
    if cursor < end:
        _append_plain_run(parent, original[cursor:end], _style_at(spans, cursor))


def _append_plain_run(parent, text: str, rpr: Any = None, *, deleted: bool = False) -> None:
    if not text:
        return
    run = OxmlElement("w:r")
    if rpr is not None:
        run.append(deepcopy(rpr))
    text_node = OxmlElement("w:delText" if deleted else "w:t")
    text_node.set(qn("xml:space"), "preserve")
    text_node.text = text
    run.append(text_node)
    parent.append(run)


def _append_revision(parent, text: str, *, kind: str, change_id: int, rpr: Any = None) -> None:
    if not text:
        return
    wrapper = OxmlElement("w:del" if kind == "delete" else "w:ins")
    wrapper.set(qn("w:id"), str(change_id))
    wrapper.set(qn("w:author"), "Patent Atölyesi")
    wrapper.set(qn("w:date"), datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"))
    _append_plain_run(wrapper, text, rpr, deleted=(kind == "delete"))
    parent.append(wrapper)


def _rewrite_paragraph_with_changes(paragraph, operations: list[dict[str, Any]], *, track_changes: bool, id_start: int) -> int:
    original = paragraph.text
    if not original:
        raise ValueError("Revizyon uygulanacak paragraf boş.")

    spans: list[tuple[int, int, dict[str, Any]]] = []
    for op in operations:
        old_text = str(op.get("old_text", "")).strip()
        span = _find_relaxed_phrase_span(original, old_text)
        if span is None:
            raise ValueError(f"Markup için eski ifade aynı paragrafta bulunamadı: {old_text[:120]}...")
        spans.append((span[0], span[1], op))
    spans.sort(key=lambda x: x[0])
    for left, right in zip(spans, spans[1:]):
        if right[0] < left[1]:
            raise ValueError("Aynı istem paragrafında birbiriyle çakışan iki revizyon önerisi bulundu.")

    style_spans = _run_style_spans(paragraph)
    p_el = paragraph._p
    for child in list(p_el):
        if child.tag != qn("w:pPr"):
            p_el.remove(child)

    cursor = 0
    change_id = id_start
    for start, end, op in spans:
        old_actual = original[start:end]
        new_text = str(op.get("new_text", ""))
        _append_unchanged_with_styles(p_el, original, cursor, start, style_spans)
        change_style = _style_at(style_spans, start)
        if track_changes:
            _append_revision(p_el, old_actual, kind="delete", change_id=change_id, rpr=change_style)
            change_id += 1
            _append_revision(p_el, new_text, kind="insert", change_id=change_id, rpr=change_style)
            change_id += 1
        else:
            _append_plain_run(p_el, new_text, change_style)
        cursor = end
    _append_unchanged_with_styles(p_el, original, cursor, len(original), style_spans)
    return change_id


def _enable_track_revisions(doc: Document) -> None:
    settings = doc.settings._element
    if settings.find(qn("w:trackRevisions")) is None:
        track = OxmlElement("w:trackRevisions")
        settings.insert(0, track)


def build_claim_revision_docx(source_docx: bytes, amendments: list[dict[str, Any]], *, track_changes: bool) -> bytes:
    if not amendments:
        raise ValueError("Uygulanacak istem revizyonu bulunamadı.")
    doc = Document(io.BytesIO(source_docx))
    claim_map = _claim_paragraph_map(doc)

    assignments: dict[int, list[dict[str, Any]]] = {}
    for op in amendments:
        claim_no = str(op.get("claim_number", "")).strip()
        old_text = str(op.get("old_text", "")).strip()
        candidates: list[int] = []
        for idx, paragraph in enumerate(doc.paragraphs):
            if claim_map.get(idx) != claim_no:
                continue
            if _find_relaxed_phrase_span(paragraph.text, old_text) is not None:
                candidates.append(idx)
        if len(candidates) != 1:
            if not candidates:
                raise ValueError(
                    f"İstem {claim_no} için Markup uygulanacak ifade ilgili istem paragrafında bulunamadı: {old_text[:120]}..."
                )
            raise ValueError(
                f"İstem {claim_no} için aynı eski ifade birden fazla paragrafta bulundu. Revizyon ifadesini daha özgül hale getirin."
            )
        assignments.setdefault(candidates[0], []).append(op)

    if track_changes:
        _enable_track_revisions(doc)
    change_id = 1
    for paragraph_index in sorted(assignments):
        change_id = _rewrite_paragraph_with_changes(
            doc.paragraphs[paragraph_index],
            assignments[paragraph_index],
            track_changes=track_changes,
            id_start=change_id,
        )

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def build_claim_revision_pair(source_docx: bytes, amendments: list[dict[str, Any]]) -> tuple[bytes, bytes]:
    markup = build_claim_revision_docx(source_docx, amendments, track_changes=True)
    clean = build_claim_revision_docx(source_docx, amendments, track_changes=False)
    return markup, clean


# -----------------------------------------------------------------------------
# TİP 3 ÖN ARAŞTIRMA MODÜLÜ
# -----------------------------------------------------------------------------
def top10_research_prompt(bbf_text: str, cutoff_date: str) -> str:
    return f"""{ARASTIRMA_RULES}
Aşağıdaki BBF için araştırma kesim tarihi {cutoff_date} olacak şekilde global patent araştırması yap ve en benzer tam 10 patent dokümanını belirle.
Google Patents, Espacenet, PATENTSCOPE, TÜRKPATENT ve ulaşılabilir resmi/yarı resmi patent kaynaklarını kapsayacak geniş web araştırması yap.
Dokümanları teknik yakınlığa göre sırala. Numara, başlık, tarih ve kaynak URL doğrulanmış olsun. JSON dışında yazma.
ŞEMA:
{{
 "subject_title":"",
 "technical_problem":"",
 "technical_effects":[""],
 "technical_features":[""],
 "method_steps":[{{"number":"1001","text":""}}],
 "keywords":[""],
 "ipc_cpc":[""],
 "documents":[{{
   "rank":1,"publication_number":"","application_number":"","title":"","date":"","jurisdiction":"","source_url":"",
   "summary":"","matching_features":[""],"missing_features":[""],"novelty_destroying":false,"novelty_reason":"","relevance_score":0
 }}],
 "totalpatent_query":"TotalPatent arama sorgusu: CN... or US...",
 "proposed_d1":"publication_number",
 "proposed_d2":"publication_number veya boş",
 "preliminary_novelty":"sağlanır/sağlanmaz",
 "preliminary_inventive_step":"sağlanır/sağlanmaz/belirsiz"
}}
BBF:\n{bbf_text}"""


def final_selection_prompt(
    bbf_text: str,
    top10: dict[str, Any],
    user_docs_text: str,
    decision_mode: str,
) -> str:
    return f"""{ARASTIRMA_RULES}
Aşağıdaki BBF, sistemin bulduğu 10 doküman ve kullanıcının varsa yüklediği dokümanları birlikte incele.
Nihai D1 ve gerekiyorsa D2'yi seç. Kullanıcı dokümanı daha yakınsa D1/D2'yi değiştir.
Tek doküman bütün esas teknik özellikleri doğrudan ve açık açıklıyorsa D1 ile yenilik sağlanmaz sonucuna git ve D2 seçme.
Aksi halde D1 ve tamamlayıcı D2 ile buluş basamağını değerlendir.
Kullanıcı sonuç modu: {decision_mode}
- 'Otomatik belirle' seçildiyse sonucu teknik analize göre belirle.
- Diğer seçeneklerde raporun sonuç yönünü kullanıcı seçimine göre kur; fakat kaynakta bulunmayan teknik özellik uydurma.
- D1 ve D2 karşılaştırma satırları aynı özellik listesine ve aynı sıraya sahip olmalıdır.
- status alanında yalnızca '+' veya '-' kullan; '±' kullanma.
JSON dışında yazma.
ŞEMA:
{{
 "d1":{{"number":"","title":"","date":"","source":"system/user","summary":"","abstract":"","figure_description":""}},
 "d2":null,
 "novelty_result":"sağlanır/sağlanmaz",
 "inventive_step_result":"sağlanır/sağlanmaz",
 "novelty_reasoning":[""],
 "inventive_step_reasoning":[""],
 "feature_list":[""],
 "comparison_rows_d1":[{{"feature":"","status":"+/-","evidence":""}}],
 "comparison_rows_d2":[{{"feature":"","status":"+/-","evidence":""}}],
 "helper_documents":[{{"number":"","title":"","role":""}}],
 "warnings":[""]
}}
BBF:\n{bbf_text}\n
TOP10:\n{json.dumps(top10, ensure_ascii=False, indent=2)}\n
KULLANICI DOKÜMANLARI:\n{user_docs_text}"""


def validate_research_selection(selection: dict[str, Any]) -> None:
    allowed = {"+", "-"}
    d1_rows = selection.get("comparison_rows_d1") or []
    d2_rows = selection.get("comparison_rows_d2") or []
    for label, rows in (("D1", d1_rows), ("D2", d2_rows)):
        for row in rows:
            status = str(row.get("status", "")).strip()
            if status not in allowed:
                raise ValueError(f"{label} karşılaştırma tablosunda yalnızca + veya - kullanılmalıdır: {status!r}")
    if selection.get("d2"):
        d1_features = [re.sub(r"\s+", " ", str(row.get("feature", ""))).strip() for row in d1_rows]
        d2_features = [re.sub(r"\s+", " ", str(row.get("feature", ""))).strip() for row in d2_rows]
        if d1_features != d2_features:
            raise ValueError("D1 ve D2 karşılaştırma tablolarındaki teknik özellik listeleri birebir aynı değildir.")


def report_drafting_prompt(
    bbf_text: str,
    top10: dict[str, Any],
    selection: dict[str, Any],
    reference: str,
    cutoff_date: str,
    decision_mode: str,
) -> str:
    return f"""{ARASTIRMA_RULES}
Aşağıdaki verilere göre Ön Araştırma Raporu_181612 biçiminde ayrıntılı rapor içeriği oluştur.
DP referans numarası: {reference}
Araştırma kesim tarihi: {cutoff_date}
Kullanıcı sonuç modu: {decision_mode}
JSON dışında yazma.
ŞEMA:
{{
 "reference":"{reference}",
 "title":"",
 "report_date":"{date.today().strftime('%d.%m.%Y')}",
 "purpose":"Belirlenen konuda araştırmanın gerçekleştirilmesi",
 "scope":"Global (Araştırma kesim tarihine kadar ilan edilmiş patent başvuruları)",
 "cutoff_date":"{cutoff_date}",
 "keywords":[""],
 "ipc_cpc":[{{"code":"","description":""}}],
 "evaluation_intro":"",
 "novelty_heading":"2.1. Yenilik Değerlendirmesi",
 "documents":[{{
   "label":"D1","number":"","alternate_number":"","title":"","date":"",
   "description":["2-3 cümle"],"abstract":"","figure_caption":"D1- Şekil",
   "comparison_rows":[{{"feature":"","status":"+/-"}}],
   "novelty_assessment":["5-10 satırlık değerlendirme"]
 }}],
 "inventive_step_paragraphs":[""],
 "conclusion_paragraphs":[""],
 "warnings":[""],
 "attachments":["Benzer Dokümanlar","Ön İnceleme Raporu","Makine Tercümeleri"]
}}
ÖZEL:
- selection.d2 null ise yalnızca D1 bölümü oluştur.
- D1 ve D2 tablolarındaki feature alanları birebir aynı ve aynı sırada olmalıdır.
- status alanında yalnızca '+' veya '-' kullan.
- Sonuçta yenilik ve buluş basamağı sonucunu açıkça yaz.
- Yardımcı dokümanları yalnızca buluş basamağına destek olarak kullan.
- Şablonda olmayan 'İncelenen diğer yakın dokümanlar...' benzeri cümle ekleme.
- Metni ikinci kez kontrol edip düzelt.
BBF:\n{bbf_text}\n
TOP10:\n{json.dumps(top10, ensure_ascii=False, indent=2)}\n
NİHAİ SEÇİM:\n{json.dumps(selection, ensure_ascii=False, indent=2)}"""


def build_research_docx(report: dict[str, Any]) -> bytes:
    doc = Document(str(ARASTIRMA_TEMPLATE))
    clear_body(doc)
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(1.8)
        section.left_margin = Cm(1.7)
        section.right_margin = Cm(1.7)

    # Kapak
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("DP Referans Numarası\n" + report.get("reference", ""))
    r.bold = True
    r.font.name = "Arial"
    r.font.size = Pt(12)
    doc.add_paragraph()
    add_text(doc, "ÖN ARAŞTIRMA RAPORU", bold=True, center=True, size=14)
    doc.add_paragraph()
    add_text(doc, report.get("title", ""), bold=True, center=True, size=13)
    doc.add_paragraph()
    add_text(doc, "RAPOR İÇERİĞİ:", bold=True)
    for item in ["1. ÖN ARAŞTIRMA KRİTERLERİ", "2. PATENTLENEBİLİRLİK DEĞERLENDİRMESİ", "2.1 Yenilik Değerlendirmesi", "2.2 Buluş Basamağı Değerlendirmesi", "3. SONUÇ"]:
        add_text(doc, item)
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run("Rapor Düzenleme Tarihi:\n" + report.get("report_date", ""))
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(11)
    doc.add_page_break()

    add_heading(doc, "1. ÖN ARAŞTIRMA KRİTERLERİ")
    criteria = doc.add_table(rows=5, cols=2)
    criteria.alignment = WD_TABLE_ALIGNMENT.LEFT
    rows = [
        ("Amaç", report.get("purpose", "")),
        ("Konu", report.get("title", "")),
        ("Kapsam", report.get("scope", "")),
        ("Anahtar Kelimeler", "\n".join(report.get("keywords") or [])),
        ("IPC/CPC Kodu", "\n".join(f"{x.get('code','')}: {x.get('description','')}" for x in report.get("ipc_cpc") or [])),
    ]
    for row, (label, value) in zip(criteria.rows, rows):
        set_cell_text(row.cells[0], label, bold=True)
        set_cell_text(row.cells[1], value)
    add_text(doc, "Araştırma kapsamının belirlenmesi aşamasında tarafımıza ulaştırılan ‘buluş bilgileri’ temel alınmış ve yukarıda belirtilen anahtar kelimeler ile türevleri kullanılarak sorgulama yapılmıştır. Ayrıca konu ile ilgili olarak belirlenen anahtar kelimeler, uluslararası IPC ve CPC patent sınıflandırmasına ait kodlamalar ile uyumlu hale getirilerek belirlenen teknik kapsam içerisinde araştırma gerçekleştirilmiştir.", italic=True)

    add_heading(doc, "2. DEĞERLENDİRME")
    add_text(doc, report.get("evaluation_intro", ""))
    add_heading(doc, report.get("novelty_heading", "2.1. Yenilik Değerlendirmesi"))
    for d in report.get("documents") or []:
        header = f"{d.get('label','')}- {d.get('number','')}"
        if d.get("alternate_number"):
            header += f" ({d.get('alternate_number')})"
        header += f"- {d.get('title','')}- {d.get('date','')}"
        add_heading(doc, header)
        for p in d.get("description") or []:
            add_text(doc, p)
        add_heading(doc, f"{d.get('label','')}-Özet")
        add_text(doc, d.get("abstract", ""))
        add_heading(doc, d.get("figure_caption", f"{d.get('label','')}- Şekil"))
        add_text(doc, "[İlgili patent şekli mevcutsa buraya eklenir.]")
        add_text(doc, f"Araştırma konusu ile {d.get('label','')} dokümanı arasında benzerlik değerlendirmesi:", italic=True)
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_cell_text(table.rows[0].cells[0], "Araştırma konusu", bold=True)
        set_cell_text(table.rows[0].cells[1], f"{d.get('label','')} Dokümanı", bold=True)
        for row in d.get("comparison_rows") or []:
            cells = table.add_row().cells
            set_cell_text(cells[0], row.get("feature", ""))
            set_cell_text(cells[1], row.get("status", row.get("status_evidence", "")), bold=True)
        add_text(doc, "(+) İlgili özelliğin benzer dokümanda yer aldığını ifade etmektedir.\n(-) İlgili özelliğin benzer dokümanda yer almadığını ifade etmektedir.", size=9)
        for p in d.get("novelty_assessment") or []:
            add_text(doc, p)

    add_heading(doc, "2.2. Buluş Basamağı Değerlendirmesi")
    for p in report.get("inventive_step_paragraphs") or []:
        add_text(doc, p)
    add_heading(doc, "3. SONUÇ")
    for p in report.get("conclusion_paragraphs") or []:
        add_text(doc, p)
    if report.get("warnings"):
        add_heading(doc, "Uyarılar")
        for warning in report["warnings"]:
            add_bullet(doc, warning, symbol="▪")
    add_text(doc, "Bilgilerinize sunar, çalışmalarınızda başarılar dileriz.")
    add_text(doc, "Saygılarımızla,")
    doc.add_page_break()
    add_heading(doc, "Ekler:")
    for item in report.get("attachments") or []:
        add_bullet(doc, item)
    add_heading(doc, "Önemli Not:")
    notes = [
        "Ön araştırma; rapor düzenleme tarihine kadar yayınlanan patent/faydalı model müracaatlarını kapsar.",
        "Raporlanan dokümanlar, aynı teknik alandaki birçok patent başvurusundan kapsamın olabildiğince geniş tutulması kaydıyla daraltılarak incelemenize sunulan patentlerdir.",
        "Orijinallerine ek olarak sunulan makine tercümelerinin anlaşılırlığı ve güvenirliği konusunda güvence verilemez. Ticari/hukuki kritik kararlar bunlara dayandırılmamalıdır.",
        "Bazı ülkelerin resmi veritabanlarını güncellememesinden kaynaklanabilecek eksiklikler olabilir.",
        "Çalışmamız resmi nitelikli bir araştırma değildir. Resmi patent araştırmaları ülkelerin patent ofisleri nezdinde yapılabilmektedir.",
    ]
    for i, note in enumerate(notes, 1):
        add_text(doc, f"{i}. {note}", size=9)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# -----------------------------------------------------------------------------
# ARAYÜZ
# -----------------------------------------------------------------------------
st.set_page_config(page_title=f"Patent Atölyesi {APP_VERSION}", page_icon="⚙️", layout="wide")
st.markdown(
    f"""
    <style>
      .block-container {{max-width: 1180px; padding-top: 1.5rem; padding-bottom: 3rem;}}
      .hero {{padding: 1.2rem 1.4rem; border:1px solid #e7e7e7; border-radius:16px; margin-bottom:1rem;}}
      .hero h1 {{margin:0; font-size:2rem;}}
      .hero p {{margin:.35rem 0 0 0; color:#666;}}
      .version {{font-size:.82rem; color:#888; margin-top:.45rem;}}
      div[data-testid="stDownloadButton"] button, div[data-testid="stFormSubmitButton"] button {{width:100%;}}
    </style>
    <div class="hero">
      <h1>Patent Atölyesi {APP_VERSION}</h1>
      <p>Tarifname, görüş ve Tip 3 ön araştırma çalışmalarını tek arayüzden oluşturun.</p>
      <div class="version">Kural sürümü: {RULESET_VERSION}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander("Bu sürümde kaydedilmiş temel kurallar"):
    st.markdown(
        """
        - BBF ve mevcut revize tarifnamedeki tüm teknik içerik korunur; formül, tablo, deneysel sonuç ve referans tablosu atlanmaz.
        - Her buluş için sistem/yöntem istem yapısı ayrı analiz edilir.
        - Paralel birinci–ikinci–k'ıncı işlemler ana istemde kapsayıcı yazılabilir; ayrıntıları tek bağımlı istemde toplanabilir.
        - Eğitim/genel aşama ile test aşaması paralel fakat ayrı teknik akışlar olarak değerlendirilir.
        - Referans listesi, detaylı yöntem adımları ve istemlerdeki numaralar birlikte kontrol edilir.
        - Tarifname metninde BBF/kaynak-form atfı kullanılmaz; “Buluşun bir yapılanmasında” terminolojisi ve yeni sayfa başlık düzeni uygulanır.
        - Görüşte ilk adım rapor analizidir; revizyon gerekiyorsa kullanıcı mutabakatı ve gerçek Track Changes/Markup aşamasından sonra görüş oluşturulur.
        - Tip 3 araştırmada tam 10 doküman, TotalPatent sorgu satırı, kullanıcı dokümanları ve D1/D2 son seçimi kullanılır.
        """
    )

if not os.getenv("OPENAI_API_KEY", "").strip():
    st.warning("OPENAI_API_KEY henüz tanımlı değil. Arayüzü inceleyebilirsiniz; üretim düğmeleri API anahtarı olmadan çalışmaz.")

work_type = st.radio(
    "İş türü",
    ["Tarifname oluşturma", "Görüş hazırlama", "Tip 3 - Ön araştırma raporu"],
    horizontal=True,
)

# TARİFNAME
if work_type == "Tarifname oluşturma":
    st.subheader("Tarifname oluşturma / revizyon")
    with st.form("tarifname_form"):
        c1, c2 = st.columns(2)
        with c1:
            bbf = st.file_uploader("BBF dosyası", type=["docx", "doc", "pdf", "txt"], key="tar_bbf")
            reference = st.text_input("DP referans numarası", value="")
            output_name = st.text_input("Çıktı dosyasının adı", value="Tarifname_XXXXXX.docx")
        with c2:
            current_draft = st.file_uploader(
                "Mevcut/revize tarifname (varsa)",
                type=["docx", "doc", "pdf", "txt"],
                key="tar_current_draft",
                help="Varsa doğrudan yükleyin; ayrıca Var/Yok sorusu sorulmaz. Kullanıcı düzenlemeleri korunur.",
            )
            claim_choice = st.selectbox(
                "İstem yapısı",
                ["BBF'ye göre otomatik belirle", "Yalnızca sistem", "Yalnızca yöntem", "Sistem ve yöntem"],
            )

        extra_technical_files = st.file_uploader(
            "Ek teknik müşteri belgeleri/notları (varsa)",
            type=["pdf", "docx", "doc", "txt", "md", "png", "jpg", "jpeg", "webp", "zip"],
            accept_multiple_files=True,
            key="tar_extra_technical",
        )
        example_files = st.file_uploader(
            "Örnek tarifnameler (yalnızca unsur/istem kurgusu için)",
            type=["pdf", "docx", "doc", "txt", "zip"],
            accept_multiple_files=True,
            key="tar_examples",
            help="Bu dosyaların teknik içeriği yeni tarifnameye aktarılmaz.",
        )

        separate_figures = st.checkbox("Şekilleri ayrı Word dosyası olarak oluştur")
        fc1, fc2 = st.columns(2)
        with fc1:
            figures_output_name = st.text_input("Şekiller dosyasının adı", value="Şekiller_XXXXXX.docx", disabled=not separate_figures)
        with fc2:
            figure_files = st.file_uploader(
                "Ayrıca kullanılacak şekil dosyaları",
                type=["png", "jpg", "jpeg", "webp", "docx", "pdf"],
                accept_multiple_files=True,
                key="tar_figures",
                disabled=not separate_figures,
            )

        literature = st.checkbox("Literatür araştırması yap ve önceki tekniğe ekle")
        lc1, lc2 = st.columns(2)
        with lc1:
            lit_count = st.number_input("Benzer patent sayısı", min_value=1, max_value=10, value=2, disabled=not literature)
        with lc2:
            jurisdiction = st.text_input("Tercih edilen ülke/veri tabanı", disabled=not literature)

        submit = st.form_submit_button("Tarifnameyi oluştur", type="primary")

    if submit:
        if bbf is None:
            st.error("BBF yükleyin.")
        else:
            try:
                progress = st.progress(0, text="Kaynak dosyalar okunuyor...")
                bbf_asset = UploadedAsset(bbf.name, bbf.getvalue(), bbf.type)
                source = extract_text_from_asset(bbf_asset)

                current_draft_text = ""
                current_draft_asset = None
                if current_draft is not None:
                    current_draft_asset = UploadedAsset(current_draft.name, current_draft.getvalue(), current_draft.type)
                    current_draft_text = extract_text_from_asset(current_draft_asset)

                technical_assets = assets_from_uploads(extra_technical_files)
                technical_text, technical_images = combine_asset_text("EK TEKNİK BELGE", technical_assets)
                example_assets = assets_from_uploads(example_files)
                example_text, _ = combine_asset_text("ÖRNEK TARİFNAME - YALNIZCA KURGU", example_assets)

                embedded_images = extract_embedded_images(bbf_asset)
                if current_draft_asset:
                    embedded_images.extend(extract_embedded_images(current_draft_asset))
                model_images = [*technical_images, *embedded_images][:12]

                progress.progress(15, text="BBF içeriği, referans tablosu ve istem çekirdeği çıkarılıyor...")
                extracted = ask_json(
                    tarifname_extraction_prompt(source, current_draft_text, technical_text, example_text),
                    images=model_images,
                )

                progress.progress(28, text="İstem yapısı belirleniyor...")
                mode = claim_choice
                if mode == "BBF'ye göre otomatik belirle":
                    recommended = str(extracted.get("recommended_claim_mode", "")).strip()
                    if recommended in {"Yalnızca sistem", "Yalnızca yöntem", "Sistem ve yöntem"}:
                        mode = recommended
                    elif extracted.get("has_system_basis") and extracted.get("has_method_basis"):
                        mode = "Sistem ve yöntem"
                    elif extracted.get("has_method_basis"):
                        mode = "Yalnızca yöntem"
                    else:
                        mode = "Yalnızca sistem"
                st.info(f"Kullanılan istem yapısı: {mode}")

                lit_docs: list[dict[str, Any]] = []
                if literature:
                    progress.progress(38, text="Patent literatürü araştırılıyor...")
                    lit_docs = (
                        ask_json(
                            tarifname_literature_prompt(extracted, int(lit_count), jurisdiction),
                            web_search=True,
                        ).get("documents")
                        or []
                    )

                progress.progress(55, text="Tarifname ve istemlerin tam taslağı hazırlanıyor...")
                draft = ask_json(
                    tarifname_drafting_prompt(
                        extracted,
                        mode,
                        lit_docs,
                        source,
                        current_draft_text,
                        technical_text,
                        example_text,
                    ),
                    images=model_images,
                )

                progress.progress(73, text="BBF ile tamlık ve istem tutarlılığı ikinci kez kontrol ediliyor...")
                draft = ask_json(
                    tarifname_quality_prompt(
                        source,
                        current_draft_text,
                        technical_text,
                        extracted,
                        draft,
                        mode,
                    ),
                    images=model_images,
                )

                if mode == "Yalnızca sistem":
                    draft["method_claim"] = None
                    draft["dependent_method_claims"] = []
                    draft["method_steps"] = []
                elif mode == "Yalnızca yöntem":
                    draft["system_claim"] = None
                    draft["dependent_system_claims"] = []

                draft = apply_tarifname_house_style(draft, mode)
                warnings = validate_tarifname_draft(draft, mode)
                for warning in warnings:
                    st.warning(warning)

                progress.progress(88, text="Word dosyası hazırlanıyor...")
                data = build_tarifname_docx(draft)

                figure_data = None
                if separate_figures:
                    all_figure_assets: list[UploadedAsset] = []
                    for uploaded in figure_files or []:
                        asset = UploadedAsset(uploaded.name, uploaded.getvalue(), uploaded.type)
                        all_figure_assets.extend(extract_embedded_images(asset))
                    if not all_figure_assets:
                        all_figure_assets = embedded_images
                    # Aynı görselin birden fazla kez eklenmesini engelle.
                    deduplicated: list[UploadedAsset] = []
                    seen_images: set[int] = set()
                    for asset in all_figure_assets:
                        marker = hash(asset.data)
                        if marker not in seen_images:
                            seen_images.add(marker)
                            deduplicated.append(asset)
                    figure_data = build_figures_docx(deduplicated)

                progress.progress(100, text="Hazır")
                st.success("Tarifname oluşturuldu ve BBF tamlık kontrolü tamamlandı.")
                st.download_button(
                    "Tarifname Word dosyasını indir",
                    data=data,
                    file_name=safe_output_name(output_name, "Tarifname.docx"),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                )
                if figure_data is not None:
                    st.download_button(
                        "Şekiller Word dosyasını indir",
                        data=figure_data,
                        file_name=safe_output_name(figures_output_name, "Şekiller.docx"),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
            except Exception as exc:
                st.exception(exc)

# GÖRÜŞ
elif work_type == "Görüş hazırlama":
    st.subheader("Görüş hazırlama")
    st.caption("Akış: dosyaları yükle → raporu analiz et → gerekiyorsa istem revizyonunda mutabakat/Markup → görüşü oluştur.")

    report_type = st.selectbox("Görüş türü", ["Araştırma raporuna karşı görüş", "İnceleme raporuna karşı görüş"])
    report_file = st.file_uploader("Araştırma / inceleme raporu", type=["pdf", "docx", "doc", "txt"], key="gor_report")

    prior_file = None
    if report_type == "İnceleme raporuna karşı görüş":
        prior_file = st.file_uploader("Önceki sunulan görüş", type=["pdf", "docx", "doc", "txt"], key="gor_prior")

    customer_yes = st.radio("Müşteriden bilgi var mı?", ["Hayır", "Evet"], horizontal=True)
    customer_files = []
    if customer_yes == "Evet":
        customer_files = st.file_uploader(
            "Müşteri bilgileri",
            type=["pdf", "docx", "doc", "txt", "png", "jpg", "jpeg", "webp", "zip"],
            accept_multiple_files=True,
            key="gor_customer",
        )

    spec_file = st.file_uploader("Tarifname", type=["pdf", "docx", "doc", "txt"], key="gor_spec")
    similar_files = st.file_uploader(
        "Rapordaki X/Y benzer dokümanlar (D1, D2 vb.)",
        type=["pdf", "docx", "doc", "txt", "zip"],
        accept_multiple_files=True,
        key="gor_sim",
    )
    reference = st.text_input("Ana dosya referansı nedir?", value="")
    output_name = st.text_input("Çıktı dosyasının adı", value="Görüş Metni_XXXXXX.docx")

    for key, default in {
        "gorus_analysis": None,
        "gorus_source": None,
        "gorus_markup_data": None,
        "gorus_clean_data": None,
        "gorus_final_spec_text": None,
        "gorus_opinion_data": None,
        "gorus_opinion_status": None,
    }.items():
        if key not in st.session_state:
            st.session_state[key] = default

    if st.button("1. Raporu analiz et", type="primary", use_container_width=True):
        if not all([reference.strip(), report_file, spec_file]) or not similar_files:
            st.error("Referans, rapor, tarifname ve X/Y dokümanlarını yükleyin.")
        elif report_type == "İnceleme raporuna karşı görüş" and prior_file is None:
            st.error("İnceleme raporu için önceki sunulan görüşü yükleyin.")
        elif customer_yes == "Evet" and not customer_files:
            st.error("Müşteriden bilgi var seçildi; müşteri bilgilerini yükleyin.")
        else:
            try:
                progress = st.progress(0, text="Dosyalar okunuyor...")
                report_text = extract_text_from_asset(UploadedAsset(report_file.name, report_file.getvalue(), report_file.type))
                spec_bytes = spec_file.getvalue()
                spec_text = extract_text_from_asset(UploadedAsset(spec_file.name, spec_bytes, spec_file.type))
                prior_text = ""
                if prior_file:
                    prior_text = extract_text_from_asset(UploadedAsset(prior_file.name, prior_file.getvalue(), prior_file.type))
                sim_assets = assets_from_uploads(similar_files)
                sim_text, sim_images = combine_asset_text("BENZER DOKÜMAN", sim_assets)
                cust_assets = assets_from_uploads(customer_files)
                cust_text, cust_images = combine_asset_text("MÜŞTERİ BİLGİSİ", cust_assets)
                model_images = [*sim_images, *cust_images]

                progress.progress(35, text="Rapor itirazları, X/Y dokümanları ve mevcut istemler analiz ediliyor...")
                analysis = ask_json(
                    gorus_analysis_prompt(
                        report_type,
                        reference,
                        report_text,
                        spec_text,
                        prior_text,
                        sim_text,
                        cust_text,
                    ),
                    images=model_images,
                )
                validate_gorus_analysis(analysis, spec_text)

                st.session_state.gorus_analysis = analysis
                st.session_state.gorus_source = {
                    "report_type": report_type,
                    "reference": reference,
                    "output_name": output_name,
                    "report_text": report_text,
                    "spec_text": spec_text,
                    "spec_name": spec_file.name,
                    "spec_bytes": spec_bytes,
                    "prior_text": prior_text,
                    "sim_text": sim_text,
                    "cust_text": cust_text,
                    "model_images": model_images,
                }
                st.session_state.gorus_markup_data = None
                st.session_state.gorus_clean_data = None
                st.session_state.gorus_final_spec_text = None
                st.session_state.gorus_opinion_data = None
                st.session_state.gorus_opinion_status = None
                progress.progress(100, text="İlk analiz tamamlandı")
            except Exception as exc:
                st.exception(exc)

    analysis = st.session_state.gorus_analysis
    source_state = st.session_state.gorus_source

    ready_to_generate = False
    final_spec_text = None
    revision_status = ""
    opinion_step = 2

    if analysis and source_state:
        st.success("İlk rapor analizi tamamlandı.")
        if analysis.get("analysis_summary"):
            st.write(analysis.get("analysis_summary"))

        issues = analysis.get("examiner_issues") or []
        if issues:
            st.markdown("**Raporda odaklanılması gereken hususlar:**")
            for item in issues:
                st.markdown(f"- {item}")

        directions = analysis.get("defense_direction") or []
        if directions:
            st.markdown("**Önerilen savunma yönü:**")
            for item in directions:
                st.markdown(f"- {item}")

        amendment_required = bool(analysis.get("amendment_required"))
        if amendment_required:
            st.warning("Analiz, istemlerde sınırlı bir revizyonun gerekli olabileceğini belirledi. Görüş henüz oluşturulmayacak.")
            if analysis.get("amendment_reason"):
                st.write(analysis.get("amendment_reason"))

            amendments = analysis.get("amendments") or []
            for idx, amendment in enumerate(amendments, 1):
                claim_no = amendment.get("claim_number", "")
                with st.expander(f"Revizyon {idx} – İstem {claim_no}", expanded=True):
                    st.markdown(f"**Gerekçe:** {amendment.get('reason','')}")
                    if amendment.get("basis_quote"):
                        st.markdown("**Tarifname dayanağı:**")
                        st.code(amendment.get("basis_quote", ""), language=None)
                    st.markdown("**Mevcut ifade:**")
                    st.code(amendment.get("old_text", ""), language=None)
                    st.markdown("**Önerilen ifade:**")
                    st.code(amendment.get("new_text", ""), language=None)

            refine_instruction = st.text_area(
                "Revizyon önerileri için düzeltme talimatı (varsa)",
                key="gor_revision_instruction",
                placeholder="Örneğin: yalnızca istem 1'i değiştir; istem 4'e dokunma; eklenen ifadeyi tarifnamedeki şu dayanakla sınırla...",
            )
            if st.button("Revizyon önerilerini yeniden analiz et", use_container_width=True):
                if not refine_instruction.strip():
                    st.error("Yeniden analiz için düzeltme talimatını yazın.")
                else:
                    try:
                        progress = st.progress(0, text="Revizyon önerileri yeniden değerlendiriliyor...")
                        revised_analysis = ask_json(
                            gorus_revision_refine_prompt(
                                source_state["report_type"],
                                source_state["report_text"],
                                source_state["spec_text"],
                                source_state["prior_text"],
                                source_state["sim_text"],
                                source_state["cust_text"],
                                analysis,
                                refine_instruction,
                            ),
                            images=source_state.get("model_images") or [],
                        )
                        validate_gorus_analysis(revised_analysis, source_state["spec_text"])
                        st.session_state.gorus_analysis = revised_analysis
                        st.session_state.gorus_markup_data = None
                        st.session_state.gorus_clean_data = None
                        st.session_state.gorus_final_spec_text = None
                        st.session_state.gorus_opinion_data = None
                        st.session_state.gorus_opinion_status = None
                        progress.progress(100, text="Revizyon önerileri güncellendi")
                        st.rerun()
                    except Exception as exc:
                        st.exception(exc)

            decision = st.radio(
                "İstem revizyonu kararı",
                ["Henüz karar vermedim", "Önerilen revizyonları uygula", "Revizyon yapmadan mevcut istemlerle devam et"],
                key="gor_revision_decision",
            )

            if decision == "Önerilen revizyonları uygula":
                if Path(source_state.get("spec_name", "")).suffix.lower() != ".docx":
                    st.error("Gerçek Word Track Changes/Markup üretimi için tarifnameyi .docx olarak yükleyin.")
                else:
                    if st.button("2. Onaylı revizyonlardan Markup ve temiz tarifnameyi oluştur", type="primary", use_container_width=True):
                        try:
                            progress = st.progress(0, text="İstem revizyonları Word dosyasına uygulanıyor...")
                            markup_data, clean_data = build_claim_revision_pair(
                                source_state["spec_bytes"],
                                analysis.get("amendments") or [],
                            )
                            st.session_state.gorus_markup_data = markup_data
                            st.session_state.gorus_clean_data = clean_data
                            st.session_state.gorus_final_spec_text = docx_text(clean_data)
                            st.session_state.gorus_opinion_data = None
                            st.session_state.gorus_opinion_status = None
                            progress.progress(100, text="Markup ve temiz sürüm hazır")
                        except Exception as exc:
                            st.exception(exc)

                    if st.session_state.gorus_markup_data and st.session_state.gorus_clean_data:
                        ref_name = re.sub(r"[^0-9A-Za-zÇĞİÖŞÜçğıöşü_-]+", "_", source_state.get("reference", "")) or "rev"
                        c1, c2 = st.columns(2)
                        with c1:
                            st.download_button(
                                "Markup (Track Changes) tarifnameyi indir",
                                data=st.session_state.gorus_markup_data,
                                file_name=f"Düzenlenen_tarifname_track_changes_{ref_name}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True,
                            )
                        with c2:
                            st.download_button(
                                "Temiz revize tarifnameyi indir",
                                data=st.session_state.gorus_clean_data,
                                file_name=f"Düzenlenen_tarifname_temiz_{ref_name}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True,
                            )

                        confirmed = st.checkbox(
                            "Revize istemleri kontrol ettim ve bu istem setini görüş için onaylıyorum.",
                            key="gor_revised_claims_confirmed",
                        )
                        if confirmed:
                            ready_to_generate = True
                            final_spec_text = st.session_state.gorus_final_spec_text
                            revision_status = "Kullanıcı tarafından onaylanmış revize istem seti"
                            opinion_step = 3

            elif decision == "Revizyon yapmadan mevcut istemlerle devam et":
                no_revision_confirm = st.checkbox(
                    "Revizyon önerisini gördüm; mevcut istemlerle görüş hazırlanmasını onaylıyorum.",
                    key="gor_no_revision_confirmed",
                )
                if no_revision_confirm:
                    ready_to_generate = True
                    final_spec_text = source_state["spec_text"]
                    revision_status = "Kullanıcının açık kararıyla mevcut istemlerle revizyonsuz devam"
                    opinion_step = 2
        else:
            st.info(analysis.get("no_amendment_reason") or "İlk analizde istem revizyonu gerekli görülmedi. Mevcut istemlerle görüş hazırlanabilir.")
            ready_to_generate = True
            final_spec_text = source_state["spec_text"]
            revision_status = "İlk analizde revizyon gerekmiyor; mevcut istemler esas alındı"
            opinion_step = 2

        if ready_to_generate and final_spec_text:
            if st.button(f"{opinion_step}. Görüş metnini oluştur", type="primary", use_container_width=True):
                try:
                    progress = st.progress(0, text="Onaylı istem seti üzerinden görüş hazırlanıyor...")
                    opinion = ask_json(
                        gorus_prompt(
                            source_state["report_type"],
                            source_state["reference"],
                            source_state["report_text"],
                            final_spec_text,
                            source_state["prior_text"],
                            source_state["sim_text"],
                            source_state["cust_text"],
                            preanalysis=analysis,
                            revision_status=revision_status,
                        ),
                        images=source_state.get("model_images") or [],
                    )
                    validate_quotes(opinion, final_spec_text)
                    data = build_gorus_docx(opinion)
                    st.session_state.gorus_opinion_data = data
                    st.session_state.gorus_opinion_status = revision_status
                    progress.progress(100, text="Görüş metni hazır")
                except Exception as exc:
                    st.exception(exc)

        if st.session_state.gorus_opinion_data and st.session_state.gorus_opinion_status == revision_status:
            st.success("Görüş metni oluşturuldu.")
            st.download_button(
                "Word görüş metnini indir",
                data=st.session_state.gorus_opinion_data,
                file_name=safe_output_name(source_state.get("output_name") or output_name, "Görüş Metni.docx"),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True,
            )

# ARAŞTIRMA
else:
    st.subheader("Tip 3 - Ön araştırma raporu")
    c1, c2 = st.columns(2)
    with c1:
        bbf = st.file_uploader("BBF dosyası", type=["docx", "doc", "pdf", "txt"], key="res_bbf")
        reference = st.text_input("DP referans numarası", value="")
    with c2:
        output_name = st.text_input("Çıktı dosyasının adı", value="Ön Araştırma Raporu_XXXXXX.docx")
        cutoff = st.date_input("Araştırma kesim tarihi", value=date.today())

    if "top10_result" not in st.session_state:
        st.session_state.top10_result = None
        st.session_state.research_bbf_text = None
        st.session_state.research_cutoff = None

    if st.button("1. Global araştırmayı yap ve en benzer 10 dokümanı bul", type="primary", use_container_width=True):
        if bbf is None:
            st.error("BBF yükleyin.")
        else:
            try:
                progress = st.progress(0, text="BBF okunuyor ve teknik çekirdek çıkarılıyor...")
                bbf_text = extract_text_from_asset(UploadedAsset(bbf.name, bbf.getvalue(), bbf.type))
                cutoff_text = cutoff.strftime("%d.%m.%Y")
                progress.progress(20, text="Global patent veritabanlarında araştırma yapılıyor...")
                top10 = ask_json(top10_research_prompt(bbf_text, cutoff_text), web_search=True)
                docs = top10.get("documents") or []
                if len(docs) != 10:
                    raise ValueError(f"Tam 10 doküman yerine {len(docs)} doküman döndü. Araştırmayı tekrar çalıştırın.")
                st.session_state.top10_result = top10
                st.session_state.research_bbf_text = bbf_text
                st.session_state.research_cutoff = cutoff_text
                progress.progress(100, text="Araştırma tamamlandı")
            except Exception as exc:
                st.exception(exc)

    if st.session_state.top10_result:
        st.success("En benzer 10 doküman bulundu.")
        docs = st.session_state.top10_result.get("documents") or []
        totalpatent_query = st.session_state.top10_result.get("totalpatent_query") or (
            "TotalPatent arama sorgusu: " + " or ".join(str(d.get("publication_number", "")) for d in docs)
        )
        st.code(totalpatent_query, language=None)

        rows = []
        for d in docs:
            rows.append({
                "Sıra": d.get("rank"),
                "Yayın no": d.get("publication_number"),
                "Başlık": d.get("title"),
                "Tarih": d.get("date"),
                "Yakınlık": d.get("relevance_score"),
                "Yeniliği bozar mı?": "Evet" if d.get("novelty_destroying") else "Hayır",
            })
        st.dataframe(rows, use_container_width=True, hide_index=True)
        st.caption(f"Önerilen D1: {st.session_state.top10_result.get('proposed_d1','')} | Önerilen D2: {st.session_state.top10_result.get('proposed_d2','') or '-'}")

        own_docs = st.radio("Sizin araştırdığınız benzer dokümanlar var mı?", ["Hayır", "Evet"], horizontal=True)
        user_files = []
        if own_docs == "Evet":
            user_files = st.file_uploader("Sizin bulduğunuz benzer dokümanlar", type=["pdf", "zip", "docx", "doc", "txt", "png", "jpg", "jpeg", "webp"], accept_multiple_files=True, key="res_user_docs")

        decision_mode = st.selectbox(
            "Rapor sonucu",
            ["Otomatik belirle", "Buluş basamağı var", "Buluş basamağı yok"],
            help="Ön analizden sonra istediğiniz sonuç yönünü seçebilirsiniz.",
        )

        if st.button("2. Nihai D1/D2'yi belirle ve raporu oluştur", type="primary", use_container_width=True):
            if not reference.strip():
                st.error("DP referans numarasını girin.")
            elif own_docs == "Evet" and not user_files:
                st.error("Benzer dokümanları yükleyin.")
            else:
                try:
                    progress = st.progress(0, text="Kullanıcı dokümanları inceleniyor...")
                    user_assets = assets_from_uploads(user_files)
                    user_text, user_images = combine_asset_text("KULLANICI BENZER DOKÜMANI", user_assets)
                    progress.progress(25, text="Nihai D1/D2 seçiliyor...")
                    selection = ask_json(
                        final_selection_prompt(
                            st.session_state.research_bbf_text,
                            st.session_state.top10_result,
                            user_text,
                            decision_mode,
                        ),
                        images=user_images,
                    )
                    if decision_mode == "Buluş basamağı var":
                        selection["inventive_step_result"] = "sağlanır"
                    elif decision_mode == "Buluş basamağı yok":
                        selection["inventive_step_result"] = "sağlanmaz"
                    validate_research_selection(selection)
                    progress.progress(55, text="Yenilik ve buluş basamağı raporu hazırlanıyor...")
                    report = ask_json(
                        report_drafting_prompt(
                            st.session_state.research_bbf_text,
                            st.session_state.top10_result,
                            selection,
                            reference,
                            st.session_state.research_cutoff or cutoff.strftime("%d.%m.%Y"),
                            decision_mode,
                        )
                    )
                    progress.progress(85, text="Word raporu oluşturuluyor...")
                    data = build_research_docx(report)
                    progress.progress(100, text="Hazır")
                    st.success(f"Nihai D1: {selection.get('d1',{}).get('number','')} | Nihai D2: {(selection.get('d2') or {}).get('number','-')}")
                    st.info(f"Yenilik: {selection.get('novelty_result','')} | Buluş basamağı: {selection.get('inventive_step_result','')}")
                    st.download_button("Word raporunu indir", data=data, file_name=safe_output_name(output_name, "Ön Araştırma Raporu.docx"), mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")
                except Exception as exc:
                    st.exception(exc)
