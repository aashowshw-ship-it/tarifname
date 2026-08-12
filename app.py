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
from urllib.request import Request, urlopen
import html as html_lib

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

from rules import APP_VERSION, RULESET_VERSION, ARASTIRMA_RULES, ARASTIRMA_GUNCELLEME_RULES, GORUS_RULES, TARIFNAME_RULES

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


def _append_word_field(run, field_name: str) -> None:
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), field_name)
    run._r.append(fld)


def _add_figures_page_counter(section) -> None:
    """Şekiller şablonundaki `1 / 3` mantığını dinamik PAGE / NUMPAGES alanlarıyla kurar."""
    header = section.header
    p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Header tekrar üretilirse eski içeriği temizle.
    for child in list(p._element):
        p._element.remove(child)
    r1 = p.add_run()
    _append_word_field(r1, "PAGE")
    p.add_run(" / ")
    r2 = p.add_run()
    _append_word_field(r2, "NUMPAGES")
    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(10)
        run.bold = True


def _figure_dimensions_cm(data: bytes, max_width_cm: float = 16.5, max_height_cm: float = 20.5) -> tuple[float, float]:
    try:
        with Image.open(io.BytesIO(data)) as im:
            w, h = im.size
        if not w or not h:
            raise ValueError
        ratio = h / w
        width = max_width_cm
        height = width * ratio
        if height > max_height_cm:
            height = max_height_cm
            width = height / ratio
        return max(4.0, width), max(2.5, height)
    except Exception:
        return max_width_cm, 10.0


def build_figures_docx(images: list[UploadedAsset], language: str = "Türkçe") -> bytes:
    """Müşteri şekillerini değiştirmeden, dinamik sayfa düzeni ve ŞEKİL N başlıklarıyla Word'e yerleştirir."""
    if not images:
        raise ValueError("Şekiller Word dosyası için BBF içinde veya ayrıca yüklenen dosyalarda kullanılabilir görsel bulunamadı.")

    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.7)
        section.right_margin = Cm(1.7)
        section.header_distance = Cm(0.65)
        _add_figures_page_counter(section)

    # Yaklaşık kullanılabilir dikey alan. Görsel büyüklüğüne göre 1 veya daha fazla şekil aynı sayfaya yerleşebilir.
    usable_height_cm = 24.2
    used_height_cm = 0.0

    for index, asset in enumerate(images, 1):
        width_cm, height_cm = _figure_dimensions_cm(asset.data)
        block_height = height_cm + 1.25  # şekil başlığı ve minimum aralık

        if used_height_cm > 0 and used_height_cm + block_height > usable_height_cm:
            doc.add_page_break()
            used_height_cm = 0.0

        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(0)
        p_img.paragraph_format.space_after = Pt(2)
        try:
            p_img.add_run().add_picture(io.BytesIO(asset.data), width=Cm(width_cm), height=Cm(height_cm))
        except Exception as exc:
            raise ValueError(f"{asset.name} şekiller dosyasına eklenemedi.") from exc

        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(0)
        p_cap.paragraph_format.space_after = Pt(10)
        r = p_cap.add_run(f"{'FIGURE' if _english_spec(language) else 'ŞEKİL'} {index}")
        r.bold = True
        r.font.name = "Arial"
        r.font.size = Pt(11)

        used_height_cm += block_height

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


def copy_template_paragraph_with_text(doc: Document, template: Document, index: int, text: str):
    """Şablon paragraf biçimini/runs formatını koruyup metni değiştir."""
    if not (0 <= index < len(template.paragraphs)):
        return
    p_el = deepcopy(template.paragraphs[index]._p)
    text_nodes = list(p_el.iter(qn("w:t")))
    if text_nodes:
        text_nodes[0].text = text
        for node in text_nodes[1:]:
            node.text = ""
    doc._element.body.insert(-1, p_el)


def _english_spec(language: str) -> bool:
    text = str(language or "").strip().casefold().replace("\u0307", "")
    return text in {"ingilizce", "english", "en"}


def add_blank(doc: Document):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    return p


def _copy_list_properties(target, source):
    """Şablondaki gerçek Word numaralandırma/madde işareti ve girintisini kopyala."""
    src_pr = source._p.pPr
    if src_pr is None:
        return
    tgt_pr = target._p.get_or_add_pPr()
    for tag in ("w:numPr", "w:ind"):
        existing = tgt_pr.find(qn(tag))
        if existing is not None:
            tgt_pr.remove(existing)
        src = src_pr.find(qn(tag))
        if src is not None:
            tgt_pr.append(deepcopy(src))


def add_template_list_item(doc: Document, template: Document, prototype_index: int, text: str):
    p = doc.add_paragraph()
    _copy_list_properties(p, template.paragraphs[prototype_index])
    p.add_run(text)
    return format_paragraph(p)


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
  "elements":[{"number":"<source reference or default 1>","name":"","description":""}],
  "method_steps":[{"number":"<source reference or default 1001>","text":""}],
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
    technical_supplement_text: str = "",
    example_structure_text: str = "",
    language: str = "Türkçe",
) -> str:
    return f"""{TARIFNAME_RULES}
ÇIKTI TARİFNAME DİLİ: {language}. Kaynak hangi dilde olursa olsun teknik envanteri kaynak terimlerine sadık kalarak çıkar.
Aşağıdaki kaynakları yalnızca yapı ve kapsam envanteri çıkarmak için incele. Teknik metni yeniden icat etme,
kısaltma nedeniyle önemli bilgi kaybettirme ve örnek tarifnamelerin teknik içeriğini kullanma. Gömülü şekil, grafik, diyagram, ısı haritası ve görsel teknik sonuçları da kaynak olarak incele; yalnız metin çıkarımına güvenme.

KAYNAK HİYERARŞİSİ:
1. BBF: temel teknik kaynak.
2. Ek teknik müşteri belgeleri: yalnızca açık teknik dayanak olarak kullanılabilir.
3. Örnek tarifnameler: yalnızca unsur, yöntem adımı, istem ve biçim kurgusunu görmek içindir; teknik içerikleri kullanılamaz.

JSON dışında hiçbir şey yazma.
ŞEMA:
{{
 "title":"",
 "technical_field":"",
 "prior_art_inventory":["BBF'deki her ayrı önceki teknik konu ve kısıt"],
 "technical_problems":[""],
 "technical_solution":[""],
 "technical_effects":[""],
 "elements":[{{"number":"<müşteri referansı; yoksa 1>","name":"","function":"","source":"BBF/ek teknik belge"}}],
 "method_steps":[{{"number":"<müşteri referansı; yoksa 1001>","text":"","stage":"","essential":true}}],
 "formulas":[{{"label":"","expression":"","variables":[""],"role":"zorunlu/tercihli"}}],
 "tables":[{{"caption":"","headers":[""],"rows":[[""]]}}],
 "experimental_results":[""],
 "alternatives":[""],
 "use_cases":[""],
 "figures":[""],
 "figure_reference_audit":[{"figure":"Şekil 1","reference_marks":["1"],"method_marks":["S101"],"symbolic_reference_marks":["UW"],"temporary_marks":[],"notes":""}],
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

EK TEKNİK BELGELER/NOTLAR:
---
{technical_supplement_text}
---

ÖRNEK TARİFNAMELER - YALNIZCA KURGU:
---
{example_structure_text}
---
"""

def tarifname_literature_prompt(extracted: dict[str, Any], count: int, jurisdiction: str, language: str = "Türkçe") -> str:
    return f"""Aşağıdaki buluş için tam olarak {count} teknik olarak yakın patent dokümanı araştır. Web araması kullan.
Nihai tarifname dili: {language}. title_en alanında patentin doğrulanmış özgün İngilizce başlığını koru; title_tr alanını yalnız Türkçe çıktı için yardımcı çeviri olarak doldur.
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
    technical_supplement_text: str,
    example_structure_text: str,
    language: str = "Türkçe",
) -> str:
    language_instruction = (
        "Nihai tarifnamenin TAMAMI İngilizce olmalıdır. Başlıklar, teknik açıklama, istemler ve özet İngilizce yazılmalı; Türkçe patent kalıpları bırakılmamalıdır. TECHNICAL FIELD ilk paragrafı yalnız 'The invention relates to ... .' cümlesinden oluşmalı, ikinci paragraf ayrı olarak 'In particular, the invention relates to ... .' ile başlamalıdır. İngilizce bağımsız istemlerde 'comprising:'; bağımlı istemlerde 'The system according to claim X, wherein...' veya 'The method according to claim X, wherein...' yapısını kullan. Yazılım ağırlıklı buluşlarda 'an electronic device', 'a processing unit' veya 'software executed on an electronic device' gibi geniş donanımsal taşıyıcı kullan."
        if _english_spec(language) else
        "Nihai tarifnamenin TAMAMI Türkçe olmalıdır ve mevcut Türkçe tarifname kuralları aynen uygulanmalıdır."
    )
    return f"""{TARIFNAME_RULES}
{language_instruction}
Aşağıdaki kaynaklara dayanarak seçilen dilde patent tarifnamesinin TAM metnini oluştur.
İstem yapısı: {claim_mode}

KRİTİK TALİMATLAR:
- BBF'nin bütün teknik bilgilerini kullan. Uzun önceki teknik, formüller, tablolar, deneysel sonuçlar, alternatifler ve referans tablosu atlanamaz.
- Yapılandırılmış envanter yalnızca yardımcıdır. Çelişki halinde ham BBF ve açık teknik müşteri belgeleri esas alınır.
- Örnek tarifnamelerden yalnızca kurguyu öğren; teknik bilgi aktarma.
- technical_field iki paragraf olmalıdır ve paragraflar \n\n ile ayrılmalıdır. Türkçe çıktıda ilk paragraf yalnız “Buluş, ... ile ilgilidir.”, ikinci paragraf “Buluş, özellikle ...” ile; İngilizce çıktıda ilk paragraf yalnız “The invention relates to ... .”, ikinci paragraf “In particular, the invention relates to ... .” ile başlamalıdır.
- ÖNCEKİ TEKNİK'teki aynı anlatımın devamı olan “Özellikle...”, “Bununla birlikte...”, “Bu nedenle...” gibi cümleleri ayrı paragraf yapma. Patent literatürü dokümanları ise ayrı ayrı paragraf olsun.
- Türkçe tarifnamede her patent literatürü paragrafında doğrulanmış İngilizce başlık ile Türkçe başlık karşılığı birlikte yazılsın. İngilizce tarifnamede özgün İngilizce patent başlığı kullanılsın; Türkçe başlık karşılığı nihai İngilizce metne eklenmesin.
- BULUŞUN DETAYLI AÇIKLAMASI'nda numaralı sistem/cihaz unsurlarını tek tek ayrı paragraf yapma; bütün unsur açıklamalarını teknik akış içinde tek sürekli paragrafta topla. Sistem unsuru-yöntem adımı ilişkisini açıklamak için “İşlem Adımı / Gerçekleştiren Unsur / Açıklama” türü tablo oluşturma. Bu ilişkiyi modül (1), sonraki modül (2) ve ilgili yöntem adımı (1001, 1002...) arasındaki veri/işlev bağlantısını gösteren doğal teknik paragraf olarak yaz. Yalnız ham kaynakta gerçekten sayısal/deneysel veri tablosu olan tabloları tables alanında koru. Gerçekten ayrı bir yapılanma/alternatif/yöntem/çalışma prensibi ayrıca paragraf olabilir.
- Ana istemde zorunlu teknik çekirdeği kapsayıcı biçimde ver. Aynı işlemin birinci/ikinci/k'ıncı tekrarlarını ana istemde gereksiz yere ayrı satırlara bölme. Bu ayrıntıları, aynı alt akışa aitse tek bağımlı istemde topla.
- Buluş ağırlıklı olarak yazılım/algoritma/modül/birimlerden oluşuyorsa bağımsız istemleri soyut yazılım olarak bırakma. Kaynakta özel donanım zorunlu değilse geniş bir donanımsal taşıyıcı kullan. Türkçede “bir elektronik cihaz üzerinde koşturulan yazılım vasıtasıyla ...”, İngilizcede “software executed on an electronic device ...” veya eşdeğer teknik taşıyıcı dili kullanılabilir. Gereksiz sunucu, cep telefonu/phone veya kişisel bilgisayar daraltması yapma; özel donanım uydurma.
- İstemleri yalnız hedeflenen sonuç veya fonksiyonla bırakma. Özellikle bağımsız istemde teknikte uzman kişinin “nasıl gerçekleştiriliyor?” sorusuna cevap verecek şekilde, kaynakta açık dayanağı bulunduğu ölçüde işlemi yapan teknik unsur/taşıyıcıyı, kullanılan girdiyi veya önceki unsurdan gelen veriyi, teknik işlem/mekanizmayı ve ortaya çıkan teknik çıktının sonraki unsurla bağlantısını yaz. “tespit eden / dönüştüren / optimize eden / classifying / transforming / determining” gibi sonuç bildiren fiiller kaynak mekanizmayı açıklıyorsa tek başına yeterli sayılmaz. Buna karşılık tercihli uygulama ayrıntılarıyla ana istemi gereksiz daraltma.
- Bağımlı istemleri kaynakta geçen her ayrıntı için çoğaltma. Yalnız ana isteme gerçek teknik daraltma/geri çekilme konumu sağlayan seçilmiş özellikleri kullan; istem bağımlılığı ana donanımsal taşıyıcıyı zaten taşıyorsa alt istemde elektronik cihaz/yazılım ifadesini gereksiz yere tekrar etme.
- Eğitim/genel aşama ile test aşamasındaki paralel akışları aynı mantıkla fakat ayrı teknik aşamalar olarak kur.
- REFERANS NUMARALARI bölümünde müşteri tarafından sistem/cihaz unsurları veya yöntem işlem adımları için verilmiş açık referansları AYNEN koru; 10, 20..., S101..., M1... veya başka bir referans ailesini sırf standartlaştırmak için değiştirme. Sistem/cihaz modüllerinde hiç referans yoksa kaynak sırasıyla 1, 2, 3... ver. Yöntem işlem adımlarında hiç referans yoksa varsayılan 1001, 1002, 1003... ailesini kullan. Kısmen numaralandırılmış kaynakta mevcut müşteri işaretlerini koru, yalnız boş kalanlara çakışmayacak varsayılan referans ata.
- “Yöntemin gerçekleştirdiği işlem adımları aşağıdaki gibidir:” bölümü için method_steps tam ve tutarlı olsun. Kaynaktaki yöntem referansları varsa aynen korunsun; yalnız kaynakta hiç yöntem referansı yoksa 1001’den başlayan varsayılan sıra oluşturulsun. REFERANS NUMARALARI, detaylı açıklamadaki yöntem listesi ve bağımsız yöntem isteminde aynı referanslı adımın teknik metni birebir aynı olsun. Detaylı açıklamadaki ara maddeler virgülle, son madde noktayla bitsin. Bağımsız yöntem istemindeki ara adımlar virgülle bitsin, son adım noktalamasız bitsin.
- Yalnızca yöntem modunda system_claim null olmalıdır. Yalnızca sistem modunda method_claim null olmalıdır.
- Her bağımlı istem ana isteme göre gerçek bir daraltma sağlamalıdır.
- Patent literatürü yalnızca ÖNCEKİ TEKNİK bölümünde kullanılsın.
- Kullanıcıya sunulan tarifname metninde “BBF”, “buluş bildirim formu” veya kaynak dokümana atıf yapan benzer ifadeler kesinlikle bulunmasın; teknik bilgi doğrudan buluş anlatımı olarak yazılsın.
- Sistem ve yöntem istemleri birlikte oluşturuluyorsa başlık seçilen dile uygun olarak “... Sistemi ve Yöntemi” veya “... System and Method” yapısını taşısın.
- REFERANS NUMARALARI bölümünde unsur adlarında yalnızca ilk kelimenin ilk harfi büyük olsun; standart teknik kısaltmaları koru. Unsur adları cümle içinde geçtiğinde küçük harfle başlat.
- Türkçe çıktıda “Buluşun bir gerçekleştirilmesinde” ifadesini kullanma; gerekli yerde “Buluşun bir yapılanmasında” yaz ve “Mevcut buluş” kullanma. İngilizce çıktıda doğal patent dili kullan.
- Önceki teknik bölümüne ham kaynakta verilen bütün teknik arka plan, eksiklik ve problem anlatımını aktar; literatür paragrafları bunların yerine geçmez.
- ŞEKİLLERİN KISA AÇIKLAMASI kısa ve işlevsel olsun; gerekli değilse yöntem adımı numara aralığını şekil açıklamasında tekrarlama.
- Bağımlı istemlerde Türkçe çıktıda “Önceki istemlerden herhangi birine” kalıbını, İngilizce çıktıda belirsiz “any preceding claim” zincirlerini varsayılan olarak kullanma; ek özellik hangi ana unsur veya işlem adımının ayrıntısıysa doğrudan onu tanımlayan gerekli isteme bağla.
- objectives alanında Türkçe çıktı için amaç gövdesi “... sağlamaktır.” gibi tam yüklemle bitsin. İngilizce çıktı için her objective baştan sona tam bir cümle olarak yazılsın, örneğin “The main objective of the invention is to ... .”
- Müşteri şekillerini teknik kaynak olarak aynen esas al. Görseldeki gerçek referans işaretlerini sayısal unsur, yöntem adımı, sembolik referans ve geçici şekil numarası olarak ayır. Gömülü grafik/ısı haritası/diyagram üzerindeki teknik sonuçları tamlık kontrolünde dikkate al. Geçici şekil numarasını yeni unsur referansı yapma.

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
    technical_supplement_text: str,
    extracted: dict[str, Any],
    draft: dict[str, Any],
    claim_mode: str,
    literature: list[dict[str, Any]],
    language: str = "Türkçe",
) -> str:
    language_instruction = (
        "Çıktı dili İngilizcedir. Tüm bölüm metinleri, istemler ve özet İngilizce kalmalı; TECHNICAL FIELD ve İngilizce claim kalıplarını kontrol et."
        if _english_spec(language) else
        "Çıktı dili Türkçedir. Türkçe tarifname ve istem kalıplarını kontrol et."
    )
    return f"""{TARIFNAME_RULES}
{language_instruction}
Aşağıdaki tarifname taslağını kaynaklarla SATIR SATIR karşılaştır ve eksik/yanlış hususları düzelterek tam JSON'u yeniden üret.
Bu bir özetleme görevi değildir. Kaynakta olup taslakta bulunmayan her teknik bilgi geri eklenmelidir.

ZORUNLU KONTROL LİSTESİ:
1. BBF'deki önceki teknik anlatımının tamamı korunmuş mu?
2. Referans tablosundaki bütün unsurlar ve yöntem adımları var mı?
3. Farklı numaralı adımlar yanlışlıkla aynı metinle mi yazılmış? Aynı veri farklı aşamada kullanılıyorsa aşama farkı açık mı?
4. “Yöntemin gerçekleştirdiği işlem adımları” tam liste ve referans tablosuyla uyumlu mu?
5. Ana istem zorunlu çekirdeği kapsıyor mu; paralel tekrarları gereksiz yere tek tek sayıyor mu?
6. Aynı alt akışa ait paralel analizler ve çıktılar gerekiyorsa tek bağımlı istemde mi?
7. Eğitim/genel ve test aşamaları paralel fakat ayrı olarak mı kurulmuş?
8. Formüller, değişken açıklamaları, tablolar, deneysel sonuçlar, alternatifler ve teknik etkiler eksiksiz mi?
9. Seçilen istem modu ({claim_mode}) ile başlık, açıklama ve istemler tutarlı mı? Sistem ve yöntem ise başlık seçilen dile uygun biçimde “Sistemi ve Yöntemi” veya “System and Method” yapısını taşıyor mu?
10. Bağımlı istemler gerçek daraltma sağlıyor mu?
11. Kullanıcıya sunulan metinde “BBF” veya “buluş bildirim formu” gibi kaynak atfı kalmış mı? Kalmışsa doğrudan buluş anlatımına dönüştür.
12. “Buluşun bir gerçekleştirilmesinde” veya “Mevcut buluş” kalıbı var mı? Varsa “Buluşun bir yapılanmasında” / “Buluş” diline dönüştür.
13. REFERANS NUMARALARI unsur adları yalnızca ilk kelime büyük olacak biçimde mi? Cümle içindeki unsur adları küçük harfle mi başlıyor?
14. Önceki teknik kaynakta verilen bütün müşteri teknik arka planını ve eksikliklerini içeriyor mu?
15. Teknik alan seçilen dilde iki kademeli mi? Türkçede ilk paragraf yalnız “Buluş, ... ile ilgilidir.”, ikinci paragraf “Buluş, özellikle ...” ile; İngilizcede ilk paragraf yalnız “The invention relates to ... .”, ikinci paragraf “In particular, the invention relates to ... .” ile başlamalıdır.
16. ÖNCEKİ TEKNİK'te “Özellikle...”, “Bununla birlikte...”, “Bu nedenle...” gibi aynı anlatımın devamları gereksiz yere ayrı paragraf yapılmış mı? Yapılmışsa birleştir.
17. Patent literatürü başlıkları seçilen dile uygun mu? Türkçe çıktıda İngilizce özgün başlık + Türkçe karşılığı; İngilizce çıktıda özgün İngilizce başlık kullanılmalı.
18. BULUŞUN DETAYLI AÇIKLAMASI'nda numaralı unsurlar gereksiz yere ayrı ayrı paragraflara bölünmüş mü? Bölündüyse tek sürekli unsur paragrafında birleştir.
19. Detaylı açıklamadaki yöntem madde listesinde ara maddeler virgül, son madde nokta ile bitiyor mu? Kaynakta yöntem referansları verilmişse aynen korunmuş mu; kaynakta hiç referans yoksa 1001, 1002... varsayılan sırası kullanılmış mı? Bağımsız yöntem istemindeki son işlem adımı noktalamasız mı?
20. Şekil açıklamaları kısa mı ve gerekli olmayan yöntem adımı numara aralıklarını tekrarlamıyor mu?
21. Müşteri şekillerindeki gerçek unsur/yöntem/sembolik referanslar REFERANS NUMARALARI ile uyumlu mu? Geçici şekil numaraları yeni referans olarak uydurulmuş mu?
22. Şekilde kullanılan UW, UW_F, UW_PL, UW_R, UW_M gibi sembolik referansların tarifnamede açık karşılığı var mı?
23. Müşterinin sistem/cihaz ve yöntem referans işaretleri aynen korunmuş mu? Kaynakta referans verilmeyen sistem/cihaz modülleri 1, 2, 3...; referans verilmeyen yöntem adımları 1001, 1002... varsayılanıyla tamamlanmış mı?
24. “İşlem Adımı / Gerçekleştiren Unsur / Açıklama” türü açıklama tablosu oluşturulmuş mu? Varsa tabloyu kaldır ve aynı içeriği modül-referans-yöntem adımı ilişkilerini koruyan doğal teknik paragrafa dönüştür.
25. Gömülü şekil, grafik, ısı haritası ve diyagramlarda bulunan teknik sonuçlar ile açıklayıcı etiketler tamlık kontrolünde değerlendirilmiş mi?
26. ŞEKİLLERİN KISA AÇIKLAMASI içindeki Şekil 1, Şekil 2, Şekil 3... satırları aralarında boş paragraf gerektirmeyecek biçimde ardışık açıklamalar olarak verilmiş mi?
27. Buluş yazılım/algoritma ağırlıklıysa bağımsız sistem ve/veya yöntem istemi geniş bir donanımsal taşıyıcıya, tercihen elektronik cihaz üzerinde koşturulan yazılıma, açıkça dayandırılmış mı? Gereksiz sunucu/telefon/bilgisayar daraltması yapılmış mı?
28. Bağımlı istemler kaynakta geçen her ayrıntıyı ayrı isteme dönüştürmek yerine yalnız stratejik ve gerçek daraltma sağlayan özelliklerle kontrollü tutulmuş mu?
29. Bağımsız istemler yalnız “ne/sonuç” anlatımıyla mı kalıyor, yoksa teknikte uzman kişinin “nasıl gerçekleştiriliyor?” sorusuna cevap verecek biçimde kaynakta dayanaklı teknik taşıyıcı, girdi/veri, işlem mekanizması ve çıktı/ilişkiyi yeterince gösteriyor mu? Özellikle yazılım/algoritma istemlerinde yalnız elektronik cihaz demekle yetinilmiş mi, yoksa yazılımın cihaz üzerinde hangi teknik yapılar ve işlemler üzerinden sonucu ürettiği de istemden anlaşılabiliyor mu? Ana istem bu amaçla gereksiz tercihli ayrıntılarla daraltılmışsa sadeleştir.

JSON dışında hiçbir şey yazma. Çıktı, aşağıdaki şemaya tam uymalıdır:
{TARIFNAME_DRAFT_SCHEMA}

YAPILANDIRILMIŞ ENVANTER:
{json.dumps(extracted, ensure_ascii=False, indent=2)}

DOĞRULANMIŞ PATENT LİTERATÜRÜ:
{json.dumps(literature, ensure_ascii=False, indent=2)}

HAM BBF:
---
{source_text}
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


def add_numbered_claim(doc: Document, template: Document, text: str):
    """İstem numarasını şablondaki gerçek Word otomatik numaralandırmasıyla oluştur."""
    p = doc.add_paragraph()
    _copy_list_properties(p, template.paragraphs[85])
    p.add_run(_strip_claim_number(text))
    return format_paragraph(p)


def _replace_in_nested(value: Any, old: str, new: str) -> Any:
    if isinstance(value, str):
        return value.replace(old, new)
    if isinstance(value, list):
        return [_replace_in_nested(x, old, new) for x in value]
    if isinstance(value, dict):
        return {k: _replace_in_nested(v, old, new) for k, v in value.items()}
    return value


def _ensure_title_for_claim_mode(title: str, claim_mode: str, language: str = "Türkçe") -> str:
    text = str(title or "").strip()
    if claim_mode != "Sistem ve yöntem" or not text:
        return text
    if _english_spec(language):
        if "method" in text.casefold():
            return text
        if re.search(r"\bsystem\b", text, flags=re.IGNORECASE):
            return re.sub(r"\bsystem\b(?!.*\bsystem\b)", "System and Method", text, count=1, flags=re.IGNORECASE)
        return text + " System and Method"
    if "yöntem" in text.lower():
        return text
    matches = list(re.finditer(r"\bsistemi\b", text, flags=re.IGNORECASE))
    if matches:
        last = matches[-1]
        return text[: last.start()] + "Sistemi ve Yöntemi" + text[last.end() :]
    return text + " Sistemi ve Yöntemi"


_CONTINUATION_STARTERS = (
    "özellikle",
    "bununla birlikte",
    "bu nedenle",
    "ayrıca",
    "böylece",
    "bu kapsamda",
    "dolayısıyla",
)


def _merge_continuation_paragraphs(paragraphs: list[str]) -> list[str]:
    out: list[str] = []
    for raw in paragraphs or []:
        text = re.sub(r"\s+", " ", str(raw or "")).strip()
        if not text:
            continue
        if out and text.casefold().startswith(_CONTINUATION_STARTERS):
            out[-1] = out[-1].rstrip() + " " + text
        else:
            out.append(text)
    return out


def _merge_initial_element_paragraphs(draft: dict[str, Any]) -> None:
    """Detaylı açıklamadaki ilk referanslı unsur anlatımlarını tek sürekli paragrafta birleştir."""
    paragraphs = [str(x or "").strip() for x in (draft.get("detailed_paragraphs") or []) if str(x or "").strip()]
    if len(paragraphs) < 2:
        draft["detailed_paragraphs"] = paragraphs
        return
    element_numbers = {str(x.get("number", "")).strip() for x in (draft.get("elements") or []) if str(x.get("number", "")).strip()}
    if not element_numbers:
        draft["detailed_paragraphs"] = paragraphs
        return
    merged: list[str] = []
    prefix_parts: list[str] = []
    for para in paragraphs:
        lower = para.casefold()
        is_separate = lower.startswith((
            "buluşun bir yapılanmasında",
            "yöntemin gerçekleştirdiği",
            "sistemin çalışma",
            "buluşun çalışma",
            "yöntemin çalışma",
            "in one embodiment",
            "in an embodiment",
            "the method performs",
            "the operation of the system",
        ))
        element_ref = any(f"({n})" in para for n in element_numbers)
        if not merged and not is_separate and element_ref:
            prefix_parts.append(para)
            continue
        if prefix_parts:
            merged.append(" ".join(prefix_parts))
            prefix_parts = []
        merged.append(para)
    if prefix_parts:
        merged.append(" ".join(prefix_parts))
    draft["detailed_paragraphs"] = merged


def _ensure_literature_titles(draft: dict[str, Any], literature: list[dict[str, Any]] | None, language: str = "Türkçe") -> None:
    paragraphs = list(draft.get("literature_paragraphs") or [])
    for idx, doc_info in enumerate(literature or []):
        if idx >= len(paragraphs):
            break
        p = str(paragraphs[idx] or "").strip()
        en = str(doc_info.get("title_en", "") or "").strip()
        tr = str(doc_info.get("title_tr", "") or "").strip()
        if _english_spec(language):
            if not en or en in p:
                continue
            number = str(doc_info.get("application_number", "") or "").strip()
            paragraphs[idx] = f"Patent literature includes {number + ' ' if number else ''}entitled ‘{en}’. " + p
            continue
        if not en or not tr or (en in p and tr in p):
            continue
        if en in p:
            replacement = f"İngilizce başlığı “{en}” ve Türkçe karşılığı “{tr}” olan"
            patterns = [
                rf"ve\s+[“\"]{re.escape(en)}[”\"]\s+başlıklı",
                rf"[“\"]{re.escape(en)}[”\"]\s+başlıklı",
            ]
            changed = False
            for pattern in patterns:
                new_p, count = re.subn(pattern, replacement, p, count=1, flags=re.IGNORECASE)
                if count:
                    p = new_p
                    changed = True
                    break
            if not changed:
                quoted_en = f"“{en}”"
                if quoted_en in p:
                    p = p.replace(quoted_en, replacement, 1)
                else:
                    p = p.replace(en, replacement, 1)
        else:
            number = str(doc_info.get("application_number", "") or "").strip()
            lead = f"Literatürde yapılan araştırmalar sonucu {number + ' ' if number else ''}numaralı, İngilizce başlığı “{en}” ve Türkçe karşılığı “{tr}” olan patent dokümanına rastlanmıştır. "
            p = lead + p
        paragraphs[idx] = p
    draft["literature_paragraphs"] = paragraphs



def _normalize_technical_field_two_paragraphs(value: str, language: str = "Türkçe") -> str:
    """TEKNİK ALAN/TECHNICAL FIELD girişini iki zorunlu paragrafa ayır."""
    raw = str(value or "").strip()
    parts = [x.strip() for x in re.split(r"\n\s*\n", raw) if x.strip()]
    combined = " ".join(parts)
    if _english_spec(language):
        combined = re.sub(r"^The invention\s+(?!relates)", "The invention relates to ", combined, count=1, flags=re.IGNORECASE)
        match = re.match(r"^(The invention relates to .+?\.)\s*(.*)$", combined, flags=re.IGNORECASE | re.DOTALL)
        if not match:
            return raw
        first, rest = match.group(1).strip(), match.group(2).strip()
        if not rest:
            return first
        rest = re.sub(r"^(?:In particular,\s*)?the invention relates to\s*", "", rest, count=1, flags=re.IGNORECASE).strip()
        second = "In particular, the invention relates to " + (rest[:1].lower() + rest[1:] if rest else "")
        return first + "\n\n" + second
    raw = re.sub(r"^Buluş\s+(?!,)", "Buluş, ", raw, count=1, flags=re.IGNORECASE)
    parts = [x.strip() for x in re.split(r"\n\s*\n", raw) if x.strip()]
    combined = " ".join(parts)
    match = re.match(r"^(Buluş,\s+.+?ile ilgilidir\.)\s*(.*)$", combined, flags=re.IGNORECASE | re.DOTALL)
    if not match:
        return raw
    first = match.group(1).strip()
    rest = match.group(2).strip()
    if not rest:
        return first
    rest = re.sub(r"^Buluş,\s*özellikle\s*", "", rest, count=1, flags=re.IGNORECASE).strip()
    rest = re.sub(r"^Buluş\s+özellikle\s*", "", rest, count=1, flags=re.IGNORECASE).strip()
    second = "Buluş, özellikle " + (rest[:1].lower() + rest[1:] if rest else "")
    return first + "\n\n" + second


def _normalize_method_step_numbers(draft: dict[str, Any]) -> None:
    """Müşteri yöntem referanslarını koru; yalnız eksik referansları 1001... varsayılanıyla tamamla."""
    steps = draft.get("method_steps") or []
    if not steps:
        return
    used = {str(step.get("number", "") or "").strip() for step in steps if str(step.get("number", "") or "").strip()}
    next_default = 1001
    for step in steps:
        number = str(step.get("number", "") or "").strip()
        if number:
            step["number"] = number
            continue
        while str(next_default) in used:
            next_default += 1
        step["number"] = str(next_default)
        used.add(str(next_default))
        next_default += 1

    method_claim = draft.get("method_claim") or {}
    if method_claim.get("steps") is not None:
        claim_steps = list(method_claim.get("steps") or [])
        normalized: list[str] = []
        for i, step in enumerate(steps):
            text = str(step.get("text", "") or "").strip().rstrip(".,;:")
            number = str(step.get("number", "") or "").strip()
            # method_steps metni numarasızdır; istemde aynı metin + kaynak referansı kullanılır.
            normalized.append(f"{text} ({number})" if number else text)
        method_claim["steps"] = normalized
        draft["method_claim"] = method_claim


def _assign_missing_element_numbers(draft: dict[str, Any]) -> None:
    """Müşteri unsur referanslarını koru; yalnız eksik olanlara 1..N varsayılanı ata."""
    elements = draft.get("elements") or []
    if not elements:
        return
    used = {str(e.get("number", "") or "").strip() for e in elements if str(e.get("number", "") or "").strip()}
    next_default = 1
    for element in elements:
        number = str(element.get("number", "") or "").strip()
        if number:
            element["number"] = number
            continue
        while str(next_default) in used:
            next_default += 1
        element["number"] = str(next_default)
        used.add(str(next_default))
        next_default += 1


def _convert_mapping_tables_to_prose(draft: dict[str, Any], language: str = "Türkçe") -> None:
    """İşlem Adımı/Gerçekleştiren Unsur açıklama tablolarını teknik paragrafa çevir."""
    kept: list[dict[str, Any]] = []
    mapping_paragraphs: list[str] = []
    elements = draft.get("elements") or []
    steps = draft.get("method_steps") or []
    for table in draft.get("tables") or []:
        headers = [str(x or "").strip().casefold() for x in (table.get("headers") or [])]
        is_mapping = (any("işlem adımı" in h for h in headers) and any("gerçekleştiren unsur" in h for h in headers)) or (any("process step" in h for h in headers) and any("performing element" in h or "performing unit" in h for h in headers))
        if not is_mapping:
            kept.append(table)
            continue
        rows = table.get("rows") or []
        sentences: list[str] = []
        for idx, row in enumerate(rows):
            row = list(row or [])
            element = elements[idx] if idx < len(elements) else {}
            step = steps[idx] if idx < len(steps) else {}
            el_name = str(element.get("name", "") or (row[2] if len(row) > 2 else "ilgili modül")).strip()
            el_num = str(element.get("number", "") or "").strip()
            step_text = str(step.get("text", "") or (row[1] if len(row) > 1 else "ilgili işlem")).strip().rstrip(".,;:")
            step_num = str(step.get("number", "") or "").strip()
            explanation = str(row[3] if len(row) > 3 else "").strip().rstrip(".")
            if _english_spec(language):
                lead = f"The {el_name}{f' ({el_num})' if el_num else ''} performs the process step of {step_text[:1].lower() + step_text[1:]}{f' ({step_num})' if step_num else ''}"
                if explanation:
                    lead += f" and, in this context, {explanation[:1].lower() + explanation[1:]}"
            else:
                lead = f"{el_name}{f' ({el_num})' if el_num else ''}, {step_text[:1].lower() + step_text[1:]} işlem adımını{f' ({step_num})' if step_num else ''} gerçekleştirmektedir"
                if explanation:
                    lead += f" ve bu kapsamda {explanation[:1].lower() + explanation[1:]}"
            sentences.append(lead.rstrip(".") + ".")
        if sentences:
            prefix = "The technical relationship between the system and the method is established as follows. " if _english_spec(language) else "Sistem ile yöntem arasındaki teknik ilişki aşağıdaki şekilde kurulmaktadır. "
            mapping_paragraphs.append(prefix + " ".join(sentences))
    draft["tables"] = kept
    if mapping_paragraphs:
        detailed = [str(x or "").strip() for x in (draft.get("detailed_paragraphs") or []) if str(x or "").strip()]
        insert_at = 1 if detailed else 0
        for paragraph in reversed(mapping_paragraphs):
            detailed.insert(insert_at, paragraph)
        draft["detailed_paragraphs"] = detailed

def apply_tarifname_house_style(
    draft: dict[str, Any],
    claim_mode: str,
    literature: list[dict[str, Any]] | None = None,
    language: str = "Türkçe",
) -> dict[str, Any]:
    """Kullanıcının sabit terminoloji, paragraf ve başlık tercihlerini çıktıdan önce uygula."""
    draft = deepcopy(draft)
    if not _english_spec(language):
        for old, new in [
            ("Buluşun bir gerçekleştirilmesinde", "Buluşun bir yapılanmasında"),
            ("buluşun bir gerçekleştirilmesinde", "buluşun bir yapılanmasında"),
            ("Mevcut buluş", "Buluş"),
            ("mevcut buluş", "buluş"),
            ("Buluş özellikle", "Buluş, özellikle"),
        ]:
            draft = _replace_in_nested(draft, old, new)
    draft["title"] = _ensure_title_for_claim_mode(draft.get("title", ""), claim_mode, language)
    _assign_missing_element_numbers(draft)
    _normalize_method_step_numbers(draft)
    _convert_mapping_tables_to_prose(draft, language)
    draft["prior_art_general_paragraphs"] = _merge_continuation_paragraphs(draft.get("prior_art_general_paragraphs") or [])
    _merge_initial_element_paragraphs(draft)
    _ensure_literature_titles(draft, literature, language)
    draft["figure_descriptions"] = [
        re.sub(r"\b\d{4}\s*[-–]\s*\d{4}\s+numaralı\s+", "", str(x or ""), flags=re.IGNORECASE)
        for x in (draft.get("figure_descriptions") or [])
    ]
    draft["technical_field"] = _normalize_technical_field_two_paragraphs(draft.get("technical_field", ""), language)
    for step in draft.get("method_steps") or []:
        step["text"] = re.sub(r"[.,;:]+$", "", str(step.get("text", "") or "").strip())
    return draft


def validate_tarifname_draft(
    draft: dict[str, Any],
    claim_mode: str,
    literature: list[dict[str, Any]] | None = None,
    language: str = "Türkçe",
) -> list[str]:
    warnings: list[str] = []
    technical_field_raw = str(draft.get("technical_field", "") or "").strip()
    tf_paragraphs = [x.strip() for x in re.split(r"\n\s*\n", technical_field_raw) if x.strip()]
    if _english_spec(language):
        if len(tf_paragraphs) < 2:
            raise ValueError('TECHNICAL FIELD must contain two paragraphs: “The invention relates to ... .” followed by “In particular, the invention relates to ... .”')
        if not re.fullmatch(r"The invention relates to .+?\.", tf_paragraphs[0], flags=re.IGNORECASE | re.DOTALL):
            raise ValueError('The first TECHNICAL FIELD paragraph must consist only of “The invention relates to ... .”')
        if not re.match(r"^In particular,\s*the invention relates to\b", tf_paragraphs[1], flags=re.IGNORECASE):
            raise ValueError('The second TECHNICAL FIELD paragraph must begin with “In particular, the invention relates to ...”.')
    else:
        technical_field = re.sub(r"\s+", " ", technical_field_raw).strip()
        if technical_field and not re.match(r"^Buluş,\s+.+?ile ilgilidir\.", technical_field, flags=re.IGNORECASE):
            raise ValueError('TEKNİK ALAN ilk cümlesi “Buluş, ... ile ilgilidir.” yapısında olmalıdır.')
        if re.search(r"\bBuluş özellikle\b", technical_field, flags=re.IGNORECASE):
            raise ValueError('TEKNİK ALAN içinde “Buluş özellikle” yerine “Buluş, özellikle” kullanılmalıdır.')
        if len(tf_paragraphs) < 2:
            raise ValueError('TEKNİK ALAN iki paragraf olmalıdır: ilk paragraf “Buluş, ... ile ilgilidir.”, ikinci paragraf “Buluş, özellikle ...” ile başlamalıdır.')
        if not re.fullmatch(r"Buluş,\s+.+?ile ilgilidir\.", tf_paragraphs[0], flags=re.IGNORECASE | re.DOTALL):
            raise ValueError('TEKNİK ALAN ilk paragrafı yalnız “Buluş, ... ile ilgilidir.” giriş cümlesinden oluşmalıdır.')
        if not re.match(r"^Buluş,\s*özellikle\b", tf_paragraphs[1], flags=re.IGNORECASE):
            raise ValueError('TEKNİK ALAN ikinci paragrafı “Buluş, özellikle ...” ile başlamalıdır.')

    steps = draft.get("method_steps") or []
    numbers = [str(x.get("number", "")).strip() for x in steps]
    if len(numbers) != len(set(numbers)):
        raise ValueError("REFERANS NUMARALARI bölümünde yinelenen yöntem adımı numarası bulundu.")
    if any(not n for n in numbers):
        raise ValueError("Numarası boş yöntem işlem adımı bulundu.")
    element_numbers = [str(x.get("number", "") or "").strip() for x in (draft.get("elements") or [])]
    if any(not n for n in element_numbers):
        raise ValueError("Sistem/cihaz unsurlarından en az birinin referans numarası boş.")
    overlap = sorted(set(element_numbers) & set(numbers))
    if overlap:
        warnings.append("Müşteri kaynaklı sistem unsuru ve yöntem adımı referansları çakışıyor: " + ", ".join(overlap) + ". Referanslar otomatik değiştirilmedi; teknik belirsizlik açısından kontrol edin.")

    for table in draft.get("tables") or []:
        headers = [str(x or "").casefold() for x in (table.get("headers") or [])]
        is_mapping = (any("işlem adımı" in h for h in headers) and any("gerçekleştiren unsur" in h for h in headers)) or (any("process step" in h for h in headers) and any("performing element" in h or "performing unit" in h for h in headers))
        if is_mapping:
            raise ValueError("Sistem-yöntem ilişki tablosu tarifname gövdesinde tablo olarak bırakılamaz; doğal teknik paragrafa dönüştürülmelidir.")

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
    missing = [n for n in numbers if n and f"({n})" not in all_claim_text and draft.get("method_claim")]
    if missing:
        raise ValueError("REFERANS NUMARALARI bölümündeki yöntem adımlarından istemde bulunmayan referanslar: " + ", ".join(missing))

    if claim_mode == "Yalnızca yöntem" and draft.get("system_claim"):
        raise ValueError("Yalnızca yöntem seçildiği halde sistem istemi üretildi.")
    if claim_mode == "Yalnızca sistem" and draft.get("method_claim"):
        raise ValueError("Yalnızca sistem seçildiği halde yöntem istemi üretildi.")
    if claim_mode in {"Yalnızca yöntem", "Sistem ve yöntem"} and not draft.get("method_claim"):
        raise ValueError("Seçilen istem yapısına rağmen bağımsız yöntem istemi üretilemedi.")
    if claim_mode in {"Yalnızca sistem", "Sistem ve yöntem"} and not draft.get("system_claim"):
        raise ValueError("Seçilen istem yapısına rağmen bağımsız sistem istemi üretilemedi.")

    hardware_anchor_re = re.compile(r"elektronik cihaz|elektronik işlem birimi|işlemci|donanım|bilgisayar|mikrodenetleyici|kontrol birimi|electronic device|processing unit|processor|hardware|computer|microcontroller|control unit", re.IGNORECASE)
    software_terms_re = re.compile(r"modül|birim|algoritma|yazılım|veri işleme|hesaplama|module|unit|algorithm|software|data processing|calculation", re.IGNORECASE)
    if draft.get("system_claim"):
        system_text = " ".join([
            str((draft.get("system_claim") or {}).get("preamble", "")),
            *map(str, (draft.get("system_claim") or {}).get("elements") or []),
        ])
        if len(software_terms_re.findall(system_text)) >= 2 and not hardware_anchor_re.search(system_text):
            raise ValueError("Yazılım/modül ağırlıklı bağımsız sistem istemi geniş bir donanımsal taşıyıcıya dayandırılmalıdır; örneğin elektronik cihaz üzerinde koşturulan yazılım vasıtasıyla.")
    if draft.get("method_claim"):
        method_text = " ".join([
            str((draft.get("method_claim") or {}).get("preamble", "")),
            *map(str, (draft.get("method_claim") or {}).get("steps") or []),
        ])
        if len(software_terms_re.findall(method_text)) >= 2 and not hardware_anchor_re.search(method_text):
            raise ValueError("Yazılım/algoritma ağırlıklı bağımsız yöntem istemi elektronik cihaz/işlemci gibi geniş bir donanımsal taşıyıcıya dayandırılmalıdır.")

    user_facing_text = json.dumps(draft, ensure_ascii=False)
    if re.search(r"\bBBF\b|buluş bildirim formu|invention disclosure form", user_facing_text, flags=re.IGNORECASE):
        raise ValueError("Tarifname taslağında kullanıcıya görünmemesi gereken kaynak form atfı bulundu.")
    if not _english_spec(language) and re.search(r"\bmevcut buluş\b", user_facing_text, flags=re.IGNORECASE):
        raise ValueError('Tarifname taslağında “mevcut buluş” ifadesi bulundu; “Buluş” dili kullanılmalıdır.')
    if claim_mode == "Sistem ve yöntem":
        required = "method" if _english_spec(language) else "yöntem"
        if required not in str(draft.get("title", "")).casefold():
            raise ValueError("Sistem ve yöntem istem yapısında başlık yöntem/method ifadesini içermiyor.")

    literature_text = " ".join(str(x or "") for x in (draft.get("literature_paragraphs") or []))
    for doc_info in literature or []:
        en = str(doc_info.get("title_en", "") or "").strip()
        tr = str(doc_info.get("title_tr", "") or "").strip()
        if en and en not in literature_text:
            raise ValueError(f"Literatür paragrafında patent başlığı eksik: {en}")
        if not _english_spec(language) and tr and tr not in literature_text:
            raise ValueError(f"Literatür paragrafında Türkçe patent başlığı eksik: {tr}")
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


def build_tarifname_docx(draft: dict[str, Any], language: str = "Türkçe") -> bytes:
    template = Document(str(TARIFNAME_TEMPLATE))
    doc = Document(str(TARIFNAME_TEMPLATE))
    clear_body(doc)
    en = _english_spec(language)
    labels = {
        "spec": "SPECIFICATION" if en else "TARİFNAME",
        "technical": "TECHNICAL FIELD" if en else "TEKNİK ALAN",
        "prior": "PRIOR ART" if en else "ÖNCEKİ TEKNİK",
        "short": "BRIEF DESCRIPTION OF THE INVENTION" if en else "BULUŞUN KISA AÇIKLAMASI",
        "figures": "BRIEF DESCRIPTION OF THE FIGURES" if en else "ŞEKİLLERİN KISA AÇIKLAMASI",
        "refs": "REFERENCE NUMERALS" if en else "REFERANS NUMARALARI",
        "detail": "DETAILED DESCRIPTION OF THE INVENTION" if en else "BULUŞUN DETAYLI AÇIKLAMASI",
        "claims": "CLAIMS" if en else "İSTEMLER",
        "abstract": "ABSTRACT" if en else "ÖZET",
    }
    for section in doc.sections:
        section.top_margin = Cm(3)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(3)
        section.right_margin = Cm(2)

    # Şablondaki boş paragraf ritmi birebir korunur.
    add_heading(doc, labels["spec"], center=True)
    add_blank(doc)
    add_text(doc, draft.get("title", ""), bold=True, center=True)
    add_blank(doc)
    copy_template_paragraph_with_text(doc, template, 4, "For preparation of the search report, the information provided in the specification and claims should be sufficiently clear and detailed to enable a person skilled in the art to carry out the subject product/method. If questions have been raised to clarify any issue, please provide the requested information. If you consider that an important element, process step or feature has not been stated before the claims section, please indicate your comments without changing the text, using highlighted text, and return them by e-mail." if en else template.paragraphs[4].text)
    add_blank(doc)

    add_heading(doc, labels["technical"])
    add_blank(doc)
    technical_field_parts = [
        x.strip() for x in re.split(r"\n\s*\n", str(draft.get("technical_field", "") or "")) if x.strip()
    ]
    for idx, paragraph in enumerate(technical_field_parts):
        add_text(doc, paragraph)
        if idx < len(technical_field_parts) - 1:
            add_blank(doc)
    add_blank(doc)

    add_heading(doc, labels["prior"])
    add_blank(doc)
    prior = draft.get("prior_art_general_paragraphs") or []
    for paragraph in prior:
        add_text(doc, paragraph)
        add_blank(doc)
    literature = draft.get("literature_paragraphs") or []
    for paragraph in literature:
        add_text(doc, paragraph)
        add_blank(doc)
    add_text(doc, "Consequently, the problems described above, which remain unresolved in view of the prior art, have created a need for an improvement in the relevant technical field." if en else "Sonuçta yukarıda bahsedilen ve mevcut teknik ışığında çözülemeyen sorunlar, ilgili teknik alanda bir yenilik yapmayı zorunlu kılmıştır.")

    add_heading(doc, labels["short"])
    add_blank(doc)
    add_text(doc, draft.get("short_description_intro", ""))
    add_blank(doc)
    for index, objective in enumerate(draft.get("objectives") or []):
        objective = str(objective).strip()
        if en:
            # English objectives are drafted as complete sentences.
            add_text(doc, objective)
        else:
            prefix = "Buluşun ana amacı, " if index == 0 else "Buluşun diğer bir amacı, "
            add_text(doc, prefix + (objective[:1].lower() + objective[1:] if objective else ""))
        add_blank(doc)
    invention_definition = draft.get("unumbered_invention_definition") or draft.get("unumbered_system_definition")
    if invention_definition:
        add_text(doc, invention_definition)
    invention_features = draft.get("unumbered_invention_features") or draft.get("unumbered_system_elements") or []
    for item in invention_features:
        add_template_list_item(doc, template, 30, item)
    if invention_features:
        if en:
            add_text(doc, "The invention comprises the foregoing technical features." if draft.get("system_claim") else "The method comprises the foregoing process steps.")
        else:
            add_text(doc, "işlem adımlarını içermesidir." if draft.get("method_claim") and not draft.get("system_claim") else "içermesidir.")
    add_blank(doc)
    add_text(doc, "For a better understanding of the configuration and advantages of the invention together with its additional elements, the invention should be considered with reference to the figures described below." if en else "Buluşun yapılanması ve ek elemanlarla birlikte avantajlarının en iyi şekilde anlaşılabilmesi için aşağıda açıklaması yapılan şekiller ile birlikte değerlendirilmesi gerekmektedir.")
    add_blank(doc)

    add_heading(doc, labels["figures"])
    add_blank(doc)
    for figure in draft.get("figure_descriptions") or (["Figure 1 is a representative illustration of the configuration of the invention."] if en else ["Şekil 1, buluşa konu yapılanmanın temsili gösterimidir."]):
        add_text(doc, figure)
    add_blank(doc)
    add_text(doc, "The drawings are not necessarily to scale, and details that are not required for understanding the invention may be omitted. Elements that are substantially identical or perform substantially identical functions may be indicated by the same reference numeral or sign." if en else "Çizimlerin mutlaka ölçeklendirilmesi gerekmemektedir ve buluşu anlamak için gerekli olmayan detaylar ihmal edilmiş olabilmektedir. Bundan başka, en azından önemli ölçüde özdeş elemanlar veya benzer fonksiyonlara sahip olan elemanlar aynı numara ile gösterilmektedir.")
    add_blank(doc)

    add_heading(doc, labels["refs"])
    add_blank(doc)
    for element in draft.get("elements") or []:
        add_text(doc, f"{element.get('number','')}. {_reference_sentence_case(element.get('name',''))}")
    method_steps = draft.get("method_steps") or []
    if draft.get("elements") and method_steps:
        add_blank(doc)
    for step in method_steps:
        number = str(step.get("number", "") or "").strip()
        text = str(step.get("text", "") or "").strip()
        if number:
            text = re.sub(rf"\s*\(\s*{re.escape(number)}\s*\)\s*$", "", text).strip()
        text = text.rstrip(".,;:")
        add_text(doc, f"{number}. {text}")
    add_blank(doc)

    add_heading(doc, labels["detail"])
    add_blank(doc)
    title = str(draft.get("title", "buluş")).strip()
    add_text(doc, f"In this detailed description, {title} is described by way of examples solely for a better understanding of the subject matter, without imposing any limiting effect." if en else f"Bu detaylı açıklamada, buluş konusu olan {title} sadece konunun daha iyi anlaşılmasına yönelik hiçbir sınırlayıcı etki oluşturmayacak örneklerle açıklanmaktadır.")
    add_blank(doc)
    detailed = [str(x or "").strip() for x in (draft.get("detailed_paragraphs") or []) if str(x or "").strip()]
    for idx, paragraph in enumerate(detailed):
        add_text(doc, paragraph)
        add_blank(doc)

    for formula in draft.get("formulas") or []:
        if formula.get("label"):
            add_text(doc, formula.get("label", ""), bold=True)
        if formula.get("expression"):
            add_text(doc, formula.get("expression", ""), center=True)
        if formula.get("explanation"):
            add_text(doc, formula.get("explanation", ""))
        add_blank(doc)

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
        add_blank(doc)

    for paragraph in draft.get("experimental_results") or []:
        add_text(doc, paragraph)
        add_blank(doc)
    for paragraph in draft.get("alternatives") or []:
        add_text(doc, paragraph)
        add_blank(doc)

    if method_steps:
        add_text(doc, "The process steps performed by the method are as follows:" if en else "Yöntemin gerçekleştirdiği işlem adımları aşağıdaki gibidir:")
        for idx, step in enumerate(method_steps):
            number = str(step.get("number", "") or "").strip()
            text = str(step.get("text", "") or "").strip()
            if number:
                text = re.sub(rf"\s*\(\s*{re.escape(number)}\s*\)\s*$", "", text).strip()
            text = text.rstrip(".,;:")
            punctuation = "." if idx == len(method_steps) - 1 else ","
            add_template_list_item(doc, template, 30, f"{text} ({number}){punctuation}")
        add_blank(doc)
    if draft.get("working_principle"):
        add_text(doc, draft.get("working_principle", ""))
        add_blank(doc)

    # Şablonda istemler öncesinde iki boş paragraf vardır; page break yeni sayfayı garanti eder.
    add_blank(doc)
    doc.add_page_break()
    add_heading(doc, labels["claims"], center=True)
    add_blank(doc)
    if en:
        english_claim_notes = {
            79: "If you consider that an important and novel element and/or process step is missing from the claims, please indicate your comments without changing the text, using highlighted text, and return them by e-mail.",
            81: "If the system can operate without one or more features recited in claim 1, or if a feature that must necessarily be present is not recited in claim 1, identifying that feature is important for determining the desired scope of protection.",
            83: "A system that implements the invention without using even one feature recited in claim 1 may fall outside the scope of that claim. Please evaluate the claims in light of this information.",
        }
        for index in (79, 81, 83):
            copy_template_paragraph_with_text(doc, template, index, english_claim_notes[index])
            add_blank(doc)
    else:
        for index in (79, 81, 83):
            copy_template_paragraph(doc, template, index)
            add_blank(doc)

    system_claim = draft.get("system_claim")
    if system_claim:
        if en:
            add_numbered_claim(doc, template, f"{system_claim.get('preamble','').rstrip(' ,;:')}, comprising:")
            items = list(system_claim.get("elements") or [])
            for idx, item in enumerate(items):
                text = str(item).rstrip(".,;:") + ("." if idx == len(items) - 1 else ";")
                add_template_list_item(doc, template, 86, text)
        else:
            add_numbered_claim(doc, template, f"{system_claim.get('preamble','')} olup, özelliği;")
            for item in system_claim.get("elements") or []:
                add_template_list_item(doc, template, 86, item)
            closing = add_text(doc, system_claim.get("closing", "içermesidir."))
            closing.paragraph_format.first_line_indent = Cm(0.5)
        add_blank(doc)
        for dependent in draft.get("dependent_system_claims") or []:
            add_numbered_claim(doc, template, dependent)
            add_blank(doc)

    method_claim = draft.get("method_claim")
    if method_claim:
        if en:
            add_numbered_claim(doc, template, f"{method_claim.get('preamble','').rstrip(' ,;:')}, comprising:")
        else:
            add_numbered_claim(doc, template, f"{method_claim.get('preamble','')} olup, özelliği;")
        claim_steps = list(method_claim.get("steps") or [])
        for idx, item in enumerate(claim_steps):
            base_step = str(item).rstrip(".,;:")
            if en:
                step_text = base_step + ("." if idx == len(claim_steps) - 1 else ";")
            else:
                step_text = base_step if idx == len(claim_steps) - 1 else base_step + ","
            add_template_list_item(doc, template, 86, step_text)
        if not en:
            closing = add_text(doc, method_claim.get("closing", "işlem adımlarını içermesidir."))
            closing.paragraph_format.first_line_indent = Cm(0.5)
        add_blank(doc)
        for dependent in draft.get("dependent_method_claims") or []:
            add_numbered_claim(doc, template, dependent)
            add_blank(doc)

    summary_heading = add_heading(doc, labels["abstract"], center=True)
    summary_heading.paragraph_format.page_break_before = True
    add_blank(doc)
    add_text(doc, draft.get("title", ""), bold=True, center=True)
    add_blank(doc)
    add_text(doc, draft.get("abstract", ""))
    add_blank(doc)

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
Dokümanları teknik yakınlığa göre sırala. Numara, başlık, tarih ve kaynak URL doğrulanmış olsun. İlk araştırma çıktısında tam 10 dokümanla birlikte `totalpatent_query`, `proposed_d1` ve `proposed_d2` alanlarını mutlaka doldur. Her doküman için yayımlanmış özgün İngilizce Abstract metni doğrulanabiliyorsa `abstract_en` alanına doğrudan aktar; Türkçeye çevirme veya yeniden özetleme. Kullanıcı dokümanlarını bu aşamada sorma veya varsayma. JSON dışında yazma.
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
   "summary":"","abstract_en":"","matching_features":[""],"missing_features":[""],"novelty_destroying":false,"novelty_reason":"","relevance_score":0
 }}],
 "totalpatent_query":"TotalPatent arama sorgusu: CN... or US...",
 "proposed_d1":"publication_number",
 "proposed_d2":"publication_number veya boş",
 "preliminary_novelty":"sağlanır/sağlanmaz",
 "preliminary_inventive_step":"sağlanır/sağlanmaz/belirsiz"
}}
BBF:\n{bbf_text}"""


# -----------------------------------------------------------------------------
# TİP 3 - ŞABLONA SIKI SADAKAT / ARAŞTIRMA GÜNCELLEME (v5.4)
# -----------------------------------------------------------------------------
def final_selection_prompt(
    bbf_text: str,
    top10: dict[str, Any],
    user_docs_text: str,
) -> str:
    return f"""{ARASTIRMA_RULES}
Aşağıdaki araştırma konusu, sistemin ilk global araştırmada bulduğu 10 doküman ve kullanıcının varsa yüklediği dokümanları birlikte incele.
İlk 10 dokümanı DEĞİŞTİRME veya yeniden sıralanmış yeni bir top-10 listesi üretme. Kullanıcı dokümanları arasından yalnız araştırma konusuna en ilgili birkaç belgeyi `10+` ek doküman olarak seç.
Nihai D1 ve gerekiyorsa D2'yi belirle. Kullanıcı dokümanı daha yakınsa D1/D2 değişebilir.
Tek doküman bütün esas teknik özellikleri doğrudan ve açık açıklıyorsa D1 ile yenilik sağlanmaz sonucuna git ve D2 seçme. Aksi halde en yakın D1 ve tamamlayıcı D2 ile buluş basamağını değerlendir.
ÖNEMLİ: Bu aşamada kullanıcı sonuç modu henüz alınmamıştır. Sonucu kullanıcı tercihine göre zorlama. Kendi teknik kanaatini bağımsız olarak ver ve `technical_opinion` alanını mutlaka `Bence buluş basamağı var.` veya `Bence buluş basamağı yok.` cümlesiyle başlat.
- `additional_relevant_user_documents` yalnız kullanıcı tarafından yüklenen en ilgili birkaç belgeyi içersin.
- `additional_query` alanı belge varsa tam olarak `10+ CN... or IN...` biçiminde olsun; ilgili kullanıcı belgesi yoksa boş bırak.
- İlk önerilen D1/D2 ile nihai D1/D2'yi karşılaştırarak `d1_changed` ve `d2_changed` alanlarını doğru doldur; değişiklik notunda eski ve yeni numarayı açık yaz.
- D1 ve D2 karşılaştırma satırları aynı özellik listesine ve aynı sıraya sahip olmalıdır.
- Sağ hücre metni çıplak + veya - olmasın. `+ Özet; İstem 1; Şekil 3 ...` veya `- Dokümanda ... açıklanmamaktadır.` mantığında somut dayanak yaz.
- Patent şekli için model tarafından üretilmiş veya temsili bir görsel verme. Kaynakta bulunan özgün şekle ait doğrudan URL bulunabiliyorsa `figure_image_url` alanına yaz; aksi halde boş bırak.
- Nihai D1/D2 için `abstract_en` alanına yalnız ilgili patentin yayımlanmış özgün İngilizce Abstract metnini koy. Kullanıcı yüklemesinde özgün İngilizce Abstract varsa onu öncelikle aynen kullan; yoksa doğrulanmış resmi/Google Patents kaynağından getir. Türkçe özet, çeviri veya model tarafından yeniden yazılmış abstract kullanma. Kaynak türünü `abstract_source` alanına yaz.
JSON dışında yazma.
ŞEMA:
{{
 "additional_relevant_user_documents":[{{"number":"","title":"","reason":""}}],
 "additional_query":"10+ CN... or IN...",
 "d1":{{"number":"","alternate_number":"","title":"","date":"","source":"system/user","source_url":"","summary":"","abstract_en":"","abstract_source":"user-file/official/espacenet/google-patents","figure_reference":"","figure_image_url":""}},
 "d2":null,
 "d1_changed":false,
 "d1_change_note":"D1 değişmedi: ... / D1 değişti: eski ...; yeni ...",
 "d2_changed":false,
 "d2_change_note":"D2 değişmedi: ... / D2 değişti: eski ...; yeni ...",
 "novelty_result":"sağlanır/sağlanmaz",
 "inventive_step_result":"sağlanır/sağlanmaz",
 "technical_opinion":"Bence buluş basamağı ... . Kısa teknik gerekçe...",
 "novelty_reasoning":[""],
 "inventive_step_reasoning":[""],
 "feature_list":[""],
 "comparison_rows_d1":[{{"feature":"","status_evidence":"+ Özet; İstem ...; Şekil ..."}}],
 "comparison_rows_d2":[{{"feature":"","status_evidence":"- ..."}}],
 "helper_documents":[{{"number":"","title":"","source_url":"","role":""}}],
 "warnings":[""]
}}
ARAŞTIRMA KONUSU:\n{bbf_text}\n
İLK TOP10 (AYNEN KORU):\n{json.dumps(top10, ensure_ascii=False, indent=2)}\n
KULLANICI DOKÜMANLARI:\n{user_docs_text}"""

def validate_research_selection(selection: dict[str, Any]) -> None:
    d1_rows = selection.get("comparison_rows_d1") or []
    d2_rows = selection.get("comparison_rows_d2") or []
    for label, rows in (("D1", d1_rows), ("D2", d2_rows)):
        for row in rows:
            evidence = str(row.get("status_evidence") or row.get("status") or "").strip()
            if not (evidence.startswith("+") or evidence.startswith("-")):
                raise ValueError(f"{label} karşılaştırma hücresi + veya - ile başlamalı ve dayanak içermelidir: {evidence!r}")
            if evidence.startswith("+") and len(evidence) < 5:
                raise ValueError(f"{label} karşılaştırma hücresinde '+' yanında dokümandaki somut yer/dayanak da belirtilmelidir.")
            row["status_evidence"] = evidence
    if selection.get("d2"):
        d1_features = [re.sub(r"\s+", " ", str(row.get("feature", ""))).strip() for row in d1_rows]
        d2_features = [re.sub(r"\s+", " ", str(row.get("feature", ""))).strip() for row in d2_rows]
        if d1_features != d2_features:
            raise ValueError("D1 ve D2 karşılaştırma tablolarındaki teknik özellik listeleri birebir aynı değildir.")
    for label, key in (("D1", "d1"), ("D2", "d2")):
        block = selection.get(key)
        if not block:
            continue
        abstract_en = str(block.get("abstract_en") or "").strip()
        if not abstract_en:
            raise ValueError(f"{label} için doğrulanmış özgün İngilizce Abstract bulunamadı.")
        if _contains_turkish_specific_chars(abstract_en):
            raise ValueError(f"{label} abstract_en alanı özgün İngilizce Abstract olmalıdır; Türkçe/çeviri metin kabul edilmez.")


def report_drafting_prompt(
    bbf_text: str,
    top10: dict[str, Any],
    selection: dict[str, Any],
    reference: str,
    cutoff_date: str,
    decision_mode: str,
) -> str:
    return f"""{ARASTIRMA_RULES}
Aşağıdaki verilere göre bağlayıcı `On_Arastirma_Raporu_181612_template.docx` içeriğine tam uyumlu Ön Araştırma Raporu metni oluştur.
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
 "scope":"Global (İlan edilmiş olan patent başvuruları)",
 "keywords":["English keyword or phrase"],
 "ipc_cpc":[{{"code":"","description":"ENGLISH IPC/CPC DESCRIPTION"}}],
 "evaluation_intro":"",
 "documents":[{{
   "label":"D1","number":"","alternate_number":"","title":"","date":"","source_url":"","figure_reference":"","figure_image_url":"",
   "description":["2-3 cümle"],"abstract":"ORIGINAL ENGLISH ABSTRACT VERBATIM","figure_caption":"D1- Şekil",
   "comparison_rows":[{{"feature":"","status_evidence":"+ Özet; İstem ...; Şekil ..."}}],
   "novelty_assessment":["5-10 satırlık değerlendirme"]
 }}],
 "inventive_step_paragraphs":[""],
 "conclusion_paragraphs":[""],
 "warnings":["uyarı paragrafı 1","uyarı paragrafı 2","uyarı paragrafı 3","uyarı paragrafı 4"],
 "attachments":["Benzer Dokümanlar","Ön İnceleme Raporu","Makine Tercümeleri"]
}}
ÖZEL:
- selection.d2 null ise yalnızca D1 bölümü oluştur.
- D1 ve D2 tablolarındaki feature alanları birebir aynı ve aynı sırada olmalıdır.
- comparison_rows sağ hücresi `status_evidence` alanıdır; + veya - ile başlar ve dokümandaki yeri açıkça yazar. Çıplak + / - kullanma.
- `documents` alanındaki D1/D2, NİHAİ SEÇİM içindeki D1/D2 ile aynı dokümanlar olmalı; yayın numarası, alternatif numara, başlık, tarih, source_url ve figure_image_url bilgilerini aynen taşı.
- `keywords` alanı yalnız İngilizce teknik anahtar kelimelerden oluşmalı; Türkçe kelime/ifade kullanma. Şablon 5x2 olduğundan en fazla 10 adet ver.
- `ipc_cpc[].description` yalnız İngilizce resmi sınıflandırma açıklaması olmalı; Türkçe açıklama yazma. Kodların kendisini değiştirme.
- `scope` alanını tam olarak `Global (İlan edilmiş olan patent başvuruları)` yaz; kesim tarihini bu sabit hücreye ekleme.
- D1/D2 `abstract` alanını ASLA yeniden yazma, Türkçeye çevirme veya özetleme. NİHAİ SEÇİM içindeki ilgili `abstract_en` değerini doğrudan kullan. `abstract_en` boşsa özgün İngilizce abstract olmadan raporu tamamlanmış gibi gösterme.
- `warnings` alanı şablondaki uyarı hücresinin dört ayrı paragraf yapısını koruyacak şekilde tam 4 paragraf olarak ver; tek paragrafta satır sonlarıyla birleştirme.
- Rapor metninde BBF/buluş bildirim formu ifadesi kullanma; `araştırma konusu` de.
- `→`, `=>` veya `özellik + özellik + özellik` gibi sembolik/yapay zekâ görünümlü anlatım kullanma.
- Yardımcı dokümanları yeni bir D3 başlığı açmadan yalnız buluş basamağı değerlendirmesinin doğal paragraf akışında kullan.
- Şablonda olmayan bölüm/başlık ekleme.
- Sonuçta yenilik ve buluş basamağı sonucunu açıkça yaz.
- Metni ikinci kez kontrol edip düzelt.
ARAŞTIRMA KONUSU:\n{bbf_text}\n
TOP10:\n{json.dumps(top10, ensure_ascii=False, indent=2)}\n
NİHAİ SEÇİM:\n{json.dumps(selection, ensure_ascii=False, indent=2)}"""


def research_update_analysis_prompt(first_text: str, revised_text: str, prior_report_text: str) -> str:
    return f"""{ARASTIRMA_GUNCELLEME_RULES}
İlk araştırma konusu, revize araştırma konusu ve ilk Ön Araştırma Raporunu birlikte analiz et.
Bu aşamada web araştırması yapma ve Word raporu yazma. İlk ve revize teknik içerik arasındaki gerçek teknik farkları çıkar; yalnız ifade/başlık değişikliklerini teknik katkı sayma.
İlk rapordaki D1/D2 ile yenilik/buluş basamağı gerekçelerini ayrıca çıkar ve her yeni farkın bu gerekçeler üzerindeki etkisini değerlendir.
JSON dışında yazma.
ŞEMA:
{{
 "first_title":"",
 "revised_title":"",
 "prior_d1":{{"number":"","alternate_number":"","title":"","date":""}},
 "prior_d2":{{"number":"","alternate_number":"","title":"","date":""}},
 "prior_novelty_result":"sağlanır/sağlanmaz/belirsiz",
 "prior_inventive_step_result":"sağlanır/sağlanmaz/belirsiz",
 "differences":[{{
   "old":"",
   "new":"",
   "technical_contribution":"evet/hayır/kısmen",
   "technical_effect":"",
   "effect_against_prior_d1_d2":""
 }}],
 "meaningful_change":true,
 "search_focus":[""],
 "preliminary_opinion":"",
 "reasons":[""]
}}
İLK ARAŞTIRMA KONUSU:\n{first_text}\n
REVİZE ARAŞTIRMA KONUSU:\n{revised_text}\n
İLK ÖN ARAŞTIRMA RAPORU:\n{prior_report_text}"""


def research_update_search_prompt(
    revised_text: str,
    prior_report_text: str,
    analysis: dict[str, Any],
    cutoff_date: str,
) -> str:
    return f"""{ARASTIRMA_GUNCELLEME_RULES}
Revize araştırma konusu için araştırma kesim tarihi {cutoff_date} olacak şekilde global patent araştırması yap.
İlk ön araştırma raporundaki D1/D2'yi başlangıç noktası olarak doğrula; özellikle revizyonda eklenen ayırt edici teknik özelliklere odaklan. Eski sorguyu aynen tekrarlama.
TR, EP, US, CN, KR, JP, GB, DE, WIPO ve ilgili diğer veri tabanlarında araştır. En yakın toplam 10 doğrulanmış patent dokümanını sırala. İlk D1/D2 bu 10 içinde olabilir.
Yeni bulunan ve ilk raporda bulunmayan dokümanları ayrıca işaretle. Nihai D1/D2'yi seç; yeni bir belge daha güçlü ise D1/D2'yi değiştir, değilse yardımcı doküman olarak tut.
D1/D2 karşılaştırma tablosunun sağ hücrelerinde yalnız + / - verme; somut yer/dayanak belirt.
Nihai D1/D2 için `abstract_en` alanına doğrulanmış yayımlanmış özgün İngilizce Abstract metnini doğrudan aktar; çeviri veya model özeti kullanma ve kaynağı `abstract_source` alanında belirt.
Özgün patent şekli için doğrulanabilir doğrudan `figure_image_url` bulunabiliyorsa yaz. Model tarafından şekil üretme.
JSON dışında yazma.
ŞEMA:
{{
 "documents":[{{"rank":1,"publication_number":"","alternate_number":"","title":"","date":"","jurisdiction":"","source_url":"","is_new_vs_prior_report":true,"summary":"","abstract_en":"","matching_revision_features":[""],"missing_revision_features":[""],"relevance_score":0}}],
 "new_documents":[{{"number":"","title":"","date":"","source_url":"","technical_relevance":""}}],
 "d1":{{"number":"","alternate_number":"","title":"","date":"","source_url":"","summary":"","abstract_en":"","abstract_source":"official/espacenet/google-patents","figure_reference":"","figure_image_url":""}},
 "d2":{{"number":"","alternate_number":"","title":"","date":"","source_url":"","summary":"","abstract_en":"","abstract_source":"official/espacenet/google-patents","figure_reference":"","figure_image_url":""}},
 "feature_list":[""],
 "comparison_rows_d1":[{{"feature":"","status_evidence":"+ Özet; İstem ...; Şekil ..."}}],
 "comparison_rows_d2":[{{"feature":"","status_evidence":"- ..."}}],
 "helper_documents":[{{"number":"","title":"","date":"","source_url":"","role":""}}],
 "novelty_result":"sağlanır/sağlanmaz",
 "inventive_step_result":"sağlanır/sağlanmaz",
 "technical_opinion":"Kanaatim: ...",
 "totalpatent_query":"TotalPatent arama sorgusu: ..."
}}
REVİZE ARAŞTIRMA KONUSU:\n{revised_text}\n
İLK ÖN ARAŞTIRMA RAPORU:\n{prior_report_text}\n
FARK ANALİZİ:\n{json.dumps(analysis, ensure_ascii=False, indent=2)}"""


def research_update_report_prompt(
    revised_text: str,
    prior_report_text: str,
    analysis: dict[str, Any],
    research: dict[str, Any],
    reference: str,
    cutoff_date: str,
    decision_mode: str,
) -> str:
    return f"""{ARASTIRMA_GUNCELLEME_RULES}
Aşağıdaki verilerden `On_Arastirma_Raporu_181612_template.docx` biçiminde nihai Ön Araştırma Raporu metni oluştur.
Bu bir güncelleme çalışması olsa da Word raporunda yeni bir format veya `Revizyon farkları` bölümü açma. Tam Tip 3 Ön Araştırma Raporu düzenini kullan.
DP referans numarası: {reference}
Araştırma kesim tarihi: {cutoff_date}
Kullanıcının buluş basamağı sonuç seçimi: {decision_mode}
Araştırmanın teknik yenilik sonucu: {research.get('novelty_result','')}

Kullanıcı 'Buluş basamağı sağlanıyor' seçtiyse inventive step sonucunu `sağlanır`, 'Buluş basamağı sağlanmıyor' seçtiyse `sağlanmaz` yaz. Ancak yenilik sonucu kaynaklarla çelişecek şekilde değiştirilmez.
İlk rapora atıf gerekiyorsa `ilk ön araştırma raporu`, teknik konuya atıf gerekiyorsa `revize araştırma konusu` de. `BBF` ifadesi kullanma.
Yeni bulunan yardımcı dokümanları şablonda olmayan D3/D4 başlığı açmadan buluş basamağı değerlendirmesinin doğal paragrafı içinde açıkla. Daha güçlü yeni doküman D1/D2 seçilmişse normal D1/D2 bölümünde kullan.
`documents` alanındaki D1 ve D2, YENİ ARAŞTIRMA içindeki `d1` ve `d2` ile aynı dokümanlar olmalı; yayın numarası, alternatif numara, başlık, tarih, source_url, figure_reference ve figure_image_url bilgilerini değiştirmeden taşı.
Karşılaştırma hücrelerinde + veya - işaretinin ardından dokümandaki somut yeri yaz.
`keywords` yalnız İngilizce olsun ve en fazla 10 adet ver; `ipc_cpc[].description` yalnız İngilizce resmi sınıflandırma açıklaması olsun.
`scope` tam olarak `Global (İlan edilmiş olan patent başvuruları)` olarak kalsın.
D1/D2 `abstract` alanına YENİ ARAŞTIRMA içindeki ilgili dokümanın `abstract_en` değerini doğrudan aktar; Türkçeye çevirme, özetleme veya yeniden yazma. Özgün İngilizce abstract yoksa raporu tamamlanmış gibi üretme.
`warnings` alanını şablondaki dört ayrı uyarı paragrafını koruyacak şekilde tam 4 paragraf olarak üret.
`→`, `=>`, oklar veya `özellik + özellik` gibi kısa sembolik anlatım kullanma.
JSON dışında yazma.
ŞEMA:
{{
 "reference":"{reference}",
 "title":"",
 "report_date":"{date.today().strftime('%d.%m.%Y')}",
 "purpose":"Belirlenen konuda araştırmanın gerçekleştirilmesi",
 "scope":"Global (İlan edilmiş olan patent başvuruları)",
 "keywords":["English keyword or phrase"],
 "ipc_cpc":[{{"code":"","description":"ENGLISH IPC/CPC DESCRIPTION"}}],
 "evaluation_intro":"",
 "documents":[{{
   "label":"D1","number":"","alternate_number":"","title":"","date":"","source_url":"","figure_reference":"","figure_image_url":"",
   "description":[""],"abstract":"ORIGINAL ENGLISH ABSTRACT VERBATIM","figure_caption":"D1- Şekil",
   "comparison_rows":[{{"feature":"","status_evidence":"+ Özet; İstem ...; Şekil ..."}}],
   "novelty_assessment":[""]
 }}],
 "inventive_step_paragraphs":[""],
 "conclusion_paragraphs":[""],
 "warnings":["uyarı paragrafı 1","uyarı paragrafı 2","uyarı paragrafı 3","uyarı paragrafı 4"],
 "attachments":["Benzer Dokümanlar","Ön İnceleme Raporu","Makine Tercümeleri"]
}}
REVİZE ARAŞTIRMA KONUSU:\n{revised_text}\n
İLK ÖN ARAŞTIRMA RAPORU:\n{prior_report_text}\n
FARK ANALİZİ:\n{json.dumps(analysis, ensure_ascii=False, indent=2)}\n
YENİ ARAŞTIRMA:\n{json.dumps(research, ensure_ascii=False, indent=2)}"""


def _iter_strings(value: Any):
    if isinstance(value, str):
        yield value
    elif isinstance(value, dict):
        for v in value.values():
            yield from _iter_strings(v)
    elif isinstance(value, (list, tuple)):
        for v in value:
            yield from _iter_strings(v)


def _contains_turkish_specific_chars(value: str) -> bool:
    return bool(re.search(r"[çğıöşüÇĞİÖŞÜ]", str(value or "")))


def _norm_ws(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def validate_research_report_language(report: dict[str, Any]) -> None:
    banned = [r"\bBBF\b", r"buluş bildirim formu", r"→", r"=>"]
    # Tablo durum/evidence hücrelerindeki + / - izinlidir; diğer alanlarda ok ve BBF dili yasaktır.
    for text in _iter_strings(report):
        for pattern in banned:
            if re.search(pattern, text, flags=re.IGNORECASE):
                raise ValueError(f"Ön araştırma raporu metninde kullanılmaması gereken ifade bulundu: {pattern}")

    keywords = [str(x).strip() for x in (report.get("keywords") or []) if str(x).strip()]
    if len(keywords) > 10:
        raise ValueError("Tip 3 anahtar kelime alanı şablondaki 5x2 yapı gereği en fazla 10 İngilizce ifade içerebilir.")
    if any(_contains_turkish_specific_chars(x) for x in keywords):
        raise ValueError("Tip 3 raporundaki anahtar kelimelerin tamamı İngilizce olmalıdır.")

    ipc_items = report.get("ipc_cpc") or []
    if len(ipc_items) > 4:
        raise ValueError("IPC/CPC alanı bağlayıcı şablonun dört satırlık yapısını aşmamalıdır.")
    for item in ipc_items:
        code = str(item.get("code") or "").strip()
        desc = str(item.get("description") or "").strip()
        if not code or not desc:
            raise ValueError("Her IPC/CPC satırında hem kod hem İngilizce açıklama bulunmalıdır.")
        if _contains_turkish_specific_chars(desc):
            raise ValueError("IPC/CPC açıklamalarının tamamı İngilizce olmalıdır.")

    scope = str(report.get("scope") or "").strip()
    if scope and scope != "Global (İlan edilmiş olan patent başvuruları)":
        raise ValueError("Tip 3 Kapsam hücresi bağlayıcı şablondaki sabit metinle aynı olmalıdır.")

    for block in report.get("documents") or []:
        abstract = str(block.get("abstract") or "").strip()
        if not abstract:
            raise ValueError(f"{block.get('label','D1/D2')} için özgün İngilizce Abstract bulunmadan rapor oluşturulamaz.")
        if _contains_turkish_specific_chars(abstract):
            raise ValueError(f"{block.get('label','D1/D2')} Abstract alanı özgün İngilizce metin olmalıdır; Türkçe/çeviri metin kabul edilmez.")


def validate_report_against_selection(report: dict[str, Any], selection: dict[str, Any]) -> None:
    """Nihai D1/D2 kimliği ve özgün İngilizce abstract metni rapora model tarafından değiştirilmeden taşınır."""
    selected_docs = [selection.get("d1"), selection.get("d2")]
    selected_docs = [d for d in selected_docs if d]
    report_docs = report.get("documents") or []
    if len(report_docs) != len(selected_docs):
        raise ValueError("Rapor D1/D2 sayısı nihai seçimle aynı değildir.")
    for idx, (src, dst) in enumerate(zip(selected_docs, report_docs), 1):
        src_num = str(src.get("number") or "").strip()
        dst_num = str(dst.get("number") or "").strip()
        if _norm_ws(src_num).lower() != _norm_ws(dst_num).lower():
            raise ValueError(f"Rapor D{idx} dokümanı nihai seçimdeki D{idx} ile aynı değildir.")
        abstract_en = str(src.get("abstract_en") or "").strip()
        if not abstract_en:
            raise ValueError(f"Nihai D{idx} için doğrulanmış özgün İngilizce Abstract bulunamadı.")
        # Modelin yeniden yazmasını engelle: rapor alanını doğrudan kaynaktaki özgün abstract ile üzerine yaz.
        dst["abstract"] = abstract_en


def _replace_paragraph_text_preserve_format(paragraph, text: str) -> None:
    proto = next((r for r in paragraph.runs if r.text or r._r.rPr is not None), None)
    proto_rpr = deepcopy(proto._r.rPr) if proto is not None and proto._r.rPr is not None else None
    paragraph.clear()
    run = paragraph.add_run(str(text or ""))
    if proto_rpr is not None:
        current = run._r.rPr
        if current is not None:
            run._r.remove(current)
        run._r.insert(0, proto_rpr)


def _replace_cell_text_preserve_format(cell, text: str) -> None:
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    _replace_paragraph_text_preserve_format(p, text)
    # Hücrede önceki metinden kalan fazladan paragrafları temizle; hücre ölçüsü/kenarlıkları korunur.
    for extra in list(cell.paragraphs[1:]):
        extra._element.getparent().remove(extra._element)


def _replace_cell_lines_preserve_format(cell, lines: list[str]) -> None:
    vals = [str(x).strip() for x in lines if str(x).strip()]
    protos = list(cell.paragraphs)
    # keep at least one paragraph
    while len(cell.paragraphs) > 1:
        extra = cell.paragraphs[-1]
        extra._element.getparent().remove(extra._element)
    if not vals:
        _replace_paragraph_text_preserve_format(cell.paragraphs[0], "")
        return
    _replace_paragraph_text_preserve_format(cell.paragraphs[0], vals[0])
    for i, value in enumerate(vals[1:], 1):
        new_p = cell.add_paragraph()
        if i < len(protos):
            src = protos[i]
            # paragraph properties
            if src._p.pPr is not None:
                new_p._p.insert(0, deepcopy(src._p.pPr))
            proto_run = next((r for r in src.runs if r.text or r._r.rPr is not None), None)
            run = new_p.add_run(value)
            if proto_run is not None and proto_run._r.rPr is not None:
                if run._r.rPr is not None:
                    run._r.remove(run._r.rPr)
                run._r.insert(0, deepcopy(proto_run._r.rPr))
        else:
            _replace_paragraph_text_preserve_format(new_p, value)


def _fill_keyword_table(cell, keywords: list[str]) -> None:
    if not cell.tables:
        return
    table = cell.tables[0]
    vals = [str(x).strip() for x in keywords if str(x).strip()][:10]
    vals += [""] * (10 - len(vals))
    k = 0
    for row in table.rows:
        for c in row.cells:
            _replace_cell_text_preserve_format(c, vals[k])
            k += 1


def _fill_ipc_cell(cell, ipc_items: list[dict[str, Any]]) -> None:
    """Şablondaki dört IPC paragrafını korur: kod kalın, İngilizce açıklama normal."""
    items = [x for x in (ipc_items or []) if str(x.get("code") or "").strip() and str(x.get("description") or "").strip()][:4]
    protos = list(cell.paragraphs)
    if not protos:
        protos = [cell.add_paragraph()]
    # Şablondaki dört paragraf geometrisini sabit tut.
    while len(cell.paragraphs) > 4:
        extra = cell.paragraphs[-1]
        extra._element.getparent().remove(extra._element)
    while len(cell.paragraphs) < 4:
        cell.add_paragraph()

    def fill_para(paragraph, item, proto):
        proto_ppr = deepcopy(proto._p.pPr) if proto is not None and proto._p.pPr is not None else None
        proto_run_rprs = [deepcopy(r._r.rPr) if r._r.rPr is not None else None for r in (list(proto.runs) if proto is not None else [])]
        if proto_ppr is not None:
            if paragraph._p.pPr is not None:
                paragraph._p.remove(paragraph._p.pPr)
            paragraph._p.insert(0, proto_ppr)
        paragraph.clear()
        if item is None:
            return
        code_run = paragraph.add_run(f"{str(item.get('code') or '').strip()}: ")
        desc_run = paragraph.add_run(str(item.get('description') or '').strip())
        if proto_run_rprs and proto_run_rprs[0] is not None:
            if code_run._r.rPr is not None:
                code_run._r.remove(code_run._r.rPr)
            code_run._r.insert(0, proto_run_rprs[0])
        code_run.bold = True
        if len(proto_run_rprs) > 1 and proto_run_rprs[1] is not None:
            if desc_run._r.rPr is not None:
                desc_run._r.remove(desc_run._r.rPr)
            desc_run._r.insert(0, proto_run_rprs[1])
        desc_run.bold = False

    for i in range(4):
        proto = protos[i] if i < len(protos) else protos[-1]
        fill_para(cell.paragraphs[i], items[i] if i < len(items) else None, proto)


def _fill_warning_cell(cell, warnings: list[str]) -> None:
    """Şablondaki dört ayrı uyarı paragrafını korur; yeni bölüm/paragraf geometrisi oluşturmaz."""
    vals = [str(x).strip() for x in (warnings or []) if str(x).strip()][:4]
    vals += [""] * (4 - len(vals))
    protos = list(cell.paragraphs)
    if not protos:
        protos = [cell.add_paragraph()]
    while len(cell.paragraphs) > 4:
        extra = cell.paragraphs[-1]
        extra._element.getparent().remove(extra._element)
    while len(cell.paragraphs) < 4:
        cell.add_paragraph()
    for i in range(4):
        p = cell.paragraphs[i]
        proto = protos[i] if i < len(protos) else protos[-1]
        proto_ppr = deepcopy(proto._p.pPr) if proto._p.pPr is not None else None
        if proto_ppr is not None:
            if p._p.pPr is not None:
                p._p.remove(p._p.pPr)
            p._p.insert(0, proto_ppr)
        _replace_paragraph_text_preserve_format(p, vals[i])

def _replace_comparison_table(table, rows: list[dict[str, Any]], label: str) -> None:
    # Şablon başlığını, satır yüksekliğini, hücre genişliklerini ve yazı biçimini koru.
    header = table.rows[0]
    _replace_cell_text_preserve_format(header.cells[0], "Araştırma konusu")
    _replace_cell_text_preserve_format(header.cells[1], f"{label} Dokümanı")
    prototype_tr = deepcopy(table.rows[1]._tr) if len(table.rows) > 1 else None
    for tr in list(table._tbl.tr_lst)[1:]:
        table._tbl.remove(tr)
    for item in rows:
        if prototype_tr is not None:
            table._tbl.append(deepcopy(prototype_tr))
            cells = table.rows[-1].cells
        else:
            cells = table.add_row().cells
        _replace_cell_text_preserve_format(cells[0], str(item.get("feature", "")))
        evidence = str(item.get("status_evidence") or item.get("status") or "").strip()
        _replace_cell_text_preserve_format(cells[1], evidence)


def _safe_remote_url(url: str) -> bool:
    url = str(url or "").strip().lower()
    return url.startswith("https://patents.google.com/") or url.startswith("https://patentimages.storage.googleapis.com/") or url.startswith("https://patentscope.wipo.int/") or url.startswith("https://worldwide.espacenet.com/")


def _fetch_remote_bytes(url: str, timeout: int = 30) -> bytes:
    if not _safe_remote_url(url):
        raise ValueError("Patent şekli için yalnız doğrulanabilir patent kaynağı URL'si kullanılabilir.")
    req = Request(url, headers={"User-Agent": "Mozilla/5.0 PatentAtolyesi/5.4"})
    with urlopen(req, timeout=timeout) as resp:
        data = resp.read(25 * 1024 * 1024 + 1)
    if len(data) > 25 * 1024 * 1024:
        raise ValueError("Patent şekli dosyası beklenenden büyük.")
    return data


def _google_patent_image_urls(source_url: str) -> list[str]:
    if not str(source_url).startswith("https://patents.google.com/"):
        return []
    try:
        html = _fetch_remote_bytes(source_url).decode("utf-8", errors="ignore")
    except Exception:
        return []
    urls = re.findall(r"https://patentimages\.storage\.googleapis\.com/[^\"'<>\s]+?\.(?:png|jpg|jpeg|webp)", html, flags=re.I)
    out: list[str] = []
    seen: set[str] = set()
    for url in urls:
        url = html_lib.unescape(url)
        if url in seen:
            continue
        seen.add(url)
        out.append(url)
    return out


def _normalize_image_bytes(data: bytes) -> bytes:
    with Image.open(io.BytesIO(data)) as im:
        im.load()
        if im.width < 300 or im.height < 180:
            raise ValueError("Patent şekli çözünürlüğü yetersiz.")
        if im.format == "PNG":
            return data
        out = io.BytesIO()
        if im.mode not in {"RGB", "RGBA", "L"}:
            im = im.convert("RGB")
        im.save(out, format="PNG")
        return out.getvalue()


def resolve_original_patent_figure(document_info: dict[str, Any]) -> bytes | None:
    """Yalnız özgün patent kaynağından şekil getirir; hiçbir zaman yapay şekil üretmez."""
    candidates: list[str] = []
    direct = str(document_info.get("figure_image_url") or "").strip()
    if direct:
        candidates.append(direct)
    source = str(document_info.get("source_url") or "").strip()
    if source:
        urls = _google_patent_image_urls(source)
        # D00000 çoğunlukla kapak/ilk çizim sayfasıdır; mümkünse ilk gerçek çizim sayfasını seç.
        preferred = [u for u in urls if "D00000" not in u.upper()]
        candidates.extend(preferred or urls)
    seen: set[str] = set()
    for url in candidates:
        if url in seen:
            continue
        seen.add(url)
        try:
            return _normalize_image_bytes(_fetch_remote_bytes(url))
        except Exception:
            continue
    return None


def _report_pdf_fallback_figures(asset: UploadedAsset | None) -> list[bytes]:
    """İlk raporda zaten bulunan özgün patent şekillerini son çare olarak yeniden kullanır."""
    if asset is None:
        return []
    imgs = extract_embedded_images(asset)
    unique: list[tuple[int, bytes]] = []
    seen: set[str] = set()
    import hashlib
    for img in imgs:
        try:
            with Image.open(io.BytesIO(img.data)) as im:
                area = im.width * im.height
                if im.width < 420 or im.height < 230:
                    continue
        except Exception:
            continue
        h = hashlib.sha1(img.data).hexdigest()
        if h in seen:
            continue
        seen.add(h)
        unique.append((area, img.data))
    # Patent şekilleri rapor logolarından belirgin biçimde daha büyük olur. Belge sırasını mümkün olduğunca korumak için
    # ilk dört büyük adayı alıp kaynak çıkarma sırasına göre kullanıyoruz.
    large_hashes = {hashlib.sha1(data).hexdigest() for _, data in sorted(unique, key=lambda x: x[0], reverse=True)[:4]}
    return [data for _, data in unique if hashlib.sha1(data).hexdigest() in large_hashes]


def _replace_figure_paragraph(paragraph, image_data: bytes | None) -> None:
    paragraph.clear()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if not image_data:
        return
    try:
        with Image.open(io.BytesIO(image_data)) as im:
            w, h = im.size
        max_w_cm, max_h_cm = 15.8, 11.5
        ratio = min(max_w_cm / max(w, 1), max_h_cm / max(h, 1))
        width_cm = max(5.0, min(max_w_cm, w * ratio))
        paragraph.add_run().add_picture(io.BytesIO(image_data), width=Cm(width_cm))
    except Exception as exc:
        raise ValueError("Özgün patent şekli Word raporuna eklenemedi.") from exc


def _validate_research_template_fidelity(doc: Document) -> None:
    """Dinamik içerik dışında bağlayıcı Tip 3 şablon geometrisinin bozulmadığını doğrular."""
    template = Document(str(ARASTIRMA_TEMPLATE))
    if len(doc.sections) != len(template.sections):
        raise ValueError("Tip 3 çıktısında section sayısı bağlayıcı şablondan farklıdır.")
    for out_sec, tpl_sec in zip(doc.sections, template.sections):
        for attr in ("top_margin", "bottom_margin", "left_margin", "right_margin", "page_width", "page_height"):
            if getattr(out_sec, attr) != getattr(tpl_sec, attr):
                raise ValueError(f"Tip 3 çıktısında {attr} bağlayıcı şablondan farklıdır.")
    if len(doc.paragraphs) != len(template.paragraphs):
        raise ValueError("Tip 3 çıktısında şablonda olmayan ana paragraf/başlık eklenmiş veya silinmiştir.")
    if len(doc.tables) != len(template.tables):
        raise ValueError("Tip 3 çıktısındaki ana tablo sayısı bağlayıcı şablondan farklıdır.")

    criteria = doc.tables[0]
    tpl_criteria = template.tables[0]
    if len(criteria.rows) != 5:
        raise ValueError("Tip 3 kriter tablosunun satır yapısı bozulmuştur.")
    for r in range(5):
        # Dinamik içerik hücreleri dışındaki sabit etiketleri koru.
        if _norm_ws(criteria.rows[r].cells[0].text) != _norm_ws(tpl_criteria.rows[r].cells[0].text):
            raise ValueError("Tip 3 kriter tablosundaki sabit alan etiketi şablondan farklıdır.")
        if _norm_ws(criteria.rows[r].cells[1].text) != _norm_ws(tpl_criteria.rows[r].cells[1].text):
            raise ValueError("Tip 3 kriter tablosundaki iki nokta/ayraç hücresi şablondan farklıdır.")

    if _norm_ws(criteria.rows[2].cells[2].text) != _norm_ws("Global (İlan edilmiş olan patent başvuruları)"):
        raise ValueError("Tip 3 Kapsam hücresi bağlayıcı şablondaki sabit metni korumuyor.")

    keyword_cell = criteria.rows[3].cells[2]
    if len(keyword_cell.tables) != 1 or len(keyword_cell.tables[0].rows) != 5 or len(keyword_cell.tables[0].columns) != 2:
        raise ValueError("Tip 3 anahtar kelime tablosunun 5x2 şablon geometrisi bozulmuştur.")

    ipc_cell = criteria.rows[4].cells[2]
    if len(ipc_cell.paragraphs) != 4:
        raise ValueError("Tip 3 IPC alanı şablondaki dört paragraf yapısını korumuyor.")
    for para in ipc_cell.paragraphs:
        if not para.text.strip():
            continue
        if len(para.runs) < 2 or para.runs[0].bold is not True or para.runs[1].bold is True:
            raise ValueError("Tip 3 IPC alanında kod kalın, İngilizce açıklama normal yazı biçimi korunmalıdır.")

    warning_cell = doc.tables[4].rows[0].cells[2]
    if len(warning_cell.paragraphs) != 4:
        raise ValueError("Tip 3 uyarı alanı şablondaki dört ayrı paragraf yapısını korumuyor.")


def build_research_docx(report: dict[str, Any], figure_fallbacks: list[bytes] | None = None) -> bytes:
    """Bağlayıcı Tip 3 şablonunu yerinde doldurur; gövdeyi yeniden kurmaz."""
    validate_research_report_language(report)
    docs = report.get("documents") or []
    if not docs:
        raise ValueError("Rapor için D1 dokümanı bulunamadı.")
    if len(docs) > 2:
        docs = docs[:2]
    d1 = docs[0]
    d2 = docs[1] if len(docs) > 1 else None

    doc = Document(str(ARASTIRMA_TEMPLATE))
    ps = doc.paragraphs
    if len(ps) < 107 or len(doc.tables) < 5:
        raise ValueError("Bağlayıcı Ön Araştırma Raporu şablonunun yapısı beklenen formatta değil.")

    # Kapak ve kriterler: paragraf/boşluk/section yapısı aynen korunur.
    _replace_paragraph_text_preserve_format(ps[1], report.get("reference", ""))
    _replace_paragraph_text_preserve_format(ps[6], report.get("title", ""))
    _replace_paragraph_text_preserve_format(ps[29], report.get("report_date", date.today().strftime('%d.%m.%Y')))

    criteria = doc.tables[0]
    _replace_cell_text_preserve_format(criteria.rows[0].cells[2], report.get("purpose", "Belirlenen konuda araştırmanın gerçekleştirilmesi"))
    _replace_cell_text_preserve_format(criteria.rows[1].cells[2], report.get("title", ""))
    _replace_cell_text_preserve_format(criteria.rows[2].cells[2], "Global (İlan edilmiş olan patent başvuruları)")
    _fill_keyword_table(criteria.rows[3].cells[2], report.get("keywords") or [])
    _fill_ipc_cell(criteria.rows[4].cells[2], report.get("ipc_cpc") or [])

    _replace_paragraph_text_preserve_format(ps[36], report.get("evaluation_intro", ""))

    def apply_doc(block: dict[str, Any], label: str, indices: dict[str, int], table_idx: int, fallback_idx: int):
        header = f"{label}- {block.get('number','')}"
        if block.get("alternate_number"):
            header += f" ({block.get('alternate_number')})"
        header += f"- {block.get('title','')}- {block.get('date','')}"
        _replace_paragraph_text_preserve_format(ps[indices['header']], header)
        _replace_paragraph_text_preserve_format(ps[indices['description']], " ".join(block.get("description") or []))
        _replace_paragraph_text_preserve_format(ps[indices['abstract']], block.get("abstract", ""))
        _replace_paragraph_text_preserve_format(ps[indices['figure_heading']], block.get("figure_caption", f"{label}- Şekil"))
        _replace_paragraph_text_preserve_format(ps[indices['comparison_heading']], f"Araştırma konusu ile {label} dokümanı arasında benzerlik değerlendirmesi:")
        _replace_comparison_table(doc.tables[table_idx], block.get("comparison_rows") or [], label)
        _replace_paragraph_text_preserve_format(ps[indices['assessment']], " ".join(block.get("novelty_assessment") or []))
        fig = resolve_original_patent_figure(block)
        if fig is None and figure_fallbacks and fallback_idx < len(figure_fallbacks):
            fig = figure_fallbacks[fallback_idx]
        if fig is None:
            raise ValueError(f"{label} için özgün patent şekli otomatik temin edilemedi. Yapay şekil oluşturulmadı; ilgili patentin orijinal PDF/şekil kaynağı gereklidir.")
        _replace_figure_paragraph(ps[indices['figure']], fig)

    apply_doc(d1, "D1", {"header":40,"description":42,"abstract":46,"figure_heading":47,"figure":48,"comparison_heading":49,"assessment":55}, 2, 0)

    if d2:
        apply_doc(d2, "D2", {"header":57,"description":59,"abstract":62,"figure_heading":64,"figure":65,"comparison_heading":66,"assessment":71}, 3, 1)
    else:
        # Şablonda D2 için ayrılmış alanı görünmez kıl; diğer format öğelerini bozma.
        for i in [57,59,61,62,64,65,66,68,69,71]:
            _replace_paragraph_text_preserve_format(ps[i], "")
        for tr in list(doc.tables[3]._tbl.tr_lst)[1:]:
            doc.tables[3]._tbl.remove(tr)
        _replace_cell_text_preserve_format(doc.tables[3].rows[0].cells[0], "")
        _replace_cell_text_preserve_format(doc.tables[3].rows[0].cells[1], "")

    inv = report.get("inventive_step_paragraphs") or []
    slots = [75, 77, 79, 81]
    for i, pidx in enumerate(slots):
        _replace_paragraph_text_preserve_format(ps[pidx], inv[i] if i < len(inv) else "")

    conclusion = report.get("conclusion_paragraphs") or []
    _replace_paragraph_text_preserve_format(ps[85], " ".join(conclusion))

    # Uyarılar tablosu şablonun kendi yerinde kalır; ayrı yeni bölüm açılmaz.
    warnings = [str(x).strip() for x in (report.get("warnings") or []) if str(x).strip()]
    _fill_warning_cell(doc.tables[4].rows[0].cells[2], warnings)

    _validate_research_template_fidelity(doc)
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
      <p>Tarifname, görüş, Tip 3 ön araştırma ve araştırma güncelleme çalışmalarını tek arayüzden oluşturun.</p>
      <div class="version">Kural sürümü: {RULESET_VERSION}</div>
    </div>
    """,
    unsafe_allow_html=True,
)


if not os.getenv("OPENAI_API_KEY", "").strip():
    st.warning("OPENAI_API_KEY henüz tanımlı değil. Arayüzü inceleyebilirsiniz; üretim düğmeleri API anahtarı olmadan çalışmaz.")

work_type = st.radio(
    "İş türü",
    ["Tarifname oluşturma", "Görüş hazırlama", "Tip 3 - Ön araştırma raporu", "Araştırma güncelleme - Tip 3"],
    horizontal=True,
)

# TARİFNAME
if work_type == "Tarifname oluşturma":
    st.subheader("Tarifname oluşturma")
    with st.form("tarifname_form"):
        c1, c2 = st.columns(2)
        with c1:
            bbf = st.file_uploader("BBF dosyası", type=["docx", "doc", "pdf", "txt"], key="tar_bbf")
            reference = st.text_input("DP referans numarası", value="")
            output_name = st.text_input("Çıktı dosyasının adı", value="Tarifname_XXXXXX.docx")
        with c2:
            language_choice = st.selectbox("Tarifname dili", ["Türkçe", "İngilizce"], index=0)
            claim_choice = st.selectbox(
                "İstem yapısı",
                ["BBF'ye göre otomatik belirle", "Yalnızca sistem", "Yalnızca yöntem", "Sistem ve yöntem"],
            )
            st.caption("Mevcut bir tarifnameyi değiştirme işlemi bu ekranda yapılmaz; tarifname düzenleme ayrı bir iş akışı olarak ele alınacaktır.")

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

                technical_assets = assets_from_uploads(extra_technical_files)
                technical_text, technical_images = combine_asset_text("EK TEKNİK BELGE", technical_assets)
                example_assets = assets_from_uploads(example_files)
                example_text, _ = combine_asset_text("ÖRNEK TARİFNAME - YALNIZCA KURGU", example_assets)

                embedded_images = extract_embedded_images(bbf_asset)
                provided_figure_assets: list[UploadedAsset] = []
                if separate_figures:
                    for uploaded in figure_files or []:
                        fig_asset = UploadedAsset(uploaded.name, uploaded.getvalue(), uploaded.type)
                        provided_figure_assets.extend(extract_embedded_images(fig_asset))

                # Nihai şekil dosyaları ayrıca yüklenmişse model bunları da görerek referans senkronizasyonunu denetler.
                model_images = [*provided_figure_assets, *technical_images, *embedded_images][:12]

                progress.progress(15, text="BBF içeriği, referans tablosu, şekiller ve istem çekirdeği çıkarılıyor...")
                extracted = ask_json(
                    tarifname_extraction_prompt(source, technical_text, example_text, language_choice),
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
                            tarifname_literature_prompt(extracted, int(lit_count), jurisdiction, language_choice),
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
                        technical_text,
                        example_text,
                        language_choice,
                    ),
                    images=model_images,
                )

                progress.progress(73, text="BBF ile tamlık ve istem tutarlılığı ikinci kez kontrol ediliyor...")
                draft = ask_json(
                    tarifname_quality_prompt(
                        source,
                        technical_text,
                        extracted,
                        draft,
                        mode,
                        lit_docs,
                        language_choice,
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

                draft = apply_tarifname_house_style(draft, mode, lit_docs, language_choice)
                warnings = validate_tarifname_draft(draft, mode, lit_docs, language_choice)
                for warning in warnings:
                    st.warning(warning)

                progress.progress(88, text="Word dosyası hazırlanıyor...")
                data = build_tarifname_docx(draft, language_choice)

                figure_data = None
                if separate_figures:
                    # Ayrıca yüklenen müşteri şekilleri birincil kaynaktır; yoksa BBF içindeki özgün görseller kullanılır.
                    all_figure_assets: list[UploadedAsset] = list(provided_figure_assets)
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
                    figure_data = build_figures_docx(deduplicated, language_choice)

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
elif work_type == "Tip 3 - Ön araştırma raporu":
    st.subheader("Tip 3 - Ön araştırma raporu")
    c1, c2 = st.columns(2)
    with c1:
        bbf = st.file_uploader("BBF dosyası", type=["docx", "doc", "pdf", "txt"], key="res_bbf")
        reference = st.text_input("DP referans numarası", value="")
    with c2:
        output_name = st.text_input("Çıktı dosyasının adı", value="Ön Araştırma Raporu_XXXXXX.docx")
        cutoff = st.date_input("Araştırma kesim tarihi", value=date.today())

    for key, default in {
        "top10_result": None,
        "research_bbf_text": None,
        "research_cutoff": None,
        "research_selection": None,
        "research_user_images": [],
    }.items():
        if key not in st.session_state:
            st.session_state[key] = default

    if st.button("1. Global araştırmayı yap ve en benzer 10 dokümanı bul", type="primary", use_container_width=True):
        if bbf is None:
            st.error("BBF yükleyin.")
        elif not reference.strip():
            st.error("DP referans numarasını girin.")
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
                st.session_state.research_selection = None
                st.session_state.research_user_images = []
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
        rows = [{
            "Sıra": d.get("rank"), "Yayın no": d.get("publication_number"), "Başlık": d.get("title"),
            "Tarih": d.get("date"), "Yakınlık": d.get("relevance_score"),
            "Yeniliği bozar mı?": "Evet" if d.get("novelty_destroying") else "Hayır",
        } for d in docs]
        st.dataframe(rows, use_container_width=True, hide_index=True)
        st.caption(f"Önerilen D1: {st.session_state.top10_result.get('proposed_d1','')} | Önerilen D2: {st.session_state.top10_result.get('proposed_d2','') or '-'}")

        own_docs = st.radio("Sizin araştırdığınız benzer dokümanlar var mı?", ["Hayır", "Evet"], horizontal=True)
        user_files = []
        if own_docs == "Evet":
            user_files = st.file_uploader("Sizin bulduğunuz benzer dokümanlar", type=["pdf", "zip", "docx", "doc", "txt", "png", "jpg", "jpeg", "webp"], accept_multiple_files=True, key="res_user_docs")

        if st.button("2. Kullanıcı dokümanlarını incele ve nihai D1/D2'yi belirle", type="primary", use_container_width=True):
            if own_docs == "Evet" and not user_files:
                st.error("Benzer dokümanları yükleyin.")
            else:
                try:
                    progress = st.progress(0, text="Kullanıcı dokümanları inceleniyor...")
                    user_assets = assets_from_uploads(user_files)
                    user_text, user_images = combine_asset_text("KULLANICI BENZER DOKÜMANI", user_assets)
                    progress.progress(45, text="10+ dokümanları ve nihai D1/D2 belirleniyor...")
                    selection = ask_json(final_selection_prompt(
                        st.session_state.research_bbf_text,
                        st.session_state.top10_result,
                        user_text,
                    ), images=user_images, web_search=True)
                    validate_research_selection(selection)
                    st.session_state.research_selection = selection
                    st.session_state.research_user_images = user_images
                    progress.progress(100, text="Teknik değerlendirme tamamlandı")
                except Exception as exc:
                    st.exception(exc)

        selection = st.session_state.research_selection
        if selection:
            additional_query = str(selection.get("additional_query") or "").strip()
            if additional_query:
                st.code(additional_query, language=None)
            elif own_docs == "Evet":
                st.caption("10+ olarak eklenecek yeterince ilgili kullanıcı dokümanı belirlenmedi.")

            d1_note = selection.get("d1_change_note") or (
                "D1 değişti." if selection.get("d1_changed") else "D1 değişmedi."
            )
            d2_note = selection.get("d2_change_note") or (
                "D2 değişti." if selection.get("d2_changed") else "D2 değişmedi."
            )
            st.write(f"**{d1_note}**")
            st.write(f"**{d2_note}**")
            st.success(f"Nihai D1: {selection.get('d1',{}).get('number','')} | Nihai D2: {(selection.get('d2') or {}).get('number','-')}")
            st.info(selection.get("technical_opinion") or f"Bence buluş basamağı {'var' if selection.get('inventive_step_result') == 'sağlanır' else 'yok'}.")

            decision_mode = st.selectbox(
                "Rapor sonucunu nasıl oluşturayım?",
                ["Otomatik belirle", "Buluş basamağı var", "Buluş basamağı yok"],
                help="Bu seçim ancak 10+, nihai D1/D2 ve sistem kanaati gösterildikten sonra yapılır.",
            )

            if st.button("3. Ön Araştırma Raporunu oluştur", type="primary", use_container_width=True):
                try:
                    final_selection = deepcopy(selection)
                    if decision_mode == "Buluş basamağı var":
                        final_selection["inventive_step_result"] = "sağlanır"
                    elif decision_mode == "Buluş basamağı yok":
                        final_selection["inventive_step_result"] = "sağlanmaz"
                    progress = st.progress(0, text="Yenilik ve buluş basamağı raporu hazırlanıyor...")
                    report = ask_json(report_drafting_prompt(
                        st.session_state.research_bbf_text,
                        st.session_state.top10_result,
                        final_selection,
                        reference,
                        st.session_state.research_cutoff or cutoff.strftime("%d.%m.%Y"),
                        decision_mode,
                    ))
                    validate_report_against_selection(report, final_selection)
                    validate_research_report_language(report)
                    progress.progress(75, text="Bağlayıcı Word şablonu dolduruluyor...")
                    data = build_research_docx(report)
                    progress.progress(100, text="Hazır")
                    st.info(f"Yenilik: {final_selection.get('novelty_result','')} | Buluş basamağı: {final_selection.get('inventive_step_result','')}")
                    effective_output_name = output_name.replace("XXXXXX", reference.strip()) if reference.strip() else output_name
                    st.download_button("Word raporunu indir", data=data, file_name=safe_output_name(effective_output_name, "Ön Araştırma Raporu.docx"), mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")
                except Exception as exc:
                    st.exception(exc)

# ARAŞTIRMA GÜNCELLEME
else:
    st.subheader("Araştırma güncelleme - Tip 3")
    st.caption("İlk araştırma konusu ile revize araştırma konusu karşılaştırılır; ilk rapordaki D1/D2 dikkate alınarak yeni araştırma yapılır ve nihai rapor standart Tip 3 Ön Araştırma Raporu formatında oluşturulur.")

    c1, c2 = st.columns(2)
    with c1:
        first_bbf = st.file_uploader("1. İlk BBF", type=["docx", "doc", "pdf", "txt"], key="upd_first_bbf")
        revised_bbf = st.file_uploader("2. Revize BBF", type=["docx", "doc", "pdf", "txt"], key="upd_revised_bbf")
        prior_report = st.file_uploader("3. İlk Ön Araştırma Raporu", type=["pdf", "docx", "doc", "txt"], key="upd_prior_report")
    with c2:
        update_reference = st.text_input("DP referans numarası", value="", key="upd_reference")
        update_output_name = st.text_input("Çıktı dosyasının adı", value="Ön Araştırma Raporu_XXXXXX_rev.docx", key="upd_output_name")
        update_cutoff = st.date_input("Araştırma kesim tarihi", value=date.today(), key="upd_cutoff")

    for key, default in {
        "update_analysis": None,
        "update_research": None,
        "update_first_text": None,
        "update_revised_text": None,
        "update_prior_report_text": None,
        "update_prior_report_asset": None,
    }.items():
        if key not in st.session_state:
            st.session_state[key] = default

    if st.button("1. Farkları ve teknik katkıyı analiz et", type="primary", use_container_width=True):
        if first_bbf is None:
            st.error("İlk BBF dosyasını yükleyin.")
        elif revised_bbf is None:
            st.error("Revize BBF dosyasını yükleyin.")
        elif prior_report is None:
            st.error("İlk Ön Araştırma Raporunu yükleyin.")
        else:
            try:
                progress = st.progress(0, text="İlk araştırma konusu okunuyor...")
                first_asset = UploadedAsset(first_bbf.name, first_bbf.getvalue(), first_bbf.type)
                revised_asset = UploadedAsset(revised_bbf.name, revised_bbf.getvalue(), revised_bbf.type)
                report_asset = UploadedAsset(prior_report.name, prior_report.getvalue(), prior_report.type)
                first_text = extract_text_from_asset(first_asset)
                progress.progress(25, text="Revize araştırma konusu okunuyor...")
                revised_text = extract_text_from_asset(revised_asset)
                progress.progress(45, text="İlk ön araştırma raporu ve D1/D2 gerekçeleri okunuyor...")
                report_text = extract_text_from_asset(report_asset)
                progress.progress(65, text="Teknik farklar ve katkılar karşılaştırılıyor...")
                analysis = ask_json(research_update_analysis_prompt(first_text, revised_text, report_text))
                st.session_state.update_analysis = analysis
                st.session_state.update_research = None
                st.session_state.update_first_text = first_text
                st.session_state.update_revised_text = revised_text
                st.session_state.update_prior_report_text = report_text
                st.session_state.update_prior_report_asset = report_asset
                progress.progress(100, text="Fark analizi tamamlandı")
            except Exception as exc:
                st.exception(exc)

    if st.session_state.update_analysis:
        analysis = st.session_state.update_analysis
        st.markdown("### Fark analizi")
        diff_rows = []
        for i, d in enumerate(analysis.get("differences") or [], 1):
            diff_rows.append({
                "No": i,
                "İlk araştırma konusu": d.get("old", ""),
                "Revize araştırma konusu": d.get("new", ""),
                "Teknik katkı": d.get("technical_contribution", ""),
                "Teknik etki": d.get("technical_effect", ""),
                "İlk D1/D2 karşısındaki etkisi": d.get("effect_against_prior_d1_d2", ""),
            })
        if diff_rows:
            st.dataframe(diff_rows, use_container_width=True, hide_index=True)
        st.write(f"**İlk rapordaki D1:** {(analysis.get('prior_d1') or {}).get('number','-')}  |  **D2:** {(analysis.get('prior_d2') or {}).get('number','-')}")
        st.info(analysis.get("preliminary_opinion", ""))

        if st.button("2. Revize konu için yeni patent araştırmasını yap", type="primary", use_container_width=True):
            try:
                progress = st.progress(0, text="Revize teknik farklara göre global araştırma yapılıyor...")
                research = ask_json(
                    research_update_search_prompt(
                        st.session_state.update_revised_text or "",
                        st.session_state.update_prior_report_text or "",
                        analysis,
                        update_cutoff.strftime("%d.%m.%Y"),
                    ),
                    web_search=True,
                )
                docs = research.get("documents") or []
                if len(docs) != 10:
                    raise ValueError(f"Araştırma güncellemede tam 10 doğrulanmış doküman beklenirken {len(docs)} doküman döndü. Araştırmayı tekrar çalıştırın.")
                temp_selection = {
                    "d1": research.get("d1") or {},
                    "d2": research.get("d2"),
                    "comparison_rows_d1": research.get("comparison_rows_d1") or [],
                    "comparison_rows_d2": research.get("comparison_rows_d2") or [],
                }
                validate_research_selection(temp_selection)
                st.session_state.update_research = research
                progress.progress(100, text="Yeni araştırma tamamlandı")
            except Exception as exc:
                st.exception(exc)

    if st.session_state.update_research:
        research = st.session_state.update_research
        st.markdown("### Yeni araştırma sonucu")
        new_docs = research.get("new_documents") or []
        if new_docs:
            st.write("**İlk raporda bulunmayan yeni yakın dokümanlar:**")
            st.dataframe([
                {
                    "Yayın no": d.get("number", ""),
                    "Başlık": d.get("title", ""),
                    "Tarih": d.get("date", ""),
                    "Teknik ilgisi": d.get("technical_relevance", ""),
                }
                for d in new_docs
            ], use_container_width=True, hide_index=True)
        else:
            st.caption("İlk rapordaki dokümanlardan daha yakın yeni bir doküman tespit edilmedi.")

        st.code(research.get("totalpatent_query", ""), language=None)
        st.write(f"**Önerilen nihai D1:** {(research.get('d1') or {}).get('number','-')}  |  **D2:** {(research.get('d2') or {}).get('number','-')}")
        st.write(f"**Yenilik ön sonucu:** {research.get('novelty_result','-')}  |  **Buluş basamağı ön sonucu:** {research.get('inventive_step_result','-')}")
        st.info(research.get("technical_opinion", ""))

        recommended = "Buluş basamağı sağlanmıyor" if str(research.get("inventive_step_result", "")).strip() == "sağlanmaz" else "Buluş basamağı sağlanıyor"
        update_decision = st.radio(
            "Raporu hangi sonuçla hazırlayayım?",
            ["Buluş basamağı sağlanıyor", "Buluş basamağı sağlanmıyor"],
            index=1 if recommended == "Buluş basamağı sağlanmıyor" else 0,
            horizontal=True,
            key="upd_decision",
        )
        st.caption(f"Sistem önerisi: {recommended}")

        if st.button("3. Ön Araştırma Raporunu oluştur", type="primary", use_container_width=True):
            if not update_reference.strip():
                st.error("DP referans numarasını girin.")
            else:
                try:
                    progress = st.progress(0, text="Standart Tip 3 rapor metni hazırlanıyor...")
                    report = ask_json(
                        research_update_report_prompt(
                            st.session_state.update_revised_text or "",
                            st.session_state.update_prior_report_text or "",
                            st.session_state.update_analysis or {},
                            research,
                            update_reference.strip(),
                            update_cutoff.strftime("%d.%m.%Y"),
                            update_decision,
                        )
                    )
                    validate_report_against_selection(report, {"d1": research.get("d1"), "d2": research.get("d2")})
                    validate_research_report_language(report)
                    progress.progress(55, text="D1/D2 özgün patent şekilleri temin ediliyor...")
                    fallbacks = _report_pdf_fallback_figures(st.session_state.update_prior_report_asset)
                    progress.progress(75, text="Bağlayıcı Ön Araştırma Raporu şablonu dolduruluyor...")
                    data = build_research_docx(report, figure_fallbacks=fallbacks)
                    progress.progress(100, text="Rapor hazır")
                    effective_output = update_output_name.replace("XXXXXX", update_reference.strip())
                    st.success("Güncelleme raporu standart Ön Araştırma Raporu formatında oluşturuldu.")
                    st.download_button(
                        "Word raporunu indir",
                        data=data,
                        file_name=safe_output_name(effective_output, "Ön Araştırma Raporu_rev.docx"),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        use_container_width=True,
                    )
                except Exception as exc:
                    st.exception(exc)

