from __future__ import annotations

import io
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
from email import policy
from email.parser import BytesParser
from html import unescape
import xml.etree.ElementTree as ET
import unicodedata
from pathlib import Path
from typing import Any

from docx import Document
from pypdf import PdfReader, PdfWriter
from PIL import Image, ImageOps

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = f"{{{_W_NS}}}"
ET.register_namespace("w", _W_NS)


def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip()).upper()


def _block_text(block: ET.Element) -> str:
    parts: list[str] = []
    for node in block.iter():
        if node.tag == _W + "t" and node.text:
            parts.append(node.text)
        elif node.tag in {_W + "tab"}:
            parts.append("\t")
        elif node.tag in {_W + "br", _W + "cr"}:
            parts.append("\n")
    return "".join(parts).strip()


def _find_heading_index(body_children: list[ET.Element], heading: str) -> int:
    target = _norm(heading)
    for idx, child in enumerate(body_children):
        if child.tag == _W + "sectPr":
            continue
        if _norm(_block_text(child)) == target:
            return idx
    raise ValueError(f"Tarifname Word dosyasında '{heading}' başlığı bulunamadı.")


def _slice_docx(data: bytes, *, start_idx: int, end_idx: int | None) -> bytes:
    """DOCX paket yapısını koruyarak document.xml gövdesini seçilen aralığa indirger."""
    src = io.BytesIO(data)
    out = io.BytesIO()
    with zipfile.ZipFile(src, "r") as zin:
        document_xml = zin.read("word/document.xml")
        root = ET.fromstring(document_xml)
        body = root.find(_W + "body")
        if body is None:
            raise ValueError("Word belgesinin gövdesi okunamadı.")
        children = list(body)
        sect_pr = next((c for c in children if c.tag == _W + "sectPr"), None)
        content = [c for c in children if c.tag != _W + "sectPr"]
        selected = content[start_idx:end_idx]
        if not selected:
            raise ValueError("Word belgesinde seçilen bölüm boş.")
        for child in list(body):
            body.remove(child)
        for child in selected:
            body.append(child)
        if sect_pr is not None:
            body.append(sect_pr)
        new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                payload = new_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, payload)
    return out.getvalue()



_TEMPLATE_COLORS = {"FF0000", "0000FF"}


def _run_color(run: ET.Element) -> str:
    rpr = run.find(_W + "rPr")
    if rpr is None:
        return ""
    color = rpr.find(_W + "color")
    if color is None:
        return ""
    return (color.attrib.get(_W + "val") or "").upper()


def _paragraph_has_visible_content(paragraph: ET.Element) -> bool:
    for node in paragraph.iter():
        local = node.tag.rsplit("}", 1)[-1]
        if local == "t" and (node.text or "").strip():
            return True
        if local in {"drawing", "object", "pict"}:
            return True
    return False


def strip_template_colored_text(data: bytes, *, colors: set[str] | None = None) -> bytes:
    """Kırmızı/mavi şablon açıklamalarını DOCX'ten OOXML düzeyinde kaldırır.

    Siyah teknik içerik, başlıklar, numaralandırma ve belge paket yapısı korunur.
    """
    target = {c.upper() for c in (colors or _TEMPLATE_COLORS)}
    src = io.BytesIO(data)
    out = io.BytesIO()
    with zipfile.ZipFile(src, "r") as zin:
        document_xml = zin.read("word/document.xml")
        root = ET.fromstring(document_xml)
        parent_map = {child: parent for parent in root.iter() for child in parent}
        touched_paragraphs: set[ET.Element] = set()
        for run in list(root.iter(_W + "r")):
            if _run_color(run) not in target:
                continue
            parent = parent_map.get(run)
            if parent is None:
                continue
            paragraph = next((a for a in parent_map_chain(run, parent_map) if a.tag == _W + "p"), None)
            if paragraph is not None:
                touched_paragraphs.add(paragraph)
            parent.remove(run)

        # Yalnız renkli şablon içeriğinden ibaret kalan paragrafları kaldır; böylece
        # PDF sayfa sayımı gerçek başvuru metnine göre yapılır.
        for paragraph in list(touched_paragraphs):
            if _paragraph_has_visible_content(paragraph):
                continue
            parent = parent_map.get(paragraph)
            if parent is not None:
                try:
                    parent.remove(paragraph)
                except ValueError:
                    pass

        new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                payload = new_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, payload)
    return out.getvalue()


def parent_map_chain(node: ET.Element, parent_map: dict[ET.Element, ET.Element]):
    current = node
    while current in parent_map:
        current = parent_map[current]
        yield current


def remove_word_line_numbering(data: bytes) -> bytes:
    """EPATS PDF'lerinde sol kenarda görünen Word otomatik satır numaralarını kaldırır.

    Yalnız section property içindeki w:lnNumType düğümünü siler. Font, stil,
    paragraf, numaralandırma, header/footer, margin ve teknik metne dokunmaz.
    """
    src = io.BytesIO(data)
    out = io.BytesIO()
    with zipfile.ZipFile(src, "r") as zin:
        document_xml = zin.read("word/document.xml")
        root = ET.fromstring(document_xml)
        changed = False
        for sect_pr in root.iter(_W + "sectPr"):
            for node in list(sect_pr):
                if node.tag == _W + "lnNumType":
                    sect_pr.remove(node)
                    changed = True
        new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True) if changed else document_xml
        with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                payload = new_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, payload)
    return out.getvalue()


def _paragraph_numbering(paragraph: ET.Element) -> tuple[str, str] | None:
    ppr = paragraph.find(_W + "pPr")
    if ppr is None:
        return None
    numpr = ppr.find(_W + "numPr")
    if numpr is None:
        return None
    num_id_node = numpr.find(_W + "numId")
    ilvl_node = numpr.find(_W + "ilvl")
    if num_id_node is None:
        return None
    return (
        num_id_node.attrib.get(_W + "val", ""),
        ilvl_node.attrib.get(_W + "val", "0") if ilvl_node is not None else "0",
    )


def count_claims_from_docx(data: bytes) -> int:
    """İSTEMLER bölümündeki gerçek istem sayısını Word liste yapısından belirler."""
    with zipfile.ZipFile(io.BytesIO(data), "r") as zin:
        root = ET.fromstring(zin.read("word/document.xml"))
        numbering_root = ET.fromstring(zin.read("word/numbering.xml")) if "word/numbering.xml" in zin.namelist() else None
    num_to_abstract: dict[str, str] = {}
    if numbering_root is not None:
        for num in numbering_root.findall(_W + "num"):
            num_id = num.attrib.get(_W + "numId", "")
            abs_node = num.find(_W + "abstractNumId")
            if num_id and abs_node is not None:
                num_to_abstract[num_id] = abs_node.attrib.get(_W + "val", "")
    body = root.find(_W + "body")
    if body is None:
        return 0
    content = [c for c in list(body) if c.tag != _W + "sectPr"]
    try:
        claims_idx = _find_heading_index(content, "İSTEMLER")
    except ValueError:
        claims_idx = -1
    claim_nodes = content[claims_idx + 1 :] if claims_idx >= 0 else content

    # Şablonda ilk gerçek numaralı paragraf bağımsız istemin numId'sini verir;
    # alt unsur listeleri farklı numId kullanır. Aynı numId'deki ilvl=0 paragraflar
    # bağımsız + bağımlı istemlerin sayısıdır.
    claim_num_id = ""
    for node in claim_nodes:
        if node.tag != _W + "p" or not _block_text(node).strip():
            continue
        numbering = _paragraph_numbering(node)
        if numbering and numbering[1] == "0":
            claim_num_id = numbering[0]
            break
    if claim_num_id:
        claim_abstract = num_to_abstract.get(claim_num_id, "")
        if claim_abstract:
            count = 0
            for node in claim_nodes:
                if node.tag != _W + "p" or not _block_text(node).strip():
                    continue
                numbering = _paragraph_numbering(node)
                if not numbering or numbering[1] != "0":
                    continue
                if num_to_abstract.get(numbering[0], "") == claim_abstract:
                    count += 1
            if count:
                return count
        count = sum(
            1
            for node in claim_nodes
            if node.tag == _W + "p"
            and _block_text(node).strip()
            and _paragraph_numbering(node) == (claim_num_id, "0")
        )
        if count:
            return count

    # Liste numaralandırması kaybolmuş dosyalar için metinsel geri dönüş.
    text = "\n".join(_block_text(n) for n in claim_nodes if n.tag == _W + "p")
    explicit = re.findall(r"(?mi)^\s*(\d+)\s*[.\-)]+\s+", text)
    if explicit:
        return len(set(explicit))
    dependent = re.findall(r"(?mi)^\s*İstem\s+\d+(?:['’`]?e|['’`]?a)\s+uygun", text)
    if dependent:
        return 1 + len(dependent)
    # Bazı Word/PDF dönüşümlerinde liste numarası metne düşmez ancak her bağımlı
    # istem ayrı paragrafta "İstem X..." diye başlar. Aynı paragrafı iki kez sayma.
    claimish = [
        _block_text(n).strip() for n in claim_nodes
        if n.tag == _W + "p" and _block_text(n).strip()
    ]
    dep_count = sum(1 for t in claimish if re.match(r"(?i)^istem\s+\d+", t))
    if dep_count:
        return 1 + dep_count
    return 0


def pdf_page_count(data: bytes) -> int:
    return len(PdfReader(io.BytesIO(data)).pages)


def epats_document_metrics(
    specification_data: bytes,
    pdfs: dict[str, bytes],
    *,
    specification_name: str = "Tarifname.docx",
) -> dict[str, Any]:
    suffix = Path(specification_name).suffix.lower()
    claims = 0
    if suffix in {".docx", ".doc"}:
        docx = specification_data if suffix == ".docx" else _libreoffice_to_docx(specification_data, specification_name)
        cleaned = strip_template_colored_text(docx)
        split_docs = split_patent_docx(cleaned, clean_template_colors=False)
        claims = count_claims_from_docx(split_docs["Istemler.docx"])
    elif suffix == ".pdf" and pdfs.get("Istemler.pdf"):
        claims = count_claims_from_pdf(pdfs["Istemler.pdf"])
    spec_pages = pdf_page_count(pdfs["Tarifname.pdf"]) if pdfs.get("Tarifname.pdf") else 0
    figures_pages = pdf_page_count(pdfs["Sekiller.pdf"]) if pdfs.get("Sekiller.pdf") else 0
    abstract_ok = bool(pdfs.get("Ozet.pdf")) and pdf_page_count(pdfs["Ozet.pdf"]) > 0
    return {
        "specification_pages": spec_pages,
        "claim_count": claims,
        "abstract_present": abstract_ok,
        "figures_pages": figures_pages,
        "codes": {
            "specification": f"T-{spec_pages}" if spec_pages else "T-?",
            "claims": f"İ-{claims}" if claims else "İ-?",
            "abstract": "Ö" if abstract_ok else "Ö-?",
            "figures": f"Ş-{figures_pages}" if figures_pages else "Ş-",
        },
    }


def _docx_source_text(data: bytes) -> str:
    """DOCX metnini paragraf/tablo sırasını ve tablo hücrelerini koruyarak çıkarır."""
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    doc = Document(io.BytesIO(data))
    parts: list[str] = []
    table_index = 0
    for child in doc.element.body.iterchildren():
        local = child.tag.rsplit("}", 1)[-1]
        if local == "p":
            p = Paragraph(child, doc)
            text = p.text.strip()
            if text:
                parts.append(text)
        elif local == "tbl":
            table_index += 1
            table = Table(child, doc)
            parts.append(f"[[TABLO {table_index}]]")
            for row in table.rows:
                vals: list[str] = []
                for cell in row.cells:
                    value = re.sub(r"\s*\n\s*", " / ", cell.text or "").strip()
                    # Birleşik hücrelerde python-docx aynı hücreyi birden çok kez döndürebilir.
                    if value and (not vals or _plain_norm(value) != _plain_norm(vals[-1])):
                        vals.append(value)
                    elif not value and not vals:
                        vals.append("")
                if any(v for v in vals):
                    parts.append("\t".join(vals))
    return "\n".join(parts)


def _html_to_text(value: str) -> str:
    value = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", value or "")
    value = re.sub(r"(?i)<br\s*/?>", "\n", value)
    value = re.sub(r"(?i)</p\s*>", "\n", value)
    value = re.sub(r"(?i)</(?:div|tr|li|h[1-6])\s*>", "\n", value)
    value = re.sub(r"(?s)<[^>]+>", " ", value)
    value = unescape(value)
    value = re.sub(r"[ \t]+", " ", value)
    value = re.sub(r"\n\s*\n+", "\n", value)
    return value.strip()


def _attachment_text(filename: str, data: bytes) -> str:
    """E-posta eklerindeki desteklenen metin belgelerini güvenli biçimde oku."""
    suffix = Path(filename or "").suffix.lower()
    if suffix in {".docx", ".doc", ".pdf", ".txt", ".md", ".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp"}:
        try:
            return extract_application_source_text(filename, data)
        except Exception:
            return ""
    return ""


def _eml_text(data: bytes) -> str:
    msg = BytesParser(policy=policy.default).parsebytes(data)
    head = [
        f"Konu: {msg.get('subject', '')}",
        f"Kimden: {msg.get('from', '')}",
        f"Kime: {msg.get('to', '')}",
        f"Cc: {msg.get('cc', '')}",
        f"Tarih: {msg.get('date', '')}",
    ]
    plain_parts: list[str] = []
    html_parts: list[str] = []
    attachment_parts: list[str] = []
    if msg.is_multipart():
        for part in msg.walk():
            disposition = part.get_content_disposition()
            filename = part.get_filename() or ""
            if disposition == "attachment" or filename:
                payload = part.get_payload(decode=True) or b""
                att = _attachment_text(filename, payload)
                if att:
                    attachment_parts.append(f"[[E-POSTA EKI: {filename}]]\n{att}")
                continue
            ctype = part.get_content_type()
            if ctype == "text/plain":
                try:
                    plain_parts.append(str(part.get_content()))
                except Exception:
                    payload = part.get_payload(decode=True) or b""
                    plain_parts.append(payload.decode(part.get_content_charset() or "utf-8", errors="replace"))
            elif ctype == "text/html":
                try:
                    html_parts.append(_html_to_text(str(part.get_content())))
                except Exception:
                    pass
    else:
        try:
            content = msg.get_content()
            (html_parts if msg.get_content_type() == "text/html" else plain_parts).append(
                _html_to_text(str(content)) if msg.get_content_type() == "text/html" else str(content)
            )
        except Exception:
            pass
    body_parts = plain_parts or html_parts
    return "\n".join([x for x in head + body_parts + attachment_parts if x and x.strip()]).strip()


def extract_application_source_text(filename: str, data: bytes) -> str:
    """Beyan formu/yazı/e-posta gibi başvuru bilgi kaynaklarından metin çıkarır."""
    suffix = Path(filename).suffix.lower()
    if suffix == ".docx":
        text = _docx_source_text(data)
    elif suffix == ".doc":
        with tempfile.TemporaryDirectory() as td:
            src = Path(td) / Path(filename).name
            src.write_bytes(data)
            proc = subprocess.run(["antiword", str(src)], stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
            text = proc.stdout.decode("utf-8", errors="replace") if proc.returncode == 0 else ""
    elif suffix == ".pdf":
        text = "\n".join((page.extract_text() or "") for page in PdfReader(io.BytesIO(data)).pages)
    elif suffix in {".txt", ".md"}:
        text = data.decode("utf-8", errors="replace")
    elif suffix == ".eml":
        text = _eml_text(data)
    elif suffix == ".msg":
        try:
            import extract_msg  # type: ignore
        except ImportError as exc:
            raise ValueError(".msg e-posta dosyasını okuyabilmek için extract-msg paketi kurulu olmalıdır.") from exc
        with tempfile.TemporaryDirectory() as td:
            src = Path(td) / Path(filename).name
            src.write_bytes(data)
            msg = extract_msg.Message(str(src))
            parts = [
                f"Konu: {getattr(msg, 'subject', '') or ''}",
                f"Kimden: {getattr(msg, 'sender', '') or ''}",
                f"Kime: {getattr(msg, 'to', '') or ''}",
                f"Cc: {getattr(msg, 'cc', '') or ''}",
                f"Tarih: {getattr(msg, 'date', '') or ''}",
                getattr(msg, "body", "") or "",
            ]
            # Outlook mesajının içindeki Word/PDF/TXT ve görsel ekleri de bilgi kaynağına kat.
            for att in getattr(msg, "attachments", []) or []:
                att_name = str(getattr(att, "longFilename", None) or getattr(att, "shortFilename", None) or "")
                att_data = getattr(att, "data", None)
                if att_name and isinstance(att_data, (bytes, bytearray)):
                    att_text = _attachment_text(att_name, bytes(att_data))
                    if att_text:
                        parts.append(f"[[E-POSTA EKI: {att_name}]]\n{att_text}")
            text = "\n".join(x for x in parts if x and str(x).strip())
            try:
                msg.close()
            except Exception:
                pass
    elif suffix in {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp"}:
        try:
            import pytesseract  # type: ignore
        except ImportError as exc:
            raise ValueError("Resimden metin okuyabilmek için yerel OCR bileşeni kurulu değildir.") from exc
        try:
            image = Image.open(io.BytesIO(data))
            image = ImageOps.exif_transpose(image).convert("L")
            image = ImageOps.autocontrast(image)
            # Türkçe + İngilizce OCR yerel çalışır; OpenAI/API kredisi tüketmez.
            text = pytesseract.image_to_string(image, lang="tur+eng", config="--psm 6")
        except Exception as exc:
            raise ValueError(f"Resim OCR ile okunamadı: {exc}") from exc
    else:
        raise ValueError("Bilgi kaynağı için desteklenen türler: .docx, .doc, .pdf, .txt, .md, .eml, .msg, .png, .jpg, .jpeg, .webp, .tif, .tiff, .bmp")
    text = (text or "").replace("\x00", " ").strip()
    if not text:
        raise ValueError(f"{filename} dosyasından metin çıkarılamadı.")
    return text



# -----------------------------------------------------------------------------
# AI'SIZ / DETERMINISTIK BASVURU BILGISI CIKARIMI
# -----------------------------------------------------------------------------

_LABEL_SPLIT_RE = re.compile(r"^\s*([^:\t|]{2,80}?)\s*(?::|\t|\|)\s*(.*?)\s*$")


def _plain_norm(value: str) -> str:
    value = (value or "").casefold()
    value = value.translate(str.maketrans({"ı": "i", "ş": "s", "ğ": "g", "ü": "u", "ö": "o", "ç": "c"}))
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def _clean_value(value: str) -> str:
    value = re.sub(r"\s+", " ", (value or "").strip(" \t|:;-"))
    return value.strip()


def _looks_like_company(name: str) -> bool:
    n = _plain_norm(name)
    return bool(re.search(r"\b(a s|anonim|ltd|limited|sanayi|ticaret|holding|sirketi|company|inc|corp|llc)\b", n))


def _entity_type(name: str, explicit: str = "") -> str:
    e = _plain_norm(explicit)
    if "tuzel" in e or "kurum" in e or "sirket" in e:
        return "Tüzel kişi"
    if "gercek" in e or "kisi" in e:
        return "Gerçek kişi"
    if _looks_like_company(name):
        return "Tüzel kişi"
    return ""


def _canonical_label(label: str) -> str:
    n = _plain_norm(label)
    if not n:
        return ""

    # Form başlıklarında sık geçen ek/sonekleri normalleştir; anlamı bozacak
    # genel kelimeleri tamamen silmek yerine rol tespitinde esnek davran.
    applicant_role = any(x in n for x in [
        "hak sahibi", "basvuru sahibi", "basvuran", "applicant", "patent sahibi",
    ])
    inventor_role = any(x in n for x in [
        "bulus sahibi", "bulusu yapan", "bulus yapan", "buluscu", "mucit", "inventor",
    ])
    name_terms = [
        "adi soyadi", "ad soyad", "adi ve soyadi", "ad ve soyad", "ad soyadi",
        "adi unvani", "ad unvan", "adi ve unvani", "unvani", "unvan", "isim", "name",
    ]
    identity_terms = [
        "tckn", "t c kimlik", "tc kimlik", "tc kimlik no", "tc kimlik numarasi", "t c kimlik numarasi",
        "vkn", "vergi no", "vergi numarasi", "vergi kimlik no", "vergi kimlik numarasi",
        "kimlik no", "kimlik numarasi", "identity", "tax number",
    ]

    # Rol + alan birleşik etiketleri: "Başvuru Sahibinin Adı/Unvanı",
    # "Buluşu Yapanın Adresi", "Hak Sahibi TCKN/VKN" vb.
    if applicant_role or inventor_role:
        role = "applicant" if applicant_role else "inventor"
        if any(x in n for x in ["adres", "address", "tebligat adres", "ikametgah"]):
            return role + "_address"
        if any(x in n for x in ["ulke", "uyruk", "tabiyet", "country", "milliyet"]):
            return role + "_country"
        if any(x in n for x in ["il ilce", "il ilçe", "il/ilce", "il/ilçe"]):
            return role + "_location"
        if re.search(r"\bilce\b", n):
            return role + "_district"
        if re.search(r"\b(il|sehir|city)\b", n):
            return role + "_city"
        if any(x in n for x in identity_terms):
            return role + "_identity"
        if any(x in n for x in ["dogum tarihi", "birth date"]):
            return role + "_birth_date"
        if any(x in n for x in ["e posta", "eposta", "email", "e mail"]):
            return role + "_email"
        if any(x in n for x in ["telefon", "gsm", "cep telefonu", "phone"]):
            return role + "_phone"
        if any(x in n for x in name_terms):
            return role + "_name"
        if applicant_role and any(x in n for x in ["kisi turu", "tuzel", "gercek", "entity type", "hukuki nitelik"]):
            return "applicant_entity_type"

    aliases = {
        "application_kind": [
            "basvuru turu", "basvuru tipi", "koruma turu", "basvuru sekli", "koruma sekli",
            "patent faydali model tercihi", "patent veya faydali model",
        ],
        "reference": [
            "dp referans", "dp ref", "dp no", "dosya referansi", "dosya referans no", "dosya no",
            "referans no", "referans numarasi", "referans", "is no", "is numarasi", "dosya kodu",
        ],
        "invention_title": [
            "bulus basligi", "bulusun basligi", "bulus adi", "bulusun adi", "buluşun adı",
            "invention title", "baslik",
        ],
        "applicant": [
            "hak sahibi", "hak sahipleri", "basvuru sahibi", "basvuru sahipleri", "basvuran",
            "applicant", "applicants", "hak sahibi bilgileri", "basvuru sahibi bilgileri",
        ],
        "inventor": [
            "bulus sahibi", "bulus sahipleri", "buluscu", "buluscular", "bulusu yapan",
            "bulusu yapanlar", "bulus yapan", "mucit", "mucitler", "inventor", "inventors",
            "bulus sahibi bilgileri", "bulusu yapan bilgileri",
        ],
        "name": name_terms,
        "identity": identity_terms,
        "address": [
            "adres", "adresi", "tebligat adresi", "ikamet adresi", "ikametgah", "merkez adresi",
            "yazisma adresi", "address",
        ],
        "country": ["ulke", "ulkesi", "uyruk", "uyrugu", "tabiyet", "tabiyeti", "milliyet", "milliyeti", "country"],
        "location": ["il ilce", "il/ilce", "il / ilce", "il ilçe", "il/ilçe", "il / ilçe", "il ve ilce", "city district"],
        "city": ["il", "ili", "sehir", "sehri", "city"],
        "district": ["ilce", "ilcesi", "ilçe", "ilçesi", "district"],
        "birth_date": ["dogum tarihi", "doğum tarihi", "birth date", "date of birth"],
        "fax": ["faks", "fax", "faks no", "fax no"],
        "postal_code": ["posta kodu", "postal code", "zip code"],
        "authorized_person": ["yetkili", "yetkili kisi", "yetkili kişi", "temsilci", "contact person"],
        "website": ["web", "web sitesi", "internet adresi", "website"],
        "entity_type": [
            "kisi turu", "hak sahibi turu", "basvuru sahibi turu", "tuzel gercek kisi",
            "gercek tuzel", "hukuki nitelik", "entity type",
        ],
        "email": ["e posta", "eposta", "e mail", "email", "mail adresi", "elektronik posta"],
        "phone": ["telefon", "telefonu", "telefon no", "telefon numarasi", "gsm", "cep telefonu", "ev telefonu", "is telefonu", "ev is telefonu", "phone"],
        "priority": ["ruchan", "ruchan durumu", "ruchan talebi", "priority", "priority claim"],
        "priority_country": ["ruchan ulkesi", "priority country"],
        "priority_number": ["ruchan numarasi", "ruchan no", "ruchan basvuru no", "priority number"],
        "priority_date": ["ruchan tarihi", "priority date"],
    }
    order = (
        "priority_country", "priority_number", "priority_date", "application_kind", "invention_title",
        "reference", "entity_type", "applicant", "inventor", "name", "identity", "address", "country",
        "location", "city", "district", "birth_date", "authorized_person", "postal_code", "email", "phone", "fax", "website", "priority",
    )
    for key in order:
        for alias in aliases[key]:
            a = _plain_norm(alias)
            if n == a or n.startswith(a + " ") or n.endswith(" " + a):
                return key
    return ""



def _is_instruction_text(value: str) -> bool:
    n = _plain_norm(value)
    if not n:
        return False
    markers = [
        "not basvuru sahibinin", "birden fazla olmasi durumunda", "ayri ayri duzenlenmelidir",
        "bu alan", "doldurulmalidir", "zorunlu alan", "gerekiyorsa", "aciklama", "ornek",
        "basvuru sahibinin birden fazla", "bulusu yapan birden fazla", "her bir basvuru sahibi",
    ]
    return any(x in n for x in markers)


def _is_role_section_heading(raw_label: str, canonical: str) -> bool:
    n = _plain_norm(raw_label)
    if canonical not in {"applicant", "inventor"}:
        return False
    # Resmî beyan formlarında bölüm başlığı bazen yalnız "... SAHİBİ(LERİ)"
    # biçimindedir ve hemen yanında açıklama/not bulunur. Bu da bölüm başlığıdır.
    return any(x in n for x in [
        "bilgi", "bilgileri", "bilgisi", "information", "detay", "detaylari",
        "sahipleri", "yapanlar", "mucitler", "buluscular",
    ])


def _line_label_value_pairs(line: str) -> list[tuple[str, str, str]]:
    """Bir görsel/tablo satırındaki bir veya birden çok etiket-değer çiftini çıkar.

    Dönüş: (ham etiket, kanonik etiket, değer). Tablo satırlarında
    "Hak Sahibi | ABC A.Ş. | Ülke | Türkiye" gibi birden çok çift desteklenir.
    """
    line = (line or "").strip()
    if not line or line.startswith("[[TABLO") or line.startswith("[[E-POSTA EKI"):
        return []

    # Önce tablo/hücre ayraçlarını değerlendir. Hücre içindeki " / " normal
    # metin olarak bırakılır; yalnız tab ve belirgin dikey çizgiler ayraçtır.
    tokens = [x.strip() for x in re.split(r"\t+|\s+\|\s+", line) if x.strip()]
    if len(tokens) >= 2:
        pairs: list[tuple[str, str, str]] = []
        i = 0
        while i < len(tokens):
            label = _canonical_label(tokens[i])
            if label:
                j = i + 1
                vals: list[str] = []
                while j < len(tokens) and not _canonical_label(tokens[j]):
                    vals.append(tokens[j])
                    j += 1
                value = _clean_value(" / ".join(vals))
                if label in {"applicant", "inventor"} and _is_instruction_text(value):
                    value = ""
                pairs.append((tokens[i], label, value))
                i = max(j, i + 1)
            else:
                i += 1
        if pairs:
            return pairs

    m = _LABEL_SPLIT_RE.match(line)
    if m:
        raw_label = m.group(1)
        label = _canonical_label(raw_label)
        if label:
            value = _clean_value(m.group(2))
            if label in {"applicant", "inventor"} and _is_instruction_text(value):
                value = ""
            return [(raw_label, label, value)]
    # PDF/antiword çıktısında tablo sütunları bazen ayraçsız tek satıra
    # düşer: "Hak Sahibi ABC A.Ş." gibi. Yalnız açık ve sabit etiket
    # başlangıçlarında ayraçsız geri dönüş uygula.
    inline_patterns = [
        (r"^(hak\s+sahibi|başvuru\s+sahibi|başvuru\s+sahibinin|hak\s+sahibinin)\s+(.+)$", "applicant"),
        (r"^(buluş\s+sahibi|buluşu\s+yapan|buluşçu|mucit)\s+(.+)$", "inventor"),
        (r"^(buluş\s+başlığı|buluşun\s+başlığı|buluş\s+adı)\s+(.+)$", "invention_title"),
        (r"^(başvuru\s+türü|başvuru\s+tipi|koruma\s+türü)\s+(.+)$", "application_kind"),
        (r"^(dp\s*(?:ref(?:erans)?|no)|dosya\s+(?:referansı|no)|referans\s+no)\s+(.+)$", "reference"),
    ]
    for pattern, default_label in inline_patterns:
        mm = re.match(pattern, line, flags=re.IGNORECASE)
        if mm:
            value = _clean_value(mm.group(2))
            if _plain_norm(value) not in {"bilgi", "bilgileri", "bilgisi"}:
                return [(mm.group(1), default_label, value)]

    label = _canonical_label(line)
    if label:
        return [(line, label, "")]
    return []


def reference_from_filename(filename: str) -> str:
    """Tarifname dosya adındaki DP/ofis referansını güvenli biçimde çıkar.

    Örn. Tarifname_181176_rev3.docx -> 181176. Yalnız açık bir numaralı
    referans bulunduğunda döner; tarih/sürüm gibi kısa parçaları seçmez.
    """
    stem = Path(filename or "").stem
    # DP181176, DP-181176, Tarifname_181176, 181176_Tarifname gibi örnekler.
    explicit = re.search(r"(?i)\bdp\s*[-_ ]?\s*(\d{4,12})\b", stem)
    if explicit:
        return explicit.group(1)
    candidates = re.findall(r"(?<!\d)(\d{5,12})(?!\d)", stem)
    if candidates:
        # En uzun aday; eşitlikte soldaki. 5.4.48 gibi sürüm parçaları zaten
        # noktalı/kısa olduğundan bu kalıba girmez.
        return sorted(enumerate(candidates), key=lambda x: (-len(x[1]), x[0]))[0][1]
    return ""


def _add_other_information(result: dict[str, Any], label: str, value: str, source: str) -> None:
    label = _clean_value(label)
    value = _clean_value(value)
    if not value:
        return
    key = (_plain_norm(label), _plain_norm(value), source)
    for row in result.get("other_information", []):
        if (_plain_norm(row.get("label", "")), _plain_norm(row.get("value", "")), row.get("source", "")) == key:
            return
    result["other_information"].append({"label": label or "Diğer bilgi", "value": value, "source": source})


def _extract_contact_information(result: dict[str, Any], text: str, source: str) -> None:
    """Serbest mail/yazı içindeki açık iletişim bilgilerini 'diğer bilgiler'e al."""
    for email in sorted(set(re.findall(r"(?i)\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", text or ""))):
        _add_other_information(result, "E-posta", email, source)
    phone_re = re.compile(r"(?<!\d)(?:\+?90\s*)?(?:\(?0?\d{3}\)?[ .-]*)\d{3}[ .-]*\d{2}[ .-]*\d{2}(?!\d)")
    for match in phone_re.finditer(text or ""):
        phone = match.group(0)
        context = _plain_norm((text or "")[max(0, match.start()-24):match.end()+24])
        # VKN/TCKN gibi 10/11 haneli kimlikleri telefon sanma.
        if any(x in context for x in ["vkn", "vergi no", "vergi numarasi", "tckn", "tc kimlik", "kimlik no"]):
            continue
        _add_other_information(result, "Telefon", phone, source)

def _split_people_value(value: str) -> list[str]:
    value = _clean_value(value)
    if not value:
        return []
    # Birden fazla kişi açık biçimde noktalı virgül, satır içi numara veya " ve " ile verilmişse ayır.
    parts = re.split(r"\s*;\s*|\s+\d+[.)]\s+", value)
    if len(parts) == 1 and " ve " in value.casefold() and not _looks_like_company(value):
        candidate = re.split(r"\s+ve\s+", value, flags=re.IGNORECASE)
        if 1 < len(candidate) <= 4 and all(len(x.split()) >= 2 for x in candidate):
            parts = candidate
    return [p.strip() for p in parts if p.strip()]


def _new_person(name: str, source: str, *, applicant: bool) -> dict[str, str]:
    row = {
        "identity": "",
        "name": _clean_value(name),
        "country": "",
        "city": "",
        "district": "",
        "address": "",
        "email": "",
        "phone": "",
        "birth_date": "",
        "source": source,
    }
    if applicant:
        row["entity_type"] = _entity_type(name)
    return row


def _split_city_district(value: str) -> tuple[str, str]:
    value = _clean_value(value)
    if not value:
        return "", ""
    parts = [x.strip() for x in re.split(r"[/|,]", value, maxsplit=1)]
    if len(parts) == 2 and all(parts):
        return parts[0], parts[1]
    return "", value


def _apply_person_field(row: dict[str, str], field: str, value: str, *, applicant: bool) -> None:
    value = _clean_value(value)
    if not value:
        return
    if field == "identity":
        row["identity"] = _valid_identity(value)
    elif field == "entity_type" and applicant:
        row["entity_type"] = _entity_type(row.get("name", ""), value)
    elif field == "location":
        city, district = _split_city_district(value)
        if city and not row.get("city"):
            row["city"] = city
        if district and not row.get("district"):
            row["district"] = district
    elif field == "district":
        city, district = _split_city_district(value)
        if city and _plain_norm(city) in _TURKEY_PROVINCES and not row.get("city"):
            row["city"] = city
        if not row.get("district"):
            row["district"] = district or value
    elif field == "birth_date":
        # Form başlığıyla komşu hücre numaralarını doğum tarihi sanma.
        if re.search(r"(?<!\d)(?:0?[1-9]|[12]\d|3[01])[./-](?:0?[1-9]|1[0-2])[./-](?:19|20)\d{2}(?!\d)", value):
            row["birth_date"] = value
    elif field in {"address", "country", "city", "email", "phone"}:
        if not row.get(field):
            row[field] = value


_TURKEY_PROVINCES = {
    _plain_norm(x) for x in """Adana Adıyaman Afyonkarahisar Ağrı Aksaray Amasya Ankara Antalya Ardahan Artvin Aydın Balıkesir Bartın Batman Bayburt Bilecik Bingöl Bitlis Bolu Burdur Bursa Çanakkale Çankırı Çorum Denizli Diyarbakır Düzce Edirne Elazığ Erzincan Erzurum Eskişehir Gaziantep Giresun Gümüşhane Hakkari Hatay Iğdır Isparta İstanbul İzmir Kahramanmaraş Karabük Karaman Kars Kastamonu Kayseri Kırıkkale Kırklareli Kırşehir Kilis Kocaeli Konya Kütahya Malatya Manisa Mardin Mersin Muğla Muş Nevşehir Niğde Ordu Osmaniye Rize Sakarya Samsun Siirt Sinop Sivas Şanlıurfa Şırnak Tekirdağ Tokat Trabzon Tunceli Uşak Van Yalova Yozgat Zonguldak""".split()
}


def _fill_country_from_explicit_location(row: dict[str, str]) -> None:
    if row.get("country"):
        return
    naddr = _plain_norm(row.get("address", ""))
    if "turkiye" in naddr or re.search(r"\bturkey\b", naddr):
        row["country"] = "Türkiye"
        return
    city = _plain_norm(row.get("city", ""))
    if city in _TURKEY_PROVINCES:
        row["country"] = "Türkiye"



def _append_unique_person(target: list[dict[str, str]], row: dict[str, str], *, applicant: bool) -> dict[str, str]:
    name_key = _plain_norm(row.get("name", ""))
    id_key = re.sub(r"\D", "", row.get("identity", ""))
    for existing in target:
        existing_id = re.sub(r"\D", "", existing.get("identity", ""))
        if (id_key and existing_id and id_key == existing_id) or (name_key and name_key == _plain_norm(existing.get("name", ""))):
            for key in ("identity", "country", "city", "district", "address", "email", "phone", "birth_date"):
                if not existing.get(key) and row.get(key):
                    existing[key] = row[key]
            if applicant and not existing.get("entity_type"):
                existing["entity_type"] = row.get("entity_type") or _entity_type(existing.get("name", ""))
            if row.get("source") and row["source"] not in (existing.get("source") or "").split("; "):
                existing["source"] = "; ".join(x for x in [existing.get("source", ""), row["source"]] if x)
            return existing
    target.append(row)
    return row


def _specification_title(specification_text: str) -> str:
    lines = [_clean_value(x) for x in (specification_text or "").splitlines()]
    lines = [x for x in lines if x]
    for idx, line in enumerate(lines[:15]):
        if _plain_norm(line) == "tarifname":
            for cand in lines[idx + 1: idx + 5]:
                n = _plain_norm(cand)
                if n and n not in {"teknik alan", "onceki teknik", "bulusun kisa aciklamasi"} and not cand.startswith("Araştırma raporunun"):
                    return cand
    # Başlık etiketi olan tarifnamelerde geri dönüş.
    for line in lines[:25]:
        m = _LABEL_SPLIT_RE.match(line)
        if m and _canonical_label(m.group(1)) == "invention_title":
            return _clean_value(m.group(2))
    return ""


def _kind_from_text(text: str) -> str:
    n = _plain_norm(text)
    if re.search(r"\bfaydali model\b", n):
        return "Faydalı Model"
    if re.search(r"\bpatent\b", n):
        return "Patent"
    return ""


def _priority_status(value: str) -> str:
    n = _plain_norm(value)
    if not n:
        return "Belirsiz"
    negative = ["yok", "hayir", "talep edilmiyor", "talep edilmeyecek", "bulunmuyor", "mevcut degil", "no"]
    positive = ["var", "evet", "talep ediliyor", "talep edilecek", "mevcut", "yes"]
    if any(x in n for x in negative):
        return "Yok"
    if any(x in n for x in positive):
        return "Var"
    return "Belirsiz"


def _valid_identity(value: str) -> str:
    digits = re.sub(r"\D", "", value or "")
    return digits if len(digits) in {10, 11} else _clean_value(value)


def _set_singleton(result: dict[str, Any], key: str, value: str, source: str, seen: dict[str, tuple[str, str]]) -> None:
    value = _clean_value(value)
    if not value:
        return
    prior = seen.get(key)
    if prior and _plain_norm(prior[0]) != _plain_norm(value):
        result["conflicts"].append(f"{key}: '{prior[0]}' ({prior[1]}) / '{value}' ({source})")
        return
    seen[key] = (value, source)
    if key == "priority.status":
        result["priority"]["status"] = value
        result["priority"]["source"] = source
    elif key.startswith("priority."):
        result["priority"][key.split(".", 1)[1]] = value
        result["priority"]["source"] = source
    else:
        result[key] = value


def extract_application_information_rule_based(
    source_blocks: list[tuple[str, str]], *, specification_text: str = "", specification_filename: str = ""
) -> dict[str, Any]:
    """Başvuru verilerini OpenAI/API kullanmadan belge yapısı + açık metin kurallarıyla çıkarır.

    Bilgi uydurmaz. Tablo hücreleri, bölüm başlıkları, etiket/değer satırları ve açık
    e-posta cümleleri desteklenir. Bulunamayan alan ön kontrol kapısı tarafından bloke edilir.
    """
    result: dict[str, Any] = {
        "application_kind": "",
        "reference": "",
        "invention_title": "",
        "applicants": [],
        "inventors": [],
        "priority": {"status": "Belirsiz", "country": "", "number": "", "date": "", "source": ""},
        "other_information": [],
        "conflicts": [],
        "source_files_used": [],
        "field_sources": {},
    }
    seen: dict[str, tuple[str, str]] = {}

    def ensure_person(role: str, source: str, *, force_new: bool = False) -> dict[str, str]:
        target = result["applicants"] if role == "applicant" else result["inventors"]
        if force_new or not target:
            row = _new_person("", source, applicant=role == "applicant")
            target.append(row)
            return row
        return target[-1]

    for source, raw_text in source_blocks:
        text = (raw_text or "").replace("\r", "\n")
        result["source_files_used"].append(source)
        _extract_contact_information(result, text, source)

        current_role = ""
        current_person: dict[str, str] | None = None
        pending_label = ""

        lines = [x.strip(" \r") for x in text.splitlines()]
        for line in lines:
            if not line or line.startswith("[[TABLO") or line.startswith("[[E-POSTA EKI"):
                continue

            pairs = _line_label_value_pairs(line)
            if not pairs and pending_label:
                # Başlık bir satır, değer sonraki satır. Ancak sonraki satır başka
                # bir tablo/bölüm başlığıysa onu değer sanma.
                if not _canonical_label(line):
                    pairs = [(pending_label, pending_label, _clean_value(line))]
                    pending_label = ""
            if not pairs:
                nline = _plain_norm(line)
                if "ruchan" in nline and result["priority"]["status"] == "Belirsiz":
                    status = _priority_status(line)
                    if status != "Belirsiz":
                        _set_singleton(result, "priority.status", status, source, seen)
                if not result["application_kind"] and any(x in nline for x in ["patent basvuru", "faydali model basvuru", "faydali model olarak", "patent olarak"]):
                    kind = _kind_from_text(line)
                    if kind:
                        _set_singleton(result, "application_kind", kind, source, seen)
                continue

            for raw_label, label, value in pairs:
                # Bölüm başlığı: "HAK SAHİBİ BİLGİLERİ" / "BULUŞU YAPAN BİLGİLERİ".
                if label in {"applicant", "inventor"} and not value and _is_role_section_heading(raw_label, label):
                    current_role = label
                    current_person = None
                    pending_label = ""
                    continue

                # Tek satır başlık + sonraki satır değer formatı.
                if not value and label in {
                    "application_kind", "reference", "invention_title", "applicant", "inventor",
                    "priority", "priority_country", "priority_number", "priority_date",
                    "applicant_name", "inventor_name", "applicant_identity", "inventor_identity",
                    "applicant_address", "inventor_address", "applicant_country", "inventor_country",
                    "applicant_city", "inventor_city", "name", "identity", "address", "country", "city",
                }:
                    pending_label = label
                    if label.startswith("applicant") or label == "applicant":
                        current_role = "applicant"
                    elif label.startswith("inventor") or label == "inventor":
                        current_role = "inventor"
                    continue

                if label == "application_kind":
                    kind = _kind_from_text(value)
                    if kind:
                        _set_singleton(result, "application_kind", kind, source, seen)
                    continue
                if label == "reference":
                    _set_singleton(result, "reference", value, source, seen)
                    continue
                if label == "invention_title":
                    _set_singleton(result, "invention_title", value, source, seen)
                    continue
                if label == "priority":
                    status = _priority_status(value)
                    if status != "Belirsiz":
                        _set_singleton(result, "priority.status", status, source, seen)
                    current_role, current_person = "priority", None
                    continue
                if label == "priority_country":
                    _set_singleton(result, "priority.country", value, source, seen)
                    current_role, current_person = "priority", None
                    continue
                if label == "priority_number":
                    _set_singleton(result, "priority.number", value, source, seen)
                    current_role, current_person = "priority", None
                    continue
                if label == "priority_date":
                    _set_singleton(result, "priority.date", value, source, seen)
                    current_role, current_person = "priority", None
                    continue

                if label in {"applicant", "inventor"}:
                    current_role = label
                    current_person = None
                    for person_name in _split_people_value(value):
                        row = _new_person(person_name, source, applicant=label == "applicant")
                        target = result["applicants"] if label == "applicant" else result["inventors"]
                        current_person = _append_unique_person(target, row, applicant=label == "applicant")
                    continue

                if label.startswith("applicant_") or label.startswith("inventor_"):
                    role, field = label.split("_", 1)
                    current_role = role
                    target = result["applicants"] if role == "applicant" else result["inventors"]
                    if field == "name":
                        # Önce TCKN/doğum tarihi gibi alanları gelen aynı kişinin boş adını
                        # yeni satır açmadan tamamla. Ancak dolu ve farklı bir ad varsa yeni kişi başlat.
                        if current_person is not None and current_person in target and not current_person.get("name"):
                            current_person["name"] = _clean_value(value)
                            if role == "applicant" and not current_person.get("entity_type"):
                                current_person["entity_type"] = _entity_type(value)
                        else:
                            if current_person is not None and current_person.get("name") and _plain_norm(current_person.get("name", "")) != _plain_norm(value):
                                current_person = None
                            row = _new_person(value, source, applicant=role == "applicant")
                            current_person = _append_unique_person(target, row, applicant=role == "applicant")
                    else:
                        if current_person is None or current_person not in target:
                            current_person = target[-1] if target else ensure_person(role, source)
                        _apply_person_field(current_person, field, value, applicant=role == "applicant")
                    continue

                if label in {"name", "identity", "address", "country", "location", "city", "district", "email", "phone", "birth_date", "entity_type"} and current_role in {"applicant", "inventor"}:
                    target = result["applicants"] if current_role == "applicant" else result["inventors"]
                    if label == "name":
                        if current_person is not None and current_person in target and not current_person.get("name"):
                            current_person["name"] = _clean_value(value)
                            if current_role == "applicant" and not current_person.get("entity_type"):
                                current_person["entity_type"] = _entity_type(value)
                        else:
                            if current_person is not None and current_person.get("name") and _plain_norm(current_person.get("name", "")) != _plain_norm(value):
                                current_person = None
                            row = _new_person(value, source, applicant=current_role == "applicant")
                            current_person = _append_unique_person(target, row, applicant=current_role == "applicant")
                    else:
                        if current_person is None or current_person not in target:
                            current_person = target[-1] if target else ensure_person(current_role, source)
                        _apply_person_field(current_person, label, value, applicant=current_role == "applicant")
                    continue

                # Tanınan fakat çekirdek kişi şemasına dahil olmayan alanları kaybetme.
                if label in {"fax", "postal_code", "authorized_person", "website"}:
                    prefix = "Hak sahibi" if current_role == "applicant" else "Buluş sahibi" if current_role == "inventor" else ""
                    nice = {"fax": "Faks", "postal_code": "Posta kodu", "authorized_person": "Yetkili", "website": "Web sitesi"}[label]
                    _add_other_information(result, f"{prefix} {nice}".strip(), value, source)
                    continue

                # Etiketli fakat çekirdek alanlara ait olmayan bilgileri ön kontrolde koru.
                if value:
                    _add_other_information(result, raw_label, value, source)

        # Serbest e-posta/yazı cümlelerinde yalnız açık rol ifadelerine izin ver.
        applicant_patterns = [
            r"(?:hak|başvuru)\s+sahibi(?:\s+olarak)?\s*(?::|-)?\s+(.{3,180}?)(?=\s+(?:olacaktır|olacak|olarak\s+belirlenmiştir|belirlenmiştir|dır|dir)\b|[\n;]|$)",
            r"başvuru\s+(.{3,180}?)\s+(?:adına|üzerinden)\s+yapılacaktır",
        ]
        inventor_patterns = [
            r"(?:buluş\s+sahibi|buluşçu|mucit|buluşu\s+yapan)(?:\s+olarak)?\s*(?::|-)?\s+(.{3,160}?)(?=\s+(?:olacaktır|olacak|olarak\s+belirlenmiştir|belirlenmiştir|dır|dir)\b|[\n;]|$)",
        ]
        if not result["applicants"]:
            for pattern in applicant_patterns:
                match = re.search(pattern, text, flags=re.IGNORECASE)
                if match:
                    for name in _split_people_value(_clean_value(match.group(1)).strip(" ,")):
                        _append_unique_person(result["applicants"], _new_person(name, source, applicant=True), applicant=True)
                    if result["applicants"]:
                        break
        if not result["inventors"]:
            for pattern in inventor_patterns:
                match = re.search(pattern, text, flags=re.IGNORECASE)
                if match:
                    for name in _split_people_value(_clean_value(match.group(1)).strip(" ,")):
                        _append_unique_person(result["inventors"], _new_person(name, source, applicant=False), applicant=False)
                    if result["inventors"]:
                        break

    # DP referansı tarifname dosya adında zaten bulunuyorsa bunu ana kaynak olarak kullan.
    file_ref = reference_from_filename(specification_filename)
    if file_ref:
        if result["reference"] and _plain_norm(result["reference"]) != _plain_norm(file_ref):
            result["conflicts"].append(
                f"DP / dosya referansı: '{result['reference']}' (başvuru kaynağı) / '{file_ref}' (Tarifname dosya adı)"
            )
        elif not result["reference"]:
            result["reference"] = file_ref
            result["field_sources"]["reference"] = "Tarifname dosya adı"

    # Buluş başlığının tek otoritesi başvuru için yüklenen TARİFNAME'dir.
    # Beyan formu/e-posta içindeki eski veya kısa başlık hiçbir zaman çatışma
    # üretmez ve EPATS başlığını değiştiremez.
    spec_title = _specification_title(specification_text)
    if spec_title:
        source_title = result.get("invention_title", "")
        if source_title and _plain_norm(source_title) != _plain_norm(spec_title):
            _add_other_information(result, "Bilgi kaynağındaki buluş başlığı", source_title, result.get("field_sources", {}).get("invention_title", "Başvuru bilgi kaynağı"))
        result["invention_title"] = spec_title
        result["field_sources"]["invention_title"] = "Tarifname"
        # Daha önce singleton izleme tablosuna bilgi kaynağından yazılmış başlık
        # varsa aşağıdaki field_sources döngüsünün bunu ezmesini engelle.
        seen.pop("invention_title", None)

    # Adreste Türkiye açıkça yazıyorsa veya il alanı 81 Türkiye ilinden biriyse ülkeyi doldur.
    for rows in (result["applicants"], result["inventors"]):
        for row in rows:
            _fill_country_from_explicit_location(row)

    for field_name in ("application_kind", "reference", "invention_title"):
        if field_name in seen and field_name not in result["field_sources"]:
            result["field_sources"][field_name] = seen[field_name][1]
    if result["priority"].get("source"):
        result["field_sources"]["priority"] = result["priority"]["source"]

    # Aynı kimlik numarasının farklı adla kullanılmasını işaretle.
    for role_key, label in (("applicants", "Hak sahibi"), ("inventors", "Buluş sahibi")):
        rows = result[role_key]
        by_id: dict[str, dict[str, str]] = {}
        for row in rows:
            digits = re.sub(r"\D", "", row.get("identity", ""))
            if not digits:
                continue
            if digits in by_id and _plain_norm(by_id[digits].get("name", "")) != _plain_norm(row.get("name", "")):
                result["conflicts"].append(f"{label} kimlik {digits} farklı adlarla bulundu.")
            else:
                by_id[digits] = row

    result["conflicts"] = list(dict.fromkeys(x for x in result["conflicts"] if x))
    return normalize_application_information(result)

def build_application_information_prompt(source_blocks: list[tuple[str, str]], *, specification_text: str = "") -> str:
    blocks = []
    for name, text in source_blocks:
        blocks.append(f"--- KAYNAK: {name} ---\n{text[:45000]}")
    if specification_text.strip():
        blocks.append(f"--- TARİFNAME (yalnız başlık/teyit kaynağı) ---\n{specification_text[:18000]}")
    joined = "\n\n".join(blocks)
    return f"""
Aşağıdaki patent/faydalı model başvuru kaynaklarından yalnız açıkça bulunan bilgileri çıkar.
Tahmin etme, şirket unvanını/adresi/kimlik numarasını uydurma, e-posta gönderen kişiyi otomatik olarak buluş sahibi veya hak sahibi sayma.
Bir bilgi birden fazla kaynakta çelişiyorsa en güvenilir açık ifadeyi seç ve conflicts listesinde çelişkiyi belirt.
Tarifname yalnız buluş başlığını teyit etmek için kullanılabilir; tarifnameden hak sahibi/buluş sahibi uydurma.
Bulabildiğin tüm diğer başvuruya ilişkin bilgileri other_information içinde de koru.

Yalnız aşağıdaki JSON şemasında cevap ver:
{{
  "application_kind": "Patent|Faydalı Model|",
  "reference": "",
  "invention_title": "",
  "applicants": [
    {{"entity_type":"Tüzel kişi|Gerçek kişi|", "identity":"", "name":"", "country":"", "city":"", "address":"", "source":""}}
  ],
  "inventors": [
    {{"identity":"", "name":"", "country":"", "city":"", "address":"", "source":""}}
  ],
  "priority": {{"status":"Var|Yok|Belirsiz", "country":"", "number":"", "date":"", "source":""}},
  "other_information": [{{"label":"", "value":"", "source":""}}],
  "conflicts": [""],
  "source_files_used": [""]
}}

KAYNAKLAR:
{joined}
""".strip()


def normalize_application_information(data: dict[str, Any] | None) -> dict[str, Any]:
    data = data if isinstance(data, dict) else {}
    app_kind = str(data.get("application_kind") or "").strip()
    if app_kind.lower().replace("ı", "i") not in {"patent", "faydali model", "faydalı model"}:
        app_kind = ""
    elif app_kind.lower() == "patent":
        app_kind = "Patent"
    else:
        app_kind = "Faydalı Model"

    def people(key: str, applicant: bool) -> list[dict[str, str]]:
        out: list[dict[str, str]] = []
        raw = data.get(key)
        if not isinstance(raw, list):
            return out
        for item in raw:
            if not isinstance(item, dict):
                continue
            row = {
                "identity": str(item.get("identity") or "").strip(),
                "name": str(item.get("name") or "").strip(),
                "country": str(item.get("country") or "").strip(),
                "city": str(item.get("city") or "").strip(),
                "district": str(item.get("district") or "").strip(),
                "address": str(item.get("address") or "").strip(),
                "email": str(item.get("email") or "").strip(),
                "phone": str(item.get("phone") or "").strip(),
                "birth_date": str(item.get("birth_date") or "").strip(),
                "source": str(item.get("source") or "").strip(),
            }
            if applicant:
                row["entity_type"] = str(item.get("entity_type") or "").strip()
            if any(row.values()):
                out.append(row)
        return out

    priority = data.get("priority") if isinstance(data.get("priority"), dict) else {}
    other = data.get("other_information") if isinstance(data.get("other_information"), list) else []
    return {
        "application_kind": app_kind,
        "reference": str(data.get("reference") or "").strip(),
        "invention_title": str(data.get("invention_title") or "").strip(),
        "applicants": people("applicants", True),
        "inventors": people("inventors", False),
        "priority": {
            "status": str(priority.get("status") or "Belirsiz").strip() or "Belirsiz",
            "country": str(priority.get("country") or "").strip(),
            "number": str(priority.get("number") or "").strip(),
            "date": str(priority.get("date") or "").strip(),
            "source": str(priority.get("source") or "").strip(),
        },
        "other_information": [
            {
                "label": str(x.get("label") or "").strip(),
                "value": str(x.get("value") or "").strip(),
                "source": str(x.get("source") or "").strip(),
            }
            for x in other if isinstance(x, dict) and (x.get("label") or x.get("value"))
        ],
        "conflicts": [str(x).strip() for x in (data.get("conflicts") or []) if str(x).strip()] if isinstance(data.get("conflicts"), list) else [],
        "source_files_used": [str(x).strip() for x in (data.get("source_files_used") or []) if str(x).strip()] if isinstance(data.get("source_files_used"), list) else [],
        "field_sources": {
            str(k): str(v).strip()
            for k, v in (data.get("field_sources") or {}).items()
            if isinstance(data.get("field_sources"), dict) and str(v).strip()
        },
    }


def application_precheck_missing(metadata: dict[str, Any], metrics: dict[str, Any], *, figures_required: bool = False) -> list[str]:
    missing: list[str] = []
    if not str(metadata.get("application_kind") or "").strip():
        missing.append("başvuru türü")
    if not str(metadata.get("invention_title") or "").strip():
        missing.append("buluş başlığı")
    applicants = metadata.get("applicants") or []
    if not applicants:
        missing.append("en az bir hak sahibi / başvuru sahibi")
    else:
        for i, applicant in enumerate(applicants, 1):
            if not str(applicant.get("name") or "").strip():
                missing.append(f"hak sahibi {i} unvan/ad soyad")
            if not str(applicant.get("country") or "").strip():
                missing.append(f"hak sahibi {i} ülke")
            if not str(applicant.get("address") or "").strip():
                missing.append(f"hak sahibi {i} adres")
    inventors = metadata.get("inventors") or []
    if not inventors:
        missing.append("en az bir buluş sahibi")
    else:
        for i, inventor in enumerate(inventors, 1):
            if not str(inventor.get("name") or "").strip():
                missing.append(f"buluş sahibi {i} ad soyad")
            if not str(inventor.get("country") or "").strip():
                missing.append(f"buluş sahibi {i} ülke")
            if not str(inventor.get("address") or "").strip():
                missing.append(f"buluş sahibi {i} adres")
    priority = metadata.get("priority") or {}
    priority_status = str(priority.get("status") or "").strip()
    if priority_status not in {"Var", "Yok"}:
        missing.append("rüçhan durumu (Var/Yok)")
    if priority_status == "Var":
        for field, label in [("country", "rüçhan ülkesi"), ("number", "rüçhan numarası"), ("date", "rüçhan tarihi")]:
            if not str(priority.get(field) or "").strip():
                missing.append(label)
    if int(metrics.get("specification_pages") or 0) <= 0:
        missing.append("Tarifname PDF")
    if int(metrics.get("claim_count") or 0) <= 0:
        missing.append("istemler / istem sayısı")
    if not metrics.get("abstract_present"):
        missing.append("özet")
    if figures_required and int(metrics.get("figures_pages") or 0) <= 0:
        missing.append("şekiller")
    return missing

def split_patent_docx(data: bytes, *, clean_template_colors: bool = True) -> dict[str, bytes]:
    """Tek tarifname DOCX'ini EPATS için Tarifname / İstemler / Özet DOCX bölümlerine ayırır."""
    if clean_template_colors:
        data = strip_template_colored_text(data)
    # LibreOffice bazı Word satır numaralarını sol kenarda "5, 10, 15..."
    # şeklinde EPATS PDF'ine taşır. Başvuru paketinde bunlar istenmediği için
    # yalnız bu otomatik section özelliğini kaldırıyoruz.
    data = remove_word_line_numbering(data)
    with zipfile.ZipFile(io.BytesIO(data), "r") as zin:
        root = ET.fromstring(zin.read("word/document.xml"))
    body = root.find(_W + "body")
    if body is None:
        raise ValueError("Word belgesinin gövdesi okunamadı.")
    content = [c for c in list(body) if c.tag != _W + "sectPr"]
    claims_idx = _find_heading_index(content, "İSTEMLER")
    abstract_idx = _find_heading_index(content, "ÖZET")
    if claims_idx <= 0 or abstract_idx <= claims_idx:
        raise ValueError("İSTEMLER / ÖZET başlık sırası beklenen yapıda değil.")
    return {
        "Tarifname.docx": _slice_docx(data, start_idx=0, end_idx=claims_idx),
        "Istemler.docx": _slice_docx(data, start_idx=claims_idx, end_idx=abstract_idx),
        "Ozet.docx": _slice_docx(data, start_idx=abstract_idx, end_idx=None),
    }


def _libreoffice_convert(data: bytes, filename: str, target_ext: str) -> bytes:
    if shutil.which("libreoffice") is None:
        raise RuntimeError("Belge dönüşümü için LibreOffice kurulu değil.")
    target_ext = target_ext.lstrip(".").lower()
    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        src = td_path / Path(filename).name
        src.write_bytes(data)
        outdir = td_path / "converted"
        outdir.mkdir()
        proc = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", target_ext, "--outdir", str(outdir), str(src)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=120,
        )
        outputs = list(outdir.glob(f"*.{target_ext}"))
        if proc.returncode != 0 or not outputs:
            err = proc.stderr.decode("utf-8", errors="replace").strip()
            raise RuntimeError(f"Belge {target_ext.upper()} biçimine dönüştürülemedi. {err}".strip())
        return outputs[0].read_bytes()


def _libreoffice_to_pdf(data: bytes, filename: str) -> bytes:
    return _libreoffice_convert(data, filename, "pdf")


def _libreoffice_to_docx(data: bytes, filename: str) -> bytes:
    suffix = Path(filename).suffix.lower()
    if suffix == ".docx":
        return data
    return _libreoffice_convert(data, filename, "docx")


def extract_specification_text(filename: str, data: bytes) -> str:
    """Tarifname kaynağından başlık/şekil kontrolünde kullanılacak düz metni çıkarır."""
    suffix = Path(filename).suffix.lower()
    if suffix == ".docx":
        doc = Document(io.BytesIO(data))
        return "\n".join(p.text for p in doc.paragraphs if p.text).strip()
    if suffix == ".doc":
        # DOC'u DOCX'e dönüştürmek antiword'e göre Unicode/tablo biçimini daha iyi korur.
        docx = _libreoffice_to_docx(data, filename)
        return extract_specification_text(Path(filename).with_suffix(".docx").name, docx)
    if suffix == ".pdf":
        return "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(data)).pages).strip()
    raise ValueError("Tarifname dosyası DOC, DOCX veya PDF olmalıdır.")


def _fitz_span_rgb(color: int) -> tuple[int, int, int]:
    return ((int(color) >> 16) & 255, (int(color) >> 8) & 255, int(color) & 255)


def _pdf_epats_cleanup(data: bytes) -> bytes:
    """PDF'de kırmızı/mavi şablon metnini ve sol marj satır numaralarını görsel olarak kaldırır.

    Sayfa boyutu, fontlar ve kalan içerik yeniden dizilmez. PDF kaynaklarında DOCX gibi
    reflow yapılamayacağından kaldırılan alanlar beyaz boşluk olarak kalabilir.
    """
    import fitz  # pymupdf

    doc = fitz.open(stream=data, filetype="pdf")
    changed = False
    for page in doc:
        page_dict = page.get_text("dict")
        for block in page_dict.get("blocks", []):
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = (span.get("text") or "").strip()
                    if not text:
                        continue
                    bbox = fitz.Rect(span.get("bbox"))
                    r, g, b = _fitz_span_rgb(span.get("color", 0))
                    is_template_color = (r >= 150 and g <= 130 and b <= 130) or (b >= 150 and r <= 130 and g <= 160)
                    is_left_line_number = (
                        bool(re.fullmatch(r"\d{1,3}", text))
                        and bbox.x0 < page.rect.width * 0.10
                        and 25 < bbox.y0 < page.rect.height - 25
                        and int(text) % 5 == 0
                    )
                    if is_template_color or is_left_line_number:
                        page.add_redact_annot(bbox + (-1, -1, 1, 1), fill=(1, 1, 1))
                        changed = True
        if changed:
            try:
                page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
            except TypeError:
                page.apply_redactions()
    if not changed:
        out = data
    else:
        out = doc.tobytes(garbage=4, deflate=True)
    doc.close()
    return out


def _find_pdf_heading(data: bytes, heading: str) -> tuple[int, float] | None:
    import fitz

    doc = fitz.open(stream=data, filetype="pdf")
    target = _norm(heading)
    try:
        for page_index, page in enumerate(doc):
            info = page.get_text("dict")
            for block in info.get("blocks", []):
                for line in block.get("lines", []):
                    spans = line.get("spans", [])
                    text = "".join((sp.get("text") or "") for sp in spans).strip()
                    if _norm(text) == target and spans:
                        y0 = min(float(sp.get("bbox", [0, 0, 0, 0])[1]) for sp in spans)
                        return page_index, y0
    finally:
        doc.close()
    return None


def _copy_pdf_vertical_range(src_doc, out_doc, page_index: int, y0: float, y1: float) -> None:
    import fitz

    page = src_doc[page_index]
    y0 = max(0.0, min(float(y0), page.rect.height))
    y1 = max(0.0, min(float(y1), page.rect.height))
    if y1 - y0 < 2:
        return
    new_page = out_doc.new_page(width=page.rect.width, height=page.rect.height)
    clip = fitz.Rect(0, y0, page.rect.width, y1)
    # İçeriği kaynak sayfadaki aynı koordinatlarda tut; yalnız diğer bölümün
    # bulunduğu alanı taşımayarak biçimi değiştirmeden ayır.
    new_page.show_pdf_page(clip, src_doc, page_index, clip=clip)


def split_patent_pdf(data: bytes) -> dict[str, bytes]:
    """Birleşik Tarifname/İstemler/Özet PDF'ini görsel biçimi bozmadan ayırır."""
    import fitz

    cleaned = _pdf_epats_cleanup(data)
    claims_pos = _find_pdf_heading(cleaned, "İSTEMLER")
    abstract_pos = _find_pdf_heading(cleaned, "ÖZET")
    if not claims_pos or not abstract_pos:
        raise ValueError("PDF içinde İSTEMLER ve ÖZET başlıkları bulunamadı.")
    if claims_pos[0] > abstract_pos[0] or (claims_pos[0] == abstract_pos[0] and claims_pos[1] >= abstract_pos[1]):
        raise ValueError("PDF içindeki İSTEMLER / ÖZET sırası beklenen yapıda değil.")

    src = fitz.open(stream=cleaned, filetype="pdf")
    outputs: dict[str, bytes] = {}
    sections = {
        "Tarifname.pdf": ((0, 0.0), claims_pos),
        "Istemler.pdf": (claims_pos, abstract_pos),
        "Ozet.pdf": (abstract_pos, (len(src) - 1, src[-1].rect.height)),
    }
    try:
        for name, (start, end) in sections.items():
            out = fitz.open()
            start_page, start_y = start
            end_page, end_y = end
            for page_index in range(start_page, end_page + 1):
                page = src[page_index]
                y0 = start_y if page_index == start_page else 0.0
                y1 = end_y if page_index == end_page else page.rect.height
                _copy_pdf_vertical_range(src, out, page_index, y0, y1)
            if len(out) == 0:
                raise ValueError(f"PDF içindeki {name} bölümü boş bulundu.")
            outputs[name] = out.tobytes(garbage=4, deflate=True)
            out.close()
    finally:
        src.close()
    return outputs


def count_claims_from_pdf(data: bytes) -> int:
    text = "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(data)).pages)
    explicit = [int(x) for x in re.findall(r"(?mi)^\s*(\d{1,3})\s*[.\-)]+\s+", text)]
    if explicit:
        # 1..N dizisi varsa maksimum istem numarası en güvenilir sayımdır.
        seq = sorted(set(x for x in explicit if 1 <= x <= 999))
        if seq and seq[0] == 1:
            return max(seq)
    deps = re.findall(r"(?mi)^\s*İstem\s+\d+(?:['’`]?e|['’`]?a)\s+uygun", text)
    if deps:
        return 1 + len(deps)
    return 0


def _ensure_pdf(data: bytes, filename: str) -> bytes:
    suffix = Path(filename).suffix.lower()
    if suffix == ".pdf":
        PdfReader(io.BytesIO(data))  # temel bütünlük doğrulaması
        return data
    if suffix in {".docx", ".doc"}:
        return _libreoffice_to_pdf(data, Path(filename).name)
    raise ValueError("Şekiller dosyası DOC, DOCX veya PDF olmalıdır.")


def build_epats_application_package(
    specification_data: bytes,
    *,
    specification_name: str = "Tarifname.docx",
    figures_data: bytes | None = None,
    figures_name: str | None = None,
    metadata: dict[str, Any] | None = None,
) -> tuple[bytes, dict[str, bytes]]:
    """EPATS'a yüklenmeye hazır PDF'leri üretip ZIP paket olarak döndürür.

    Tarifname kaynağı DOC/DOCX/PDF olabilir. DOC önce DOCX'e dönüştürülür;
    PDF ise görsel biçimi korunarak İSTEMLER/ÖZET sınırlarından ayrılır.
    """
    suffix = Path(specification_name).suffix.lower()
    pdfs: dict[str, bytes] = {}
    if suffix in {".docx", ".doc"}:
        docx = specification_data if suffix == ".docx" else _libreoffice_to_docx(specification_data, specification_name)
        split_docs = split_patent_docx(docx)
        for docx_name, docx_data in split_docs.items():
            pdf_name = Path(docx_name).with_suffix(".pdf").name
            pdfs[pdf_name] = _libreoffice_to_pdf(docx_data, docx_name)
    elif suffix == ".pdf":
        pdfs.update(split_patent_pdf(specification_data))
    else:
        raise ValueError("Tarifname dosyası DOC, DOCX veya PDF olmalıdır.")

    if figures_data is not None:
        safe_name = figures_name or "Sekiller.docx"
        pdfs["Sekiller.pdf"] = _ensure_pdf(figures_data, safe_name)

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, payload in pdfs.items():
            zf.writestr(name, payload)
        if metadata:
            zf.writestr("basvuru_verileri.json", json.dumps(metadata, ensure_ascii=False, indent=2).encode("utf-8"))
    return out.getvalue(), pdfs
