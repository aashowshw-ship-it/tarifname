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
from pathlib import Path
from typing import Any

from docx import Document
from pypdf import PdfReader

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
    dependent = re.findall(r"(?mi)^\s*İstem\s+\d+['’`]?e\s+uygun", text)
    if dependent:
        return 1 + len(dependent)
    return 0


def pdf_page_count(data: bytes) -> int:
    return len(PdfReader(io.BytesIO(data)).pages)


def epats_document_metrics(specification_docx: bytes, pdfs: dict[str, bytes]) -> dict[str, Any]:
    cleaned = strip_template_colored_text(specification_docx)
    split_docs = split_patent_docx(cleaned, clean_template_colors=False)
    claims = count_claims_from_docx(split_docs["Istemler.docx"])
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


def _html_to_text(value: str) -> str:
    value = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", value or "")
    value = re.sub(r"(?i)<br\s*/?>", "\n", value)
    value = re.sub(r"(?i)</p\s*>", "\n", value)
    value = re.sub(r"(?s)<[^>]+>", " ", value)
    return re.sub(r"[ \t]+", " ", unescape(value)).strip()


def _eml_text(data: bytes) -> str:
    msg = BytesParser(policy=policy.default).parsebytes(data)
    head = [
        f"Konu: {msg.get('subject', '')}",
        f"Kimden: {msg.get('from', '')}",
        f"Kime: {msg.get('to', '')}",
        f"Cc: {msg.get('cc', '')}",
        f"Tarih: {msg.get('date', '')}",
    ]
    body_parts: list[str] = []
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_disposition() == "attachment":
                continue
            ctype = part.get_content_type()
            if ctype == "text/plain":
                try:
                    body_parts.append(part.get_content())
                except Exception:
                    payload = part.get_payload(decode=True) or b""
                    body_parts.append(payload.decode(part.get_content_charset() or "utf-8", errors="replace"))
            elif ctype == "text/html" and not body_parts:
                try:
                    body_parts.append(_html_to_text(part.get_content()))
                except Exception:
                    pass
    else:
        try:
            content = msg.get_content()
            body_parts.append(_html_to_text(content) if msg.get_content_type() == "text/html" else str(content))
        except Exception:
            pass
    return "\n".join([x for x in head + body_parts if x and x.strip()]).strip()


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
        text = "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(data)).pages)
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
            text = "\n".join(
                x for x in [
                    f"Konu: {getattr(msg, 'subject', '') or ''}",
                    f"Kimden: {getattr(msg, 'sender', '') or ''}",
                    f"Kime: {getattr(msg, 'to', '') or ''}",
                    f"Cc: {getattr(msg, 'cc', '') or ''}",
                    f"Tarih: {getattr(msg, 'date', '') or ''}",
                    getattr(msg, "body", "") or "",
                ] if x.strip()
            )
            try:
                msg.close()
            except Exception:
                pass
    else:
        raise ValueError("Bilgi kaynağı için desteklenen türler: .docx, .doc, .pdf, .txt, .md, .eml, .msg")
    text = (text or "").replace("\x00", " ").strip()
    if not text:
        raise ValueError(f"{filename} dosyasından metin çıkarılamadı.")
    return text


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
                "address": str(item.get("address") or "").strip(),
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


def _libreoffice_to_pdf(data: bytes, filename: str) -> bytes:
    if shutil.which("libreoffice") is None:
        raise RuntimeError("PDF üretimi için LibreOffice kurulu değil.")
    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        src = td_path / filename
        src.write_bytes(data)
        outdir = td_path / "pdf"
        outdir.mkdir()
        proc = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(src)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=120,
        )
        pdfs = list(outdir.glob("*.pdf"))
        if proc.returncode != 0 or not pdfs:
            err = proc.stderr.decode("utf-8", errors="replace").strip()
            raise RuntimeError(f"PDF üretilemedi. {err}".strip())
        return pdfs[0].read_bytes()


def _ensure_pdf(data: bytes, filename: str) -> bytes:
    suffix = Path(filename).suffix.lower()
    if suffix == ".pdf":
        PdfReader(io.BytesIO(data))  # temel bütünlük doğrulaması
        return data
    if suffix == ".docx":
        return _libreoffice_to_pdf(data, Path(filename).name)
    raise ValueError("Şekiller dosyası PDF veya DOCX olmalıdır.")


def build_epats_application_package(
    specification_docx: bytes,
    *,
    figures_data: bytes | None = None,
    figures_name: str | None = None,
    metadata: dict[str, Any] | None = None,
) -> tuple[bytes, dict[str, bytes]]:
    """EPATS'a yüklenmeye hazır PDF'leri üretip ZIP paket olarak döndürür."""
    split_docs = split_patent_docx(specification_docx)
    pdfs: dict[str, bytes] = {}
    for docx_name, docx_data in split_docs.items():
        pdf_name = Path(docx_name).with_suffix(".pdf").name
        pdfs[pdf_name] = _libreoffice_to_pdf(docx_data, docx_name)

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
