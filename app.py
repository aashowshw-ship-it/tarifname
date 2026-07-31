from __future__ import annotations

import base64
import io
import json
import os
import re
import subprocess
import tempfile
import zipfile
from copy import deepcopy
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable

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
from pypdf import PdfReader

BASE_DIR = Path(__file__).resolve().parent
TARIFNAME_TEMPLATE = BASE_DIR / "Tarifname_181176_template.docx"
GORUS_TEMPLATE = BASE_DIR / "Gorus_metni_696809_template.docx"
ARASTIRMA_TEMPLATE = BASE_DIR / "On_Arastirma_Raporu_181612_template.docx"
MODEL = os.getenv("OPENAI_MODEL", "gpt-5.6")
MAX_TEXT_PER_FILE = int(os.getenv("MAX_TEXT_PER_FILE", "70000"))
MAX_TOTAL_TEXT = int(os.getenv("MAX_TOTAL_TEXT", "260000"))


# -----------------------------------------------------------------------------
# GENEL KURALLAR
# -----------------------------------------------------------------------------
TARIFNAME_RULES = r"""
TÜRK PATENT TARİFNAME KURALLARI
1. Yalnızca yüklenen BBF'deki bilgi ve ifadeleri kullan. BBF'de bulunmayan teknik unsur, algoritma, değer, bağlantı, özellik veya kullanım biçimi ekleme.
2. İngilizce teknik terimi ilk geçtiği yerde Türkçesini önce, İngilizcesini parantez içinde ver. Örnek: frekans aralığı 2 (frequency range 2, FR2). Sonraki kullanımlarda yalnızca Türkçe karşılığını kullan. AI yerine yapay zekâ yaz.
3. Unsur adlarını normal cümle düzeninde yaz. Her unsur ve yöntem adımının ilk harfi büyük olsun.
4. BBF'deki unsur numaralandırmasını aynen koru. 1,2,3 ise aynı; 10,20,30 ise aynı. BBF'de unsur olarak verilmeyen sisteme/yönteme ek numara verme.
5. REFERANS NUMARALARI bölümünden önce unsur veya işlem adımı referansı kullanma.
6. Referans listesinde önce sistem unsurları, bir boş paragraf sonra yöntem işlem adımları yer alır.
7. Yöntem işlem adımları REFERANS NUMARALARI bölümünde '1001. Entegre ... toplanması' biçiminde yer alır ve burada 10,20,30 gibi modül referansları kullanılmaz.
8. Yöntem isteminde kullanılan 1001,1002... işlem metinleri ile REFERANS NUMARALARI bölümündeki aynı numaralı işlem metinleri birebir aynı olmalıdır. Yöntem isteminde adım sonunda '(1001)' biçiminde gösterilir.
9. Detaylı açıklamada yöntem işlem adımı referansı yalnızca işlem ifadesinin sonunda '(1001)' biçiminde kullanılabilir.
10. Sistem istemindeki unsurları BBF sırasıyla tanımla. Bir unsur tanımlanırken yalnızca daha önce tanımlanmış unsurlarla teknik ilişki kur. Henüz tanımlanmamış unsuru önceki unsurun içinde kullanma.
11. Ana sistem istemindeki unsurlar veri, kontrol, sinyal veya işlem ilişkisi içinde kurulmalıdır.
12. Alt istemler kısa olmalı, ana istemi tekrar etmemeli ve BBF'deki ek teknik özelliklere dayanmalıdır. Gereksiz alt istem yazma. Sistem alt istemlerini 'bir modül olmasıdır' veya 'içermesidir' şeklinde bitir; 'yapmasıdır/etmesidir/belirlemesidir' kullanma.
13. Açık ve sıralı bilgisayar tarafından gerçekleştirilen işlem akışı varsa yöntem istemi oluştur; yoksa yalnızca sistem istemi oluştur.
14. Şablondaki kırmızı/mavi açıklama metinlerini ve biçimlerini koru.
15. İnsan veya operatör eylemlerini teknik araç üzerinden yaz: 'elektronik cihaz üzerinden kullanıcıya sunan', 'elektronik cihaz üzerinden operatöre ileten' gibi.
16. Detaylı açıklamadaki modül açıklamalarını gereksiz biçimde ayrı paragraflara bölme; teknik akış elverdiği ölçüde peş peşe tek paragrafta açıkla.
17. 'Yöntemin gerçekleştirdiği işlevler aşağıdaki gibidir:' ifadesinden sonra yöntem adımlarını küçük yuvarlak veya tire ile alt alta yaz.
18. Buluşun çalışma prensibini gereksiz biçimde farklı paragraflara bölme.
19. İSTEMLER ve ÖZET başlıklarını ortala.
20. Sistem ve yöntem istemlerini oluşturduktan sonra ikinci bir istem kalite kontrolü yap: kapsam, teknik taşıyıcı, unsur sırası, tekrar, dayanak ve dil bakımından hataları düzelt.
"""

GORUS_RULES = r"""
TÜRK PATENT GÖRÜŞ ÇALIŞMASI KURALLARI
1. Yalnızca raporda X veya Y olarak gösterilen dokümanlara karşı savunma yap. A kategorisi veya itiraz dayanağı yapılmayan dokümanlara karşı görüş yazma.
2. Araştırma raporuna karşı görüşte rapor, tarifname, D1/D2 ve varsa müşteri bilgilerini birlikte analiz et.
3. İnceleme raporuna karşı görüşte bunlara ek olarak önceki görüşü analiz et; uzmanın ikna olmadığı savunmaları aynen tekrarlamak yerine farklı teknik ayrım ve dayanaklar geliştir.
4. Müşteri bilgisini yalnızca tarifname/istemlerde açık dayanağı varsa doğrudan kullan. Dayanağı yoksa teknik gerçek gibi yazma; uygunsa çıkarım olduğunu belirterek yumuşat veya kullanma.
5. Teknik farklara, teknik etkiye ve unsurlar arasındaki işlevsel ilişkiye odaklan.
6. Tarifname dayanağı verilecek yerde şu kalıbı kullan: 'Tarifnamede bu durum şu şekilde belirtilmektedir:' Ardından tarifnamedeki ilgili cümle/pasajı tırnak içinde ve kalın ver.
7. Tarifname alıntısını kesme, değiştirme, sadeleştirme veya kelime ekleyip çıkarma. Alıntı tarifname metninde birebir bulunmalıdır.
8. Yenilik itirazında ilgili istemin tüm özelliklerinin tek dokümanda doğrudan ve açık biçimde açıklanmadığını göster.
9. Buluş basamağı itirazında D1 ve D2'yi tek başına ve birlikte değerlendir; teknik fark, teknik etki, objektif teknik problem, birleştirme motivasyonu ve geriye dönük değerlendirme riskini ele al.
10. Başvuru numarası ve başvuru sahibi rapordan çekilsin. Referans kullanıcıdan alınsın.
11. Görüş formatı yüklenen Görüş metni_696809 örneğine sadık kalsın.
12. Çıktı oluşturulduktan sonra ikinci bir kalite kontrolü yap: yanlış doküman, dayanağı olmayan müşteri bilgisi, eksik alıntı, tekrar eden savunma ve sonuç tutarlılığı bakımından düzelt.
"""

ARASTIRMA_RULES = r"""
TİP 3 ÖN ARAŞTIRMA RAPORU KURALLARI
1. BBF'deki teknik problem, unsurlar, işlevler, işlem adımları ve teknik etkiler üzerinden global patent araştırması yap.
2. En benzer 10 patent dokümanını yayın/başvuru numarası, başlık, tarih, ülke/otorite ve doğrulanabilir kaynak bağlantısıyla belirle. Doküman uydurma.
3. Tek bir doküman araştırma konusu buluşun bütün esas teknik özelliklerini ve aralarındaki ilişkiyi doğrudan ve açık biçimde açıklıyorsa bu dokümanı D1 seç ve yenilik kriterinin sağlanmadığı sonucuna göre rapor hazırla. Bu durumda D2 zorunlu değildir.
4. Yeniliği tek başına bozan doküman yoksa en yakın D1 ve tamamlayıcı D2'yi seç; yenilik değerlendirmesini ayrı ayrı, buluş basamağını D1 ve D2 birlikte düşünülerek yap.
5. Kullanıcının yüklediği benzer dokümanları da incele; ilk seçilen D1/D2'nin yerini alabilecek daha yakın veya daha güçlü doküman varsa nihai seçimi değiştir.
6. Yardımcı dokümanlar buluş basamağı değerlendirmesinde yalnızca destekleyici olabilir; nihai D1/D2 açıkça belirtilmelidir.
7. Rapor, Ön Araştırma Raporu_181612 formatına sadık kalsın: kapak, kriterler, anahtar kelimeler, IPC/CPC, değerlendirme, D1/D2 tanıtımı, özet/şekil alanı, karşılaştırma tabloları, yenilik, buluş basamağı, sonuç ve gerekiyorsa uyarılar.
8. Sonuç açık olsun: yenilik sağlanır/sağlanmaz; buluş basamağı sağlanır/sağlanmaz.
9. Teknik değerlendirmeyi BBF'ye sadık yap; tanıtım veya salt iş kuralı niteliğindeki yönleri teknik katkı gibi abartma.
10. Rapor metnini oluşturduktan sonra ikinci kalite kontrolü yap: D1/D2 seçimi, özellik eşleştirmesi, yenilik mantığı, birleştirme motivasyonu ve sonuç tutarlılığı bakımından düzelt.
"""


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
        run.bold = bold or run.bold
        run.italic = italic or run.italic
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
    name = (name or default).strip()
    if not name.lower().endswith(".docx"):
        name += ".docx"
    return Path(name).name


# -----------------------------------------------------------------------------
# TARİFNAME MODÜLÜ
# -----------------------------------------------------------------------------
def tarifname_extraction_prompt(source_text: str) -> str:
    return f"""{TARIFNAME_RULES}
Aşağıdaki BBF'yi yapılandırılmış veri olarak çıkar. Bilgileri düzeltme veya genişletme. JSON dışında yazma.
ŞEMA:
{{
 "title":"", "technical_field":"", "prior_art":[""], "problems":[""], "advantages":[""],
 "elements":[{{"number":"10","name":"","function":""}}],
 "method_steps":[{{"number":"1001","text":""}}], "working_principle":[""], "keywords":[""],
 "figures":[""], "has_method_basis":true, "method_basis_reason":""
}}
BBF:\n---\n{source_text}\n---"""


def tarifname_literature_prompt(extracted: dict[str, Any], count: int, jurisdiction: str) -> str:
    return f"""Aşağıdaki buluş için tam olarak {count} teknik olarak yakın patent dokümanı araştır. Web araması kullan.
Doküman uydurma; yayın/başvuru numarasını, başlığı ve kaynak bağlantısını doğrula. Tercih: {jurisdiction or 'global'}.
Her dokümanın yakın yönünü ve buluşta bulunup dokümanda bulunmayan temel teknik farkı yaz.
JSON dışında yazma.
ŞEMA: {{"documents":[{{"application_number":"","title":"","jurisdiction":"","summary":"","difference":"","source_url":""}}]}}
BULUŞ: {json.dumps(extracted, ensure_ascii=False)}"""


def tarifname_drafting_prompt(extracted: dict[str, Any], claim_mode: str, literature: list[dict[str, Any]]) -> str:
    return f"""{TARIFNAME_RULES}
Aşağıdaki BBF verilerinden Türk patent tarifnamesi oluştur. İstem yapısı: {claim_mode}.
Literatür dokümanlarını yalnızca ÖNCEKİ TEKNİK bölümünde kullan. BBF dışında yeni teknik özellik ekleme.
JSON dışında yazma.
ŞEMA:
{{
 "title":"", "technical_field":"", "prior_art_general_paragraphs":[""], "literature_paragraphs":[""],
 "short_description_intro":"", "objectives":[""], "unumbered_system_definition":"", "unumbered_system_elements":[""],
 "figure_descriptions":[""], "elements":[{{"number":"10","name":"","description":""}}],
 "method_steps":[{{"number":"1001","text":""}}], "detailed_paragraphs":[""],
 "method_functions":[""], "working_principle":"",
 "system_claim":{{"preamble":"","elements":[""],"closing":"içermesidir."}},
 "dependent_system_claims":[""],
 "method_claim":{{"preamble":"","steps":[""],"closing":"işlem adımlarını içermesidir."}},
 "dependent_method_claims":[""], "abstract":""
}}
ÖZEL:
- method_steps.text alanında 10,20 gibi sistem referans numarası kullanma.
- method_claim.steps içindeki her adım method_steps.text ile birebir aynı olsun ve sonuna ilgili (1001) eklensin.
- Alt istemleri yalnızca ana istemden farklı ve BBF'de dayanaklı ek özellikler için yaz.
- İstemleri ikinci kez kontrol edip düzelt.
BBF: {json.dumps(extracted, ensure_ascii=False, indent=2)}
LİTERATÜR: {json.dumps(literature, ensure_ascii=False, indent=2)}"""


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
    add_text(doc, draft["title"], bold=True, center=True)
    doc.add_paragraph()
    copy_template_paragraph(doc, template, 4)
    doc.add_paragraph()

    add_heading(doc, "TEKNİK ALAN")
    add_text(doc, draft.get("technical_field", ""))
    add_heading(doc, "ÖNCEKİ TEKNİK")
    for p in draft.get("prior_art_general_paragraphs") or []:
        add_text(doc, p)
        doc.add_paragraph()
    for i, p in enumerate(draft.get("literature_paragraphs") or []):
        add_text(doc, p)
        if i < len(draft.get("literature_paragraphs") or []) - 1:
            doc.add_paragraph()
    add_text(doc, "Sonuçta yukarıda bahsedilen ve mevcut teknik ışığında çözülemeyen sorunlar, ilgili teknik alanda bir yenilik yapmayı zorunlu kılmıştır.")

    add_heading(doc, "BULUŞUN KISA AÇIKLAMASI")
    add_text(doc, draft.get("short_description_intro", ""))
    for i, objective in enumerate(draft.get("objectives") or []):
        prefix = "Buluşun ana amacı, " if i == 0 else "Buluşun diğer bir amacı, "
        objective = str(objective).strip()
        add_text(doc, prefix + (objective[:1].lower() + objective[1:] if objective else ""))
    if draft.get("unumbered_system_definition"):
        add_text(doc, draft["unumbered_system_definition"])
    for item in draft.get("unumbered_system_elements") or []:
        add_bullet(doc, item)
    if draft.get("unumbered_system_elements"):
        add_text(doc, "içermesidir.")

    add_heading(doc, "ŞEKİLLERİN KISA AÇIKLAMASI")
    for figure in draft.get("figure_descriptions") or ["Şekil 1, buluşa konu sistemin temsili gösterimidir."]:
        add_text(doc, figure)

    add_heading(doc, "REFERANS NUMARALARI")
    for element in draft.get("elements") or []:
        add_text(doc, f"{element['number']}. {element['name']}")
    if draft.get("method_steps"):
        doc.add_paragraph()
    for step in draft.get("method_steps") or []:
        text = re.sub(r"\s*\(\s*\d+\s*\)\s*", "", str(step.get("text", ""))).strip()
        add_text(doc, f"{step['number']}. {text}")

    add_heading(doc, "BULUŞUN DETAYLI AÇIKLAMASI")
    add_text(doc, f"Bu detaylı açıklamada, buluş konusu olan {draft['title'].lower()} sadece konunun daha iyi anlaşılmasına yönelik hiçbir sınırlayıcı etki oluşturmayacak örneklerle açıklanmaktadır.")
    for p in draft.get("detailed_paragraphs") or []:
        add_text(doc, p)
    if draft.get("method_functions"):
        add_text(doc, "Yöntemin gerçekleştirdiği işlevler aşağıdaki gibidir:")
        for item in draft["method_functions"]:
            add_bullet(doc, item, symbol="-")
    if draft.get("working_principle"):
        add_text(doc, draft["working_principle"])

    add_heading(doc, "İSTEMLER", center=True)
    for index in (79, 81, 83):
        copy_template_paragraph(doc, template, index)
    claim_no = 1
    sc = draft.get("system_claim") or {}
    add_text(doc, f"{claim_no}. {sc.get('preamble','')} olup, özelliği;")
    for item in sc.get("elements") or []:
        add_bullet(doc, item, symbol="-")
    add_text(doc, sc.get("closing", "içermesidir."))
    claim_no += 1
    for dep in draft.get("dependent_system_claims") or []:
        add_text(doc, f"{claim_no}. {dep}")
        claim_no += 1
    mc = draft.get("method_claim")
    if mc:
        add_text(doc, f"{claim_no}. {mc.get('preamble','')} olup, özelliği;")
        for item in mc.get("steps") or []:
            add_bullet(doc, item, symbol="-")
        add_text(doc, mc.get("closing", "işlem adımlarını içermesidir."))
        claim_no += 1
        for dep in draft.get("dependent_method_claims") or []:
            add_text(doc, f"{claim_no}. {dep}")
            claim_no += 1

    add_heading(doc, "ÖZET", center=True)
    add_text(doc, draft["title"], bold=True, center=True)
    add_text(doc, draft.get("abstract", ""))
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# -----------------------------------------------------------------------------
# GÖRÜŞ MODÜLÜ
# -----------------------------------------------------------------------------
def gorus_prompt(
    report_type: str,
    reference: str,
    report_text: str,
    spec_text: str,
    prior_opinion_text: str,
    similar_text: str,
    customer_text: str,
) -> str:
    return f"""{GORUS_RULES}
Aşağıdaki dosyalara dayanarak Türk Patent ve Marka Kurumu için ayrıntılı görüş metni hazırla.
Görüş türü: {report_type}
Ana dosya referansı: {reference}
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
- Metni ikinci kez kalite kontrolünden geçir.

RAPOR:\n{report_text}\n
TARİFNAME:\n{spec_text}\n
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
    doc = Document(str(GORUS_TEMPLATE))
    clear_body(doc)
    for section in doc.sections:
        section.top_margin = Cm(2.2)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(2.0)

    add_text(doc, "TÜRK PATENT VE MARKA KURUMU", bold=True, center=True)
    add_text(doc, "Patent Dairesi Başkanlığına", bold=True, center=True)
    doc.add_paragraph()
    table = doc.add_table(rows=3, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    labels = ["Başvuru No", "Başvuru Sahibi", "Referans"]
    values = [opinion.get("application_no", ""), opinion.get("applicant", ""), opinion.get("reference", "")]
    for row, label, value in zip(table.rows, labels, values):
        set_cell_text(row.cells[0], label, bold=True)
        set_cell_text(row.cells[1], f":   {value}")
    doc.add_paragraph()
    add_text(doc, "Sayın Uzman,")
    add_text(doc, opinion.get("intro", ""))

    docs = opinion.get("cited_documents") or []
    if docs:
        for d in docs:
            add_text(doc, f"{d.get('label','')}: {d.get('number','')} {d.get('title','')}")
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
        add_text(doc, line)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# -----------------------------------------------------------------------------
# TİP 3 ÖN ARAŞTIRMA MODÜLÜ
# -----------------------------------------------------------------------------
def top10_research_prompt(bbf_text: str) -> str:
    return f"""{ARASTIRMA_RULES}
Aşağıdaki BBF için global patent araştırması yap ve en benzer tam 10 patent dokümanını belirle.
Google Patents, Espacenet, PATENTSCOPE, TÜRKPATENT ve ulaşılabilir resmi/yarı resmi patent kaynaklarını kapsayacak geniş web araştırması yap.
Dokümanları teknik yakınlığa göre sırala. Numara, başlık, tarih ve kaynak URL doğrulanmış olsun. JSON dışında yazma.
ŞEMA:
{{
 "subject_title":"", "technical_features":[""], "keywords":[""], "ipc_cpc":[""],
 "documents":[{{
   "rank":1,"publication_number":"","application_number":"","title":"","date":"","jurisdiction":"","source_url":"",
   "summary":"","matching_features":[""],"missing_features":[""],"novelty_destroying":false,"novelty_reason":"","relevance_score":0
 }}],
 "proposed_d1":"publication_number", "proposed_d2":"publication_number veya boş",
 "preliminary_novelty":"sağlanır/sağlanmaz", "preliminary_inventive_step":"sağlanır/sağlanmaz/belirsiz"
}}
BBF:\n{bbf_text}"""


def final_selection_prompt(bbf_text: str, top10: dict[str, Any], user_docs_text: str) -> str:
    return f"""{ARASTIRMA_RULES}
Aşağıdaki BBF, sistemin bulduğu 10 doküman ve kullanıcının varsa yüklediği dokümanları birlikte incele.
Nihai D1 ve gerekiyorsa D2'yi seç. Kullanıcı dokümanı daha yakınsa D1/D2'yi değiştir.
Tek doküman bütün esas teknik özellikleri doğrudan ve açık açıklıyorsa D1 ile yenilik sağlanmaz sonucuna git ve D2 seçme.
Aksi halde D1 ve tamamlayıcı D2 ile buluş basamağını değerlendir.
JSON dışında yazma.
ŞEMA:
{{
 "d1":{{"number":"","title":"","date":"","source":"system/user","summary":"","abstract":"","figure_description":""}},
 "d2":null,
 "novelty_result":"sağlanır/sağlanmaz", "inventive_step_result":"sağlanır/sağlanmaz",
 "novelty_reasoning":[""], "inventive_step_reasoning":[""],
 "comparison_rows_d1":[{{"feature":"","status":"+/-","evidence":""}}],
 "comparison_rows_d2":[{{"feature":"","status":"+/-","evidence":""}}],
 "helper_documents":[{{"number":"","title":"","role":""}}],
 "warnings":[""]
}}
BBF:\n{bbf_text}\n
TOP10:\n{json.dumps(top10, ensure_ascii=False, indent=2)}\n
KULLANICI DOKÜMANLARI:\n{user_docs_text}"""


def report_drafting_prompt(bbf_text: str, top10: dict[str, Any], selection: dict[str, Any], reference: str) -> str:
    return f"""{ARASTIRMA_RULES}
Aşağıdaki verilere göre Ön Araştırma Raporu_181612 biçiminde ayrıntılı rapor içeriği oluştur.
DP referans numarası: {reference}
JSON dışında yazma.
ŞEMA:
{{
 "reference":"{reference}", "title":"", "report_date":"{date.today().strftime('%d.%m.%Y')}",
 "purpose":"Belirlenen konuda araştırmanın gerçekleştirilmesi", "scope":"Global (İlan edilmiş olan patent başvuruları)",
 "keywords":[""], "ipc_cpc":[{{"code":"","description":""}}],
 "evaluation_intro":"", "novelty_heading":"2.1. Yenilik Değerlendirmesi",
 "documents":[{{
   "label":"D1","number":"","alternate_number":"","title":"","date":"","description":[""],"abstract":"",
   "figure_caption":"D1- Şekil", "comparison_rows":[{{"feature":"","status_evidence":""}}], "novelty_assessment":[""]
 }}],
 "inventive_step_paragraphs":[""], "conclusion_paragraphs":[""], "warnings":[""],
 "attachments":["Benzer Dokümanlar","Ön İnceleme Raporu","Makine Tercümeleri"]
}}
ÖZEL:
- selection.d2 null ise yalnızca D1 bölümü oluştur.
- Sonuçta yenilik ve buluş basamağı sonucunu açıkça yaz.
- Yardımcı dokümanları yalnızca buluş basamağına destek olarak kullan.
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
            set_cell_text(cells[1], row.get("status_evidence", ""))
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
st.set_page_config(page_title="Patent Atölyesi", page_icon="⚙️", layout="wide")
st.markdown(
    """
    <style>
      .block-container {max-width: 1180px; padding-top: 1.5rem; padding-bottom: 3rem;}
      .hero {padding: 1.2rem 1.4rem; border:1px solid #e7e7e7; border-radius:16px; margin-bottom:1rem;}
      .hero h1 {margin:0; font-size:2rem;}
      .hero p {margin:.35rem 0 0 0; color:#666;}
      div[data-testid="stDownloadButton"] button, div[data-testid="stFormSubmitButton"] button {width:100%;}
    </style>
    <div class="hero"><h1>Patent Atölyesi</h1><p>Tarifname, görüş ve Tip 3 ön araştırma çalışmalarını tek arayüzden oluşturun.</p></div>
    """,
    unsafe_allow_html=True,
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
    st.subheader("Tarifname oluşturma")
    with st.form("tarifname_form"):
        bbf = st.file_uploader("BBF dosyası", type=["docx", "doc", "pdf", "txt"], key="tar_bbf")
        output_name = st.text_input("Çıktı dosyasının adı", value="Tarifname_XXXXXX.docx")
        claim_choice = st.selectbox("İstem yapısı", ["BBF'ye göre otomatik belirle", "Yalnızca sistem", "Sistem ve yöntem"])
        literature = st.checkbox("Literatür araştırması yap ve önceki tekniğe ekle")
        c1, c2 = st.columns(2)
        with c1:
            lit_count = st.number_input("Benzer patent sayısı", min_value=1, max_value=10, value=2, disabled=not literature)
        with c2:
            jurisdiction = st.text_input("Tercih edilen ülke/veri tabanı", disabled=not literature)
        submit = st.form_submit_button("Tarifnameyi oluştur", type="primary")
    if submit:
        if bbf is None:
            st.error("BBF yükleyin.")
        else:
            try:
                progress = st.progress(0, text="BBF okunuyor...")
                source = extract_text_from_asset(UploadedAsset(bbf.name, bbf.getvalue(), bbf.type))
                extracted = ask_json(tarifname_extraction_prompt(source))
                progress.progress(25, text="İstem yapısı belirleniyor...")
                mode = claim_choice
                if mode == "BBF'ye göre otomatik belirle":
                    mode = "Sistem ve yöntem" if extracted.get("has_method_basis") else "Yalnızca sistem"
                lit_docs: list[dict[str, Any]] = []
                if literature:
                    progress.progress(40, text="Patent literatürü araştırılıyor...")
                    lit_docs = (ask_json(tarifname_literature_prompt(extracted, int(lit_count), jurisdiction), web_search=True).get("documents") or [])
                progress.progress(65, text="Tarifname ve istemler hazırlanıyor...")
                draft = ask_json(tarifname_drafting_prompt(extracted, mode, lit_docs))
                if mode == "Yalnızca sistem":
                    draft["method_claim"] = None
                    draft["dependent_method_claims"] = []
                    draft["method_steps"] = []
                progress.progress(90, text="Word dosyası hazırlanıyor...")
                data = build_tarifname_docx(draft)
                progress.progress(100, text="Hazır")
                st.success("Tarifname oluşturuldu.")
                st.download_button("Word dosyasını indir", data=data, file_name=safe_output_name(output_name, "Tarifname.docx"), mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")
            except Exception as exc:
                st.exception(exc)

# GÖRÜŞ
elif work_type == "Görüş hazırlama":
    st.subheader("Görüş hazırlama")
    report_type = st.selectbox("Görüş türü", ["Araştırma raporuna karşı görüş", "İnceleme raporuna karşı görüş"])
    reference = st.text_input("Ana dosya referansı nedir?", value="")
    output_name = st.text_input("Çıktı dosyasının adı", value="Görüş metni_XXXXXX.docx")
    report_file = st.file_uploader("Araştırma / inceleme raporu", type=["pdf", "docx", "doc", "txt"], key="gor_report")
    spec_file = st.file_uploader("Tarifname", type=["pdf", "docx", "doc", "txt"], key="gor_spec")
    similar_files = st.file_uploader("Rapordaki X/Y benzer dokümanlar (D1, D2 vb.)", type=["pdf", "docx", "doc", "txt", "zip"], accept_multiple_files=True, key="gor_sim")
    customer_yes = st.radio("Müşteriden bilgi var mı?", ["Hayır", "Evet"], horizontal=True)
    customer_files = []
    if customer_yes == "Evet":
        customer_files = st.file_uploader("Müşteri bilgileri", type=["pdf", "docx", "doc", "txt", "png", "jpg", "jpeg", "webp", "zip"], accept_multiple_files=True, key="gor_customer")
    prior_file = None
    if report_type == "İnceleme raporuna karşı görüş":
        prior_file = st.file_uploader("Önceki sunulan görüş", type=["pdf", "docx", "doc", "txt"], key="gor_prior")
    if st.button("Görüş metnini oluştur", type="primary", use_container_width=True):
        if not all([reference.strip(), report_file, spec_file]) or not similar_files:
            st.error("Referans, rapor, tarifname ve X/Y dokümanlarını yükleyin.")
        elif report_type == "İnceleme raporuna karşı görüş" and prior_file is None:
            st.error("İnceleme raporu için önceki sunulan görüşü yükleyin.")
        else:
            try:
                progress = st.progress(0, text="Dosyalar okunuyor...")
                report_text = extract_text_from_asset(UploadedAsset(report_file.name, report_file.getvalue(), report_file.type))
                spec_text = extract_text_from_asset(UploadedAsset(spec_file.name, spec_file.getvalue(), spec_file.type))
                prior_text = ""
                if prior_file:
                    prior_text = extract_text_from_asset(UploadedAsset(prior_file.name, prior_file.getvalue(), prior_file.type))
                sim_assets = assets_from_uploads(similar_files)
                sim_text, sim_images = combine_asset_text("BENZER DOKÜMAN", sim_assets)
                cust_assets = assets_from_uploads(customer_files)
                cust_text, cust_images = combine_asset_text("MÜŞTERİ BİLGİSİ", cust_assets)
                progress.progress(35, text="Teknik farklar ve tarifname dayanakları analiz ediliyor...")
                opinion = ask_json(gorus_prompt(report_type, reference, report_text, spec_text, prior_text, sim_text, cust_text), images=[*sim_images, *cust_images])
                validate_quotes(opinion, spec_text)
                progress.progress(85, text="Word görüş metni hazırlanıyor...")
                data = build_gorus_docx(opinion)
                progress.progress(100, text="Hazır")
                st.success("Görüş metni oluşturuldu.")
                st.download_button("Word dosyasını indir", data=data, file_name=safe_output_name(output_name, "Görüş metni.docx"), mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")
            except Exception as exc:
                st.exception(exc)

# ARAŞTIRMA
else:
    st.subheader("Tip 3 - Ön araştırma raporu")
    bbf = st.file_uploader("BBF dosyası", type=["docx", "doc", "pdf", "txt"], key="res_bbf")
    reference = st.text_input("DP referans numarası", value="")
    output_name = st.text_input("Çıktı dosyasının adı", value="Ön Araştırma Raporu_XXXXXX.docx")

    if "top10_result" not in st.session_state:
        st.session_state.top10_result = None
        st.session_state.research_bbf_text = None

    if st.button("1. Global araştırmayı yap ve en benzer 10 dokümanı bul", type="primary", use_container_width=True):
        if bbf is None:
            st.error("BBF yükleyin.")
        else:
            try:
                progress = st.progress(0, text="BBF okunuyor...")
                bbf_text = extract_text_from_asset(UploadedAsset(bbf.name, bbf.getvalue(), bbf.type))
                progress.progress(20, text="Global patent veritabanlarında araştırma yapılıyor...")
                top10 = ask_json(top10_research_prompt(bbf_text), web_search=True)
                docs = top10.get("documents") or []
                if len(docs) < 10:
                    raise ValueError(f"10 doküman yerine {len(docs)} doküman döndü. Araştırmayı tekrar çalıştırın.")
                st.session_state.top10_result = top10
                st.session_state.research_bbf_text = bbf_text
                progress.progress(100, text="Araştırma tamamlandı")
            except Exception as exc:
                st.exception(exc)

    if st.session_state.top10_result:
        st.success("En benzer 10 doküman bulundu.")
        rows = []
        for d in st.session_state.top10_result.get("documents") or []:
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

        own_docs = st.radio("Sizin araştırdığınız benzer doküman var mı?", ["Hayır", "Evet"], horizontal=True)
        user_files = []
        if own_docs == "Evet":
            user_files = st.file_uploader("Sizin bulduğunuz benzer dokümanlar", type=["pdf", "zip", "docx", "doc", "txt", "png", "jpg", "jpeg", "webp"], accept_multiple_files=True, key="res_user_docs")

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
                    selection = ask_json(final_selection_prompt(st.session_state.research_bbf_text, st.session_state.top10_result, user_text), images=user_images)
                    progress.progress(55, text="Yenilik ve buluş basamağı raporu hazırlanıyor...")
                    report = ask_json(report_drafting_prompt(st.session_state.research_bbf_text, st.session_state.top10_result, selection, reference))
                    progress.progress(85, text="Word raporu oluşturuluyor...")
                    data = build_research_docx(report)
                    progress.progress(100, text="Hazır")
                    st.success(f"Nihai D1: {selection.get('d1',{}).get('number','')} | Nihai D2: {(selection.get('d2') or {}).get('number','-')}")
                    st.info(f"Yenilik: {selection.get('novelty_result','')} | Buluş basamağı: {selection.get('inventive_step_result','')}")
                    st.download_button("Word raporunu indir", data=data, file_name=safe_output_name(output_name, "Ön Araştırma Raporu.docx"), mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")
                except Exception as exc:
                    st.exception(exc)
