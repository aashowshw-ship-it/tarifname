from __future__ import annotations

import io
import json
import os
import re
import subprocess
import tempfile
from copy import deepcopy
from pathlib import Path
from typing import Any

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt
from openai import OpenAI
from pypdf import PdfReader

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "Tarifname_181176_template.docx"
MODEL = os.getenv("OPENAI_MODEL", "gpt-5.6")

CORE_RULES = r"""
TÜRK PATENT TARİFNAME KURALLARI
1. Yalnızca yüklenen BBF'deki bilgi ve ifadeleri kullan. BBF'de bulunmayan teknik unsur, algoritma, değer, bağlantı, özellik veya kullanım biçimi ekleme.
2. İngilizce teknik terimi tarifnamede ilk geçtiği yerde İngilizcesi ve Türkçesiyle bir kez açıkla. Örnek: handover (hücre geçişi). Sonraki kullanımlarda yalnızca Türkçe karşılığını kullan. AI yerine yapay zekâ yaz.
3. Unsur adlarını normal cümle yazımıyla yaz; her kelimeyi büyük harfle başlatma.
4. BBF'deki unsur numaralandırmasını aynen koru. 1, 2, 3 ise aynı; 10, 20, 30 ise aynı. BBF'de unsur olarak verilmeyen sisteme, yönteme veya genel yapıya ek referans numarası verme.
5. REFERANS NUMARALARI bölümünden önce unsur veya işlem adımı referansı kullanma.
6. Referans listesinde önce unsurlar, ardından varsa yöntem işlem adımları yer alır. İşlem adımları '1001. ... toplanması' biçiminde yazılır.
7. Detaylı açıklamada ve yöntem isteminde işlem adımı referansı işlem ifadesinin sonunda bulunur: '... toplanması (1001)'. Numara cümlenin başına alınmaz.
8. Sistem istemindeki unsurları BBF sırasıyla tanımla. Bir unsur tanımlanırken yalnızca daha önce tanımlanmış unsurlarla teknik ilişki kur. Henüz tanımlanmamış bir unsuru önceki unsurun içinde kullanma.
9. Ana sistem isteminde tüm unsurlar birbirleriyle veri, kontrol, sinyal veya işlem ilişkisi içinde kurulmalıdır.
10. Alt istemler kısa olmalı, ana istemdeki aynı işlevi tekrar etmemelidir. Sistem alt istemleri uygun olduğunda 'bir modül olmasıdır' veya 'içermesidir' şeklinde bitmelidir. 'Yapmasıdır', 'etmesidir' veya 'belirlemesidir' şeklinde bitirme.
11. Açık ve sıralı bir bilgisayar tarafından gerçekleştirilen işlem akışı BBF'de mevcutsa yöntem istemi oluştur. İşlem akışı yoksa yalnızca sistem istemi oluştur.
12. Yöntem istemindeki işlem adımları BBF'deki 1001, 1002... sırasına ve metnine sadık olmalı; her işlem adımının numarası sonda parantez içinde gösterilmelidir.
13. Şablondaki tarifname giriş açıklamasını ve İSTEMLER altındaki kırmızı/mavi açıklamaları biçimleriyle aynen koru.
14. Teknik kişi veya operatöre yönelik işlemlerde teknik araç belirt: örneğin 'elektronik cihaz üzerinden operatöre ileten'.
15. Önceki teknik, detaylı açıklama, istemler, referans listesi ve özet birbiriyle tutarlı olmalıdır.
"""


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


def get_client() -> OpenAI:
    key = os.getenv("OPENAI_API_KEY", "").strip()
    if not key:
        raise RuntimeError(
            "OPENAI_API_KEY Render servisine tanımlı değil. Render > tarifname-atolyesi > Environment bölümünde "
            "Key alanına OPENAI_API_KEY, Value alanına sk- ile başlayan anahtarı girip Save, rebuild and deploy seçin."
        )
    return OpenAI(api_key=key)


def ask_json(prompt: str, *, web_search: bool = False) -> dict[str, Any]:
    client = get_client()
    kwargs: dict[str, Any] = {"model": MODEL, "input": prompt}
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
            values = [cell.text.strip().replace("\n", " | ") for cell in row.cells]
            if any(values):
                parts.append("\t".join(values))
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
                ["antiword", str(source)],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=60,
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
            raise ValueError("Eski .doc dosyası okunamadı. Dosyayı .docx olarak kaydedip yeniden yükleyin.") from exc
    raise ValueError("Eski .doc dosyası okunamadı.")


def extract_text(filename: str, data: bytes) -> str:
    suffix = Path(filename).suffix.lower()
    if suffix == ".docx":
        text = docx_text(data)
    elif suffix == ".doc":
        text = legacy_doc_text(data, filename)
    elif suffix == ".pdf":
        text = pdf_text(data)
    elif suffix in {".txt", ".md"}:
        text = data.decode("utf-8", errors="replace")
    else:
        raise ValueError("Desteklenen dosya türleri: .docx, .doc, .pdf ve .txt")
    text = text.replace("\x00", " ").strip()
    if not text:
        raise ValueError("Dosyadan metin çıkarılamadı.")
    return text


def extraction_prompt(source_text: str) -> str:
    return f"""{CORE_RULES}

Aşağıdaki BBF'yi yalnızca yapılandırılmış veri çıkarmak için incele. Bilgileri düzeltme, genişletme veya yeni teknik bilgi ekleme. JSON dışında hiçbir şey yazma.

ÇIKTI ŞEMASI:
{{
  "title":"",
  "technical_field":"",
  "prior_art":[""],
  "problems":[""],
  "advantages":[""],
  "elements":[{{"number":"10","name":"","function":""}}],
  "method_steps":[{{"number":"1001","text":""}}],
  "working_principle":[""],
  "keywords":[""],
  "figures":["Şekil 1, ... gösterimidir."],
  "has_method_basis":true,
  "method_basis_reason":""
}}

BBF:
---
{source_text}
---
"""


def literature_prompt(extracted: dict[str, Any], count: int, jurisdiction: str) -> str:
    return f"""Aşağıdaki buluş için patent literatüründe tam olarak {count} farklı ve teknik olarak en yakın patent dokümanı araştır.
Web araması yap. Doküman uydurma. Yayın/başvuru numarasını ve başlığı doğrulanabilir patent kaynağından kontrol et.
Tercih edilen ülke veya kaynak kapsamı: {jurisdiction or 'ülke sınırlaması yok'}.
Her doküman için buluşun kaynaktaki gerçek içeriğine dayalı kısa açıklama ve mevcut buluştan eksik kalan temel sistemi belirt.
'Difference' alanı, '... sistem' veya '... sistemi' şeklinde bitmeli ve 'ile ilgili bir emareye rastlanmamıştır' ifadesinden önce dilbilgisel olarak kullanılabilmelidir.
JSON dışında hiçbir şey yazma.

ŞEMA:
{{
  "documents":[
    {{
      "application_number":"",
      "title":"",
      "jurisdiction":"Türkiye/ABD/Avrupa/WIPO/Çin vb.",
      "summary":"... sistemi",
      "difference":"... sistemi",
      "source_url":""
    }}
  ]
}}

BULUŞ VERİSİ:
{json.dumps(extracted, ensure_ascii=False, indent=2)}
"""


def format_literature_paragraphs(documents: list[dict[str, Any]], count: int) -> list[str]:
    paragraphs: list[str] = []
    for doc in documents[:count]:
        number = str(doc.get("application_number", "")).strip()
        title = str(doc.get("title", "")).strip()
        jurisdiction = str(doc.get("jurisdiction", "")).strip() or "ilgili"
        summary = str(doc.get("summary", "")).strip().rstrip(".")
        difference = str(doc.get("difference", "")).strip().rstrip(".")
        if not all([number, title, summary, difference]):
            continue
        paragraphs.append(
            f"Literatürde yapılan araştırmalar sonucu “{number}” başvuru numaralı ve “{title}” buluş başlıklı "
            f"{jurisdiction} patent müracaatına rastlanmıştır. Söz konusu başvuru, {summary} ile ilgilidir. "
            f"Ancak bahsedilen başvuruda {difference} ile ilgili bir emareye rastlanmamıştır."
        )
    if len(paragraphs) != count:
        raise ValueError(f"İstenen {count} benzer dokümandan yalnızca {len(paragraphs)} tanesi doğrulanabildi.")
    return paragraphs


def drafting_prompt(extracted: dict[str, Any], claim_mode: str) -> str:
    return f"""{CORE_RULES}

Aşağıdaki yapılandırılmış BBF verilerinden Türk patent tarifnamesi oluştur. BBF'de olmayan teknik bilgi ekleme.
İstem yapısı: {claim_mode}
JSON dışında hiçbir şey yazma.

ÇIKTI ŞEMASI:
{{
  "title":"",
  "technical_field":"",
  "prior_art_general_paragraphs":[""],
  "short_description_intro":"",
  "objectives":[""],
  "unumbered_system_definition":"",
  "unumbered_system_elements":[""],
  "figure_descriptions":[""],
  "elements":[{{"number":"10","name":"","description":""}}],
  "method_steps":[{{"number":"1001","text":""}}],
  "detailed_paragraphs":[""],
  "system_claim":{{"preamble":"","elements":[""],"closing":"içermesidir."}},
  "dependent_system_claims":[""],
  "method_claim":null,
  "dependent_method_claims":[""],
  "abstract":""
}}

ÖZEL TALİMATLAR:
- prior_art_general_paragraphs yalnızca BBF'deki mevcut uygulamalar ve teknik problemlerden oluşturulsun; patent dokümanı uydurulmasın.
- unnumbered_system_definition ve unnumbered_system_elements REFERANS NUMARALARI bölümünden önce kullanılacağı için hiçbir referans numarası içermesin.
- elements alanındaki numara ve unsur adları BBF ile aynen uyumlu olsun.
- method_steps alanında numara ayrı olsun; text alanı numarasız ve 'toplanması', 'analiz edilmesi' gibi isim-fiille bitsin.
- Yöntem istemi oluşturulacaksa her işlem adımı BBF sırasıyla ve sonunda '(1001)' biçiminde yer alsın.
- Sistem istemi unsurları BBF sırasıyla tanımlansın. İlk unsur kendi teknik işleviyle tanımlansın; sonraki her unsur yalnızca daha önce tanımlanan unsurlarla ilişkilendirilsin.
- Alt istemler ana istemi tekrar etmesin ve gereksiz ayrıntı içermesin.

BBF VERİSİ:
{json.dumps(extracted, ensure_ascii=False, indent=2)}
"""


def validate_draft(draft: dict[str, Any]) -> None:
    elements = draft.get("elements") or []
    numbers = [str(x.get("number", "")).strip() for x in elements]
    if not numbers or any(not n for n in numbers):
        raise ValueError("Unsur listesinde eksik referans numarası bulundu.")
    if len(numbers) != len(set(numbers)):
        raise ValueError("Unsur listesinde tekrarlanan referans numarası bulundu.")

    before_refs = "\n".join([
        str(draft.get("technical_field", "")),
        *map(str, draft.get("prior_art_paragraphs") or []),
        str(draft.get("short_description_intro", "")),
        *map(str, draft.get("objectives") or []),
        str(draft.get("unumbered_system_definition", "")),
        *map(str, draft.get("unumbered_system_elements") or []),
    ])
    if re.search(r"\(\s*(?:\d{1,3}|10\d{2})\s*\)", before_refs):
        raise ValueError("REFERANS NUMARALARI bölümünden önce parantezli referans numarası kullanılmış.")

    method = draft.get("method_claim")
    steps = draft.get("method_steps") or []
    if method:
        text = " ".join(method.get("steps") or [])
        for step in steps:
            number = str(step.get("number", "")).strip()
            if number and f"({number})" not in text:
                raise ValueError(f"Yöntem isteminde ({number}) işlem adımı referansı eksik.")

    for claim in draft.get("dependent_system_claims") or []:
        if re.search(r"(?:yapmasıdır|etmesidir|belirlemesidir)\.?$", claim.strip(), re.I):
            raise ValueError("Sistem alt istemlerinden biri yanlış fiil sonuyla bitiyor.")


def clear_body(doc: Document) -> None:
    body = doc._element.body
    for child in list(body):
        if child.tag.endswith("sectPr"):
            continue
        body.remove(child)


def format_paragraph(p, *, bold: bool = False, center: bool = False):
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(11)
        if bold:
            run.bold = True
    return p


def add_text(doc: Document, text: str, *, bold: bool = False, center: bool = False):
    p = doc.add_paragraph()
    p.add_run(text)
    return format_paragraph(p, bold=bold, center=center)


def add_heading(doc: Document, text: str):
    return add_text(doc, text, bold=True)


def add_bullet(doc: Document, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(1.905)
    p.paragraph_format.first_line_indent = Cm(-0.635)
    p.paragraph_format.line_spacing = 1.5
    p.add_run("•\t" + text)
    return format_paragraph(p)


def copy_template_paragraph(doc: Document, template: Document, index: int):
    doc._element.body.insert(-1, deepcopy(template.paragraphs[index]._p))


def numbered_claim(doc: Document, number: int, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(f"{number}. {text}")
    run.font.name = "Arial"
    run.font.size = Pt(11)


def add_page_number(section):
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType", "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instruction.text = "PAGE"
    end = OxmlElement("w:fldChar")
    end.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType", "end")
    run._r.extend([begin, instruction, end])


def build_docx(draft: dict[str, Any]) -> bytes:
    template = Document(str(TEMPLATE_PATH))
    doc = Document(str(TEMPLATE_PATH))
    clear_body(doc)
    for section in doc.sections:
        section.top_margin = Cm(3)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(3)
        section.right_margin = Cm(2)

    add_text(doc, "TARİFNAME", bold=True, center=True)
    doc.add_paragraph()
    add_text(doc, draft["title"], bold=True, center=True)
    doc.add_paragraph()
    copy_template_paragraph(doc, template, 4)
    doc.add_paragraph()

    add_heading(doc, "TEKNİK ALAN")
    doc.add_paragraph()
    add_text(doc, draft.get("technical_field", ""))
    doc.add_paragraph()

    add_heading(doc, "ÖNCEKİ TEKNİK")
    doc.add_paragraph()
    for paragraph in draft.get("prior_art_paragraphs") or []:
        add_text(doc, paragraph)
        doc.add_paragraph()
    add_text(doc, "Sonuçta yukarıda bahsedilen ve mevcut teknik ışığında çözülemeyen sorunlar, ilgili teknik alanda bir yenilik yapmayı zorunlu kılmıştır.")

    add_heading(doc, "BULUŞUN KISA AÇIKLAMASI")
    doc.add_paragraph()
    add_text(doc, draft.get("short_description_intro", ""))
    for i, objective in enumerate(draft.get("objectives") or []):
        prefix = "Buluşun ana amacı, " if i == 0 else "Buluşun diğer bir amacı, "
        objective = str(objective).strip()
        add_text(doc, prefix + (objective[:1].lower() + objective[1:] if objective else ""))
        doc.add_paragraph()

    if draft.get("unumbered_system_definition"):
        add_text(doc, draft["unumbered_system_definition"])
    for element in draft.get("unumbered_system_elements") or []:
        add_bullet(doc, element)
    if draft.get("unumbered_system_elements"):
        add_text(doc, "içermesidir.")
    doc.add_paragraph()
    add_text(doc, "Mevcut buluşun yapılanması ve ek elemanlarla birlikte avantajlarının en iyi şekilde anlaşılabilmesi için aşağıda açıklaması yapılan şekiller ile birlikte değerlendirilmesi gerekir.")

    add_heading(doc, "ŞEKİLLERİN KISA AÇIKLAMASI")
    doc.add_paragraph()
    figures = draft.get("figure_descriptions") or ["Şekil 1, buluşa konu sistemin temsili bir gösterimidir."]
    for figure in figures:
        add_text(doc, figure)
    add_text(doc, "Çizimlerin mutlaka ölçeklendirilmesi gerekmemektedir ve mevcut buluşu anlamak için gerekli olmayan detaylar ihmal edilmiş olabilmektedir. Bundan başka, en azından büyük ölçüde özdeş olan veya en azından büyük ölçüde özdeş işlevleri olan elemanlar, aynı numara ile gösterilmektedir.")

    add_heading(doc, "REFERANS NUMARALARI")
    doc.add_paragraph()
    for element in draft.get("elements") or []:
        add_text(doc, f"{element['number']}. {element['name']}")
        doc.add_paragraph()
    for step in draft.get("method_steps") or []:
        add_text(doc, f"{step['number']}. {step['text']}")
        doc.add_paragraph()

    add_heading(doc, "BULUŞUN DETAYLI AÇIKLAMASI")
    doc.add_paragraph()
    add_text(doc, f"Bu detaylı açıklamada, buluş konusu olan {draft['title'].lower()} sadece konunun daha iyi anlaşılmasına yönelik hiçbir sınırlayıcı etki oluşturmayacak örneklerle açıklanmaktadır.")
    for paragraph in draft.get("detailed_paragraphs") or []:
        doc.add_paragraph()
        add_text(doc, paragraph)

    add_heading(doc, "İSTEMLER")
    doc.add_paragraph()
    for index in (79, 81, 83):
        copy_template_paragraph(doc, template, index)
        doc.add_paragraph()

    claim_no = 1
    system_claim = draft.get("system_claim") or {}
    numbered_claim(doc, claim_no, system_claim.get("preamble", "") + " olup, özelliği;")
    for element in system_claim.get("elements") or []:
        add_bullet(doc, element)
    add_text(doc, system_claim.get("closing", "içermesidir."))
    claim_no += 1

    for dependent in draft.get("dependent_system_claims") or []:
        numbered_claim(doc, claim_no, dependent)
        claim_no += 1

    method_claim = draft.get("method_claim")
    if method_claim:
        numbered_claim(doc, claim_no, method_claim.get("preamble", "") + " olup, özelliği;")
        for step in method_claim.get("steps") or []:
            add_bullet(doc, step)
        add_text(doc, method_claim.get("closing", "işlem adımlarını içermesidir."))
        claim_no += 1
        for dependent in draft.get("dependent_method_claims") or []:
            numbered_claim(doc, claim_no, dependent)
            claim_no += 1

    add_heading(doc, "ÖZET")
    doc.add_paragraph()
    add_text(doc, draft["title"], bold=True, center=True)
    doc.add_paragraph()
    add_text(doc, draft.get("abstract", ""))

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


st.set_page_config(page_title="Tarifname Atölyesi", page_icon="📄", layout="centered")
st.markdown(
    """
    <style>
      .block-container {max-width: 850px; padding-top: 2rem;}
      .hero {padding: 1.3rem 1.5rem; border:1px solid #e8e8e8; border-radius:16px; margin-bottom:1.2rem;}
      .hero h1 {margin:0 0 .25rem 0; font-size:2rem;}
      .hero p {margin:0; color:#666;}
    </style>
    <div class="hero"><h1>Tarifname Atölyesi</h1><p>BBF'yi yükleyin ve doğrudan Word tarifnamesini oluşturun.</p></div>
    """,
    unsafe_allow_html=True,
)

if not os.getenv("OPENAI_API_KEY", "").strip():
    st.error(
        "OPENAI_API_KEY tanımlı değil. Render'da tarifname-atolyesi servisini açın → Environment → "
        "Add Environment Variable → Key: OPENAI_API_KEY → Value: sk- ile başlayan anahtar → Save, rebuild and deploy."
    )
    st.stop()

with st.form("draft_form"):
    uploaded = st.file_uploader("BBF dosyası", type=["docx", "doc", "pdf", "txt"])
    output_name = st.text_input("Çıktı dosyasının adı", value="Tarifname_181710.docx")
    claim_choice = st.selectbox(
        "İstem yapısı",
        ["BBF'ye göre otomatik belirle", "Yalnızca sistem", "Sistem ve yöntem"],
    )
    research = st.checkbox("Literatür araştırması yap ve benzer patentleri önceki tekniğe ekle", value=False)
    col1, col2 = st.columns(2)
    with col1:
        similar_count = st.number_input("Benzer doküman sayısı", min_value=1, max_value=10, value=2, disabled=not research)
    with col2:
        jurisdiction = st.text_input("Tercih edilen ülke/veri tabanı", value="", disabled=not research, placeholder="Örn. Türkiye, EP, WO")
    submit = st.form_submit_button("Tarifnameyi oluştur", type="primary", use_container_width=True)

if submit:
    if uploaded is None:
        st.error("Önce BBF dosyasını yükleyin.")
        st.stop()
    safe_name = output_name.strip() or "Tarifname.docx"
    if not safe_name.lower().endswith(".docx"):
        safe_name += ".docx"

    try:
        progress = st.progress(0, text="BBF okunuyor...")
        source_text = extract_text(uploaded.name, uploaded.getvalue())
        progress.progress(15, text="BBF bilgileri çıkarılıyor...")
        extracted = ask_json(extraction_prompt(source_text))

        if claim_choice == "BBF'ye göre otomatik belirle":
            claim_mode = "Sistem ve yöntem" if extracted.get("has_method_basis") else "Yalnızca sistem"
        else:
            claim_mode = claim_choice

        literature_paragraphs: list[str] = []
        if research:
            progress.progress(35, text=f"{int(similar_count)} benzer patent araştırılıyor...")
            literature = ask_json(
                literature_prompt(extracted, int(similar_count), jurisdiction.strip()),
                web_search=True,
            )
            literature_paragraphs = format_literature_paragraphs(
                literature.get("documents") or [], int(similar_count)
            )

        progress.progress(60, text="Tarifname ve istemler hazırlanıyor...")
        draft = ask_json(drafting_prompt(extracted, claim_mode))
        draft["prior_art_paragraphs"] = (draft.pop("prior_art_general_paragraphs", []) or []) + literature_paragraphs

        if claim_mode == "Yalnızca sistem":
            draft["method_claim"] = None
            draft["dependent_method_claims"] = []
        validate_draft(draft)

        progress.progress(85, text="Word dosyası oluşturuluyor...")
        docx_bytes = build_docx(draft)
        progress.progress(100, text="Tarifname hazır.")
        st.success("Tarifname oluşturuldu.")
        st.download_button(
            "Word tarifnamesini indir",
            data=docx_bytes,
            file_name=safe_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
    except Exception as exc:
        st.error(str(exc))
