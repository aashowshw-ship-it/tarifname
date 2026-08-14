from __future__ import annotations

import io
import re
import subprocess
import tempfile
from pathlib import Path
from typing import Any, Iterable

import fitz
from docx import Document
from docx.oxml.ns import qn


def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def _to_pdf_bytes(filename: str, data: bytes) -> bytes:
    suffix = Path(filename).suffix.lower()
    if suffix == ".pdf":
        return data
    if suffix not in {".doc", ".docx"}:
        raise ValueError("Sayfa/satır doğrulaması için tarifname PDF, DOC veya DOCX olmalıdır.")
    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        src = td_path / Path(filename).name
        src.write_bytes(data)
        outdir = td_path / "pdf"
        outdir.mkdir()
        proc = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(src)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=120,
        )
        if proc.returncode != 0:
            raise ValueError("Tarifname sayfa/satır doğrulaması için PDF'e dönüştürülemedi.")
        pdfs = list(outdir.glob("*.pdf"))
        if not pdfs:
            raise ValueError("Tarifname PDF dönüşümü üretilemedi.")
        return pdfs[0].read_bytes()


def _page_lines(page: fitz.Page) -> list[dict[str, Any]]:
    words = page.get_text("words")
    grouped: dict[tuple[int, int], list[tuple]] = {}
    for w in words:
        grouped.setdefault((int(w[5]), int(w[6])), []).append(w)
    lines: list[dict[str, Any]] = []
    for ws in grouped.values():
        ws = sorted(ws, key=lambda x: x[0])
        text = " ".join(str(w[4]) for w in ws).strip()
        if not text:
            continue
        lines.append({
            "text": text,
            "x0": min(float(w[0]) for w in ws),
            "y0": min(float(w[1]) for w in ws),
            "y1": max(float(w[3]) for w in ws),
        })
    lines.sort(key=lambda x: (round(x["y0"], 1), x["x0"]))
    return lines


def build_page_line_index(filename: str, data: bytes) -> list[dict[str, Any]]:
    """Build page/physical-line index from Word-style printed line numbers.

    Patent specification templates commonly print every fifth line in the left
    margin. We anchor to those numbers and interpolate the intervening rendered
    text lines. This makes page/line citations deterministic rather than LLM guesses.
    """
    pdf = fitz.open(stream=_to_pdf_bytes(filename, data), filetype="pdf")
    indexed: list[dict[str, Any]] = []
    for page_no, page in enumerate(pdf, start=1):
        raw = _page_lines(page)
        width = float(page.rect.width)
        anchors: list[tuple[int, float]] = []
        body: list[dict[str, Any]] = []
        for line in raw:
            t = line["text"].strip()
            m = re.fullmatch(r"(\d{1,3})", t)
            if m and line["x0"] < width * 0.22 and int(m.group(1)) % 5 == 0:
                anchors.append((int(m.group(1)), (line["y0"] + line["y1"]) / 2))
            else:
                # Header page number at top center is not part of line-numbered body.
                if re.fullmatch(r"\d+", t) and line["y0"] < 90:
                    continue
                body.append(line)
        if not anchors:
            # No printed line numbers. Preserve page structure but mark line unknown.
            for line in body:
                indexed.append({"page": page_no, "line": None, "text": line["text"], "y": line["y0"]})
            continue
        # Word line numbering also counts blank rendered lines. Therefore an every-fifth
        # printed number can sit on a blank baseline. Infer the vertical line grid from
        # the printed anchors instead of snapping anchors only to text lines.
        pitches: list[float] = []
        a_sorted = sorted(anchors, key=lambda x: x[0])
        for (n1, y1), (n2, y2) in zip(a_sorted, a_sorted[1:]):
            if n2 > n1 and y2 > y1:
                pitches.append((y2 - y1) / float(n2 - n1))
        if pitches:
            pitches.sort()
            pitch = pitches[len(pitches)//2]
        else:
            # Typical 11-pt / 1.5-spaced patent template fallback.
            pitch = 17.5
        for line in body:
            y = (line["y0"] + line["y1"]) / 2
            n0, y0 = min(anchors, key=lambda a: abs(a[1] - y))
            line_no = int(n0 + round((y - y0) / max(1.0, pitch)))
            if line_no < 1 or line_no > 80:
                line_no = None
            indexed.append({"page": page_no, "line": line_no, "text": line["text"], "y": line["y0"]})
    return indexed


def locate_quote_page_lines(filename: str, data: bytes, quote: str) -> tuple[int, int, int]:
    q = _norm(quote)
    if not q:
        raise ValueError("Boş tarifname alıntısı için sayfa/satır bulunamaz.")
    index = build_page_line_index(filename, data)
    pages = sorted({int(x["page"]) for x in index})
    for pno in pages:
        lines = [x for x in index if int(x["page"]) == pno]
        pieces: list[str] = []
        spans: list[tuple[int, int, dict[str, Any]]] = []
        cursor = 0
        for ln in lines:
            t = _norm(ln["text"])
            if not t:
                continue
            if pieces:
                cursor += 1
            start = cursor
            pieces.append(t)
            cursor += len(t)
            spans.append((start, cursor, ln))
        joined = " ".join(pieces)
        pos = joined.find(q)
        if pos < 0:
            continue
        end = pos + len(q)
        hit = [ln for s, e, ln in spans if e > pos and s < end]
        nums = [int(ln["line"]) for ln in hit if ln.get("line") is not None]
        if not nums:
            raise ValueError(f"Tarifname alıntısı sayfa {pno}'da bulundu ancak basılı satır numaraları doğrulanamadı.")
        return pno, min(nums), max(nums)
    raise ValueError(f"Tarifname alıntısı fiziksel sayfa metninde birebir bulunamadı: {q[:140]}...")


def _iter_quote_objects(opinion: dict[str, Any]):
    for section in opinion.get("sections") or []:
        for block in section.get("blocks") or []:
            if str(block.get("type", "")).lower() == "quote":
                yield block
        for quote in section.get("quotes") or []:  # legacy schema compatibility
            yield quote


def annotate_quote_locations(opinion: dict[str, Any], spec_filename: str, spec_bytes: bytes) -> None:
    for q in _iter_quote_objects(opinion):
        text = str(q.get("text", "")).strip()
        if not text:
            continue
        page, start, end = locate_quote_page_lines(spec_filename, spec_bytes, text)
        dash = "-"
        q["page"] = page
        q["line_start"] = start
        q["line_end"] = end
        q["lead"] = f"Tarifname sayfa {page}, satır {start}{dash}{end}’te bu durum şu şekilde belirtilmiştir:"



def _edit_distance_le1(a: str, b: str) -> bool:
    a, b = str(a), str(b)
    if a == b:
        return True
    if abs(len(a) - len(b)) > 1:
        return False
    if len(a) == len(b):
        return sum(x != y for x, y in zip(a, b)) <= 1
    if len(a) > len(b):
        a, b = b, a
    # b is longer by one
    i = j = diff = 0
    while i < len(a) and j < len(b):
        if a[i] == b[j]:
            i += 1; j += 1
        else:
            diff += 1; j += 1
            if diff > 1:
                return False
    return True

def validate_opinion_payload(opinion: dict[str, Any], report_text: str, spec_text: str) -> None:
    for field, label in [("application_no", "Başvuru No"), ("applicant", "Başvuru Sahibi"), ("reference", "Referans")]:
        if not _norm(opinion.get(field, "")):
            raise ValueError(f"Görüş metadata kapısı: {label} boş bırakılamaz. Kaynaktan bulunamıyorsa arayüzden girin.")
    report_norm = _norm(report_text).casefold()
    report_keys = re.findall(r"[A-Z]{2}[0-9A-Z./-]{6,}", str(report_text or "").upper())
    report_keys = [_publication_key(x) for x in report_keys]
    spec_norm = _norm(spec_text)
    docs = opinion.get("cited_documents") or []
    for d in docs:
        num = _norm(d.get("number", ""))
        key = _publication_key(num)
        if num and key not in _publication_key(report_norm) and not any(_edit_distance_le1(key, rk) for rk in report_keys):
            raise ValueError(f"Görüşte raporda doğrulanamayan doküman numarası var: {num}")
    for q in _iter_quote_objects(opinion):
        text = _norm(q.get("text", ""))
        if text and text not in spec_norm:
            raise ValueError(f"Tarifname alıntısı birebir doğrulanamadı: {text[:140]}...")
    combined = opinion.get("combined_assessment") or {}
    combined_text = _norm(" ".join(combined.get("paragraphs") or []))
    if "buluş basamağı" in report_norm or "buluş basamagi" in report_norm:
        required_concepts = [
            ("teknik fark", "ayırt edici", "farklı"),
            ("teknik etki", "etki"),
            ("teknik problem", "problem"),
            ("motivasyon", "yönlendirme"),
            ("geriye dönük", "hindsight"),
        ]
        low = combined_text.casefold()
        missing = [alts[0] for alts in required_concepts if not any(a in low for a in alts)]
        if missing:
            raise ValueError("Dokümanların birlikte değerlendirilmesi bölümü buluş basamağı zincirini eksik kuruyor: " + ", ".join(missing))
        if len(combined_text) < 1200:
            raise ValueError("Dokümanların birlikte değerlendirilmesi bölümü buluş basamağı itirazı için yeterince ayrıntılı değil.")


def _asset_name(asset: Any) -> str:
    if isinstance(asset, dict):
        return str(asset.get("name", ""))
    return str(getattr(asset, "name", ""))


def _asset_data(asset: Any) -> bytes:
    if isinstance(asset, dict):
        return bytes(asset.get("data", b""))
    return bytes(getattr(asset, "data", b""))


def _publication_key(text: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(text or "").upper())


def _best_figure_page_png(pdf_data: bytes) -> bytes:
    pdf = fitz.open(stream=pdf_data, filetype="pdf")
    if not pdf.page_count:
        raise ValueError("Patent PDF boş.")
    scored: list[tuple[float, int]] = []
    pats = [
        re.compile(r"(?i)\bFIG(?:URE)?\.?\s*\d+"),
        re.compile(r"【\s*図\s*\d+\s*】"),
        re.compile(r"图\s*\d+"),
    ]
    for i, page in enumerate(pdf):
        text = page.get_text("text") or ""
        markers = sum(len(p.findall(text)) for p in pats)
        # Figure sheets tend to have many figure labels and relatively little prose.
        score = markers * 100.0 - min(len(text), 6000) / 800.0
        if markers:
            score += i * 0.01
        scored.append((score, i))
    score, idx = max(scored)
    if score <= 0:
        # Safe fallback: choose page with least text after bibliographic first page.
        candidates = [(len((pdf[i].get_text("text") or "")), i) for i in range(min(1, pdf.page_count - 1), pdf.page_count)]
        idx = min(candidates)[1]
    page = pdf[idx]
    pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2), alpha=False)
    png = pix.tobytes("png")
    # Crop only uniform outer whitespace; no technical geometry is altered.
    try:
        from PIL import Image, ImageChops
        im = Image.open(io.BytesIO(png)).convert("RGB")
        bg = Image.new("RGB", im.size, "white")
        diff = ImageChops.difference(im, bg).convert("L")
        bbox = diff.point(lambda x: 0 if x < 12 else 255).getbbox()
        if bbox:
            left, top, right, bottom = bbox
            pad = 24
            crop = im.crop((max(0, left-pad), max(0, top-pad), min(im.width, right+pad), min(im.height, bottom+pad)))
            out = io.BytesIO(); crop.save(out, format="PNG"); png = out.getvalue()
    except Exception:
        pass
    return png


def extract_cited_original_figure_pages(cited_documents: list[dict[str, Any]], assets: Iterable[Any]) -> dict[str, bytes]:
    asset_list = list(assets or [])
    out: dict[str, bytes] = {}
    for d in cited_documents or []:
        label = str(d.get("label", "")).strip()
        number = str(d.get("number", "")).strip()
        if not label or not number:
            continue
        key = _publication_key(number)
        candidates = []
        for a in asset_list:
            name = _asset_name(a)
            if Path(name).suffix.lower() != ".pdf":
                continue
            name_key = _publication_key(name)
            tokens = re.findall(r"[A-Z]{2}[0-9A-Z]{6,}", name_key) or [name_key]
            if key and (key in name_key or any(_edit_distance_le1(key, tok) for tok in tokens)):
                # Prefer original-language patent PDFs over browser EN exports for figures.
                penalty = 1 if re.search(r"(?i)(?:^|[-_ ])EN(?:[-_ .]|$)", Path(name).stem) else 0
                candidates.append((penalty, len(name), a))
        if not candidates:
            continue
        candidates.sort(key=lambda x: (x[0], x[1]))
        out[label] = _best_figure_page_png(_asset_data(candidates[0][2]))
    return out


def _geom(section) -> tuple[int, ...]:
    return tuple(int(x) for x in (
        section.page_width, section.page_height,
        section.top_margin, section.bottom_margin, section.left_margin, section.right_margin,
        section.header_distance, section.footer_distance,
    ))


def _paragraph_text_from_element(el) -> str:
    return "".join(t.text or "" for t in el.iter() if t.tag == qn("w:t")).strip()


def _body_sequence(doc: Document) -> list[tuple[str, str, Any]]:
    seq: list[tuple[str, str, Any]] = []
    table_by_id = {id(t._tbl): t for t in doc.tables}
    for el in doc._element.body:
        if el.tag == qn("w:p"):
            seq.append(("p", _paragraph_text_from_element(el), el))
        elif el.tag == qn("w:tbl"):
            t = table_by_id.get(id(el))
            cap = ""
            if t is not None and t.rows and t.rows[0].cells:
                cap = t.rows[0].cells[0].text.strip()
            else:
                cap = _paragraph_text_from_element(el)
            seq.append(("t", cap, el))
    return seq


def validate_gorus_template_fidelity(docx_data: bytes, template_path: str | Path, opinion: dict[str, Any], required_figure_labels: Iterable[str] = ()) -> None:
    doc = Document(io.BytesIO(docx_data))
    tpl = Document(str(template_path))
    if not doc.sections or _geom(doc.sections[0]) != _geom(tpl.sections[0]):
        raise ValueError("Görüş şablon kapısı: sayfa geometrisi/marj/header-footer mesafeleri 696809 şablonuyla aynı değil.")
    seq = _body_sequence(doc)
    # Binding opening archetype: two institutional titles -> metadata table -> physical blank -> salutation -> intro -> physical blank.
    if len(seq) < 7 or [x[0] for x in seq[:7]] != ["p", "p", "t", "p", "p", "p", "p"]:
        raise ValueError("Görüş şablon kapısı: giriş öğelerinin paragraf/tablo sırası 696809 taslağıyla aynı değil.")
    if seq[0][1] != tpl.paragraphs[0].text.strip() or seq[1][1] != tpl.paragraphs[1].text.strip():
        raise ValueError("Görüş şablon kapısı: kurum başlıkları bağlayıcı taslakla aynı değil.")
    if seq[3][1] != "" or seq[4][1] != "Sayın Uzman," or seq[6][1] != "":
        raise ValueError("Görüş şablon kapısı: girişteki fiziksel boş paragraf / `Sayın Uzman,` düzeni taslağa uymuyor.")
    texts = [p.text.strip() for p in doc.paragraphs]
    if len(texts) < 6 or texts[0] != tpl.paragraphs[0].text.strip() or texts[1] != tpl.paragraphs[1].text.strip():
        raise ValueError("Görüş şablon kapısı: kurum başlıkları bağlayıcı şablonla aynı değil.")
    if "Sayın Uzman," not in texts:
        raise ValueError("Görüş şablon kapısı: `Sayın Uzman,` girişi eksik.")
    if not doc.tables or len(doc.tables[0].rows) != 3 or len(doc.tables[0].columns) != 3:
        raise ValueError("Görüş şablon kapısı: metadata tablosu 3x3 değil.")
    labels = [doc.tables[0].rows[i].cells[0].text.strip() for i in range(3)]
    if labels != ["Başvuru No", "Başvuru Sahibi", "Referans"]:
        raise ValueError("Görüş şablon kapısı: metadata etiketleri bozulmuş.")
    # Body paragraphs must remain Arial 11 and 1.5-spaced; first two institutional headings are exempt.
    sal_idx = texts.index("Sayın Uzman,")
    for p in doc.paragraphs[sal_idx:]:
        if not p.text.strip():
            continue
        if p.text.strip() in {"Saygılarımızla,", "DESTEK PATENT A.Ş."}:
            continue
        spacing = p.paragraph_format.line_spacing
        if spacing is not None and abs(float(spacing) - 1.5) > 0.01:
            raise ValueError(f"Görüş şablon kapısı: 1,5 satır aralığından sapma: {p.text[:80]}")
        for r in p.runs:
            if not r.text:
                continue
            if r.font.name not in {None, "Arial"}:
                raise ValueError(f"Görüş şablon kapısı: Arial dışı font bulundu: {r.font.name}")
            if r.font.size is not None and abs(r.font.size.pt - 11) > 0.2:
                raise ValueError(f"Görüş şablon kapısı: 11 punto dışı gövde metni bulundu: {r.font.size.pt}")
    # Every required D-label must have a 2-row original-figure table with a drawing.
    found: set[str] = set()
    # Template figure archetype has two physical blank paragraphs immediately before the D-figure table.
    for idx, item in enumerate(seq):
        if item[0] != "t" or not re.match(r"^D\d+\s+dokümanı\s*-\s*Şekil", item[1], flags=re.I):
            continue
        if idx < 2 or seq[idx-1][0] != "p" or seq[idx-2][0] != "p" or seq[idx-1][1] != "" or seq[idx-2][1] != "":
            raise ValueError(f"Görüş şablon kapısı: `{item[1]}` tablosundan önce taslaktaki iki fiziksel boş paragraf yok.")
    for t in doc.tables[1:]:
        if len(t.rows) < 2:
            continue
        cap = t.rows[0].cells[0].text.strip() if t.rows and t.rows[0].cells else ""
        m = re.match(r"^(D\d+)\s+dokümanı\s*-\s*Şekil", cap, flags=re.I)
        if not m:
            continue
        label = m.group(1).upper()
        xml = t.rows[1].cells[0]._tc.xml
        if "w:drawing" not in xml and "w:pict" not in xml:
            raise ValueError(f"Görüş şekil kapısı: {label} şekil tablosunda özgün görsel yok.")
        found.add(label)
    required = {str(x).upper() for x in required_figure_labels if str(x).strip()}
    missing = sorted(required - found)
    if missing:
        raise ValueError("Görüş şekil kapısı: özgün şekli eksik dokümanlar: " + ", ".join(missing))
    # Physical page/line quote lead + bold verbatim quote must be visible in the same paragraph.
    quote_count = 0
    for p in doc.paragraphs:
        if re.search(r"Tarifname sayfa\s+\d+,\s*satır\s+\d+-\d+’te", p.text):
            quote_count += 1
            if "“" not in p.text or "”" not in p.text:
                raise ValueError("Görüş dayanak kapısı: sayfa/satır atfının yanında tırnak içi birebir pasaj yok.")
            if not any(bool(r.bold) and ("“" in r.text or "”" in r.text or len(r.text.strip()) > 20) for r in p.runs):
                raise ValueError("Görüş dayanak kapısı: tarifname alıntısı kalın biçimde değil.")
    expected_quotes = sum(1 for _ in _iter_quote_objects(opinion))
    if expected_quotes and quote_count < expected_quotes:
        raise ValueError("Görüş dayanak kapısı: bazı birebir tarifname alıntılarında sayfa/satır görünmüyor.")
    # Combined inventive-step section depth / key reasoning chain.
    combined = opinion.get("combined_assessment") or {}
    combined_text = _norm(" ".join(combined.get("paragraphs") or []))
    low = combined_text.casefold()
    for concept in ["teknik", "problem"]:
        if concept not in low:
            raise ValueError(f"Görüş buluş basamağı kapısı: birlikte değerlendirmede `{concept}` eksik.")
    if not ("motivasyon" in low or "yönlendirme" in low):
        raise ValueError("Görüş buluş basamağı kapısı: birleştirme motivasyonu/yönlendirmesi tartışılmamış.")
    if not ("geriye dönük" in low or "hindsight" in low):
        raise ValueError("Görüş buluş basamağı kapısı: geriye dönük değerlendirme riski tartışılmamış.")


def render_gorus_docx_smoke_test(data: bytes) -> int:
    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        docx = td_path / "gorus.docx"
        docx.write_bytes(data)
        proc = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", str(td_path), str(docx)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=120,
        )
        pdf_path = td_path / "gorus.pdf"
        if proc.returncode != 0 or not pdf_path.exists():
            raise ValueError("Görüş render kapısı: Word PDF'e dönüştürülemedi.")
        pdf = fitz.open(pdf_path)
        pages = pdf.page_count
        if pages < 1:
            raise ValueError("Görüş render kapısı: sıfır sayfa üretildi.")
        return pages
