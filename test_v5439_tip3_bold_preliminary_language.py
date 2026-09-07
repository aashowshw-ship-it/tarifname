from __future__ import annotations

import io

import pytest
from docx import Document
from PIL import Image

from app_core import build_research_docx, validate_research_report_language


def _png_bytes() -> bytes:
    im = Image.new('L', (900, 500), 255)
    out = io.BytesIO()
    im.save(out, format='PNG')
    return out.getvalue()


def _report() -> dict:
    common_rows = [
        {"feature": "çoklu veri kaynağından risk değerlendirmesi", "status_evidence": "+ Abstract, Claim 1"},
        {"feature": "olay kaydının oluşturulması", "status_evidence": "- Açık ve doğrudan açıklama bulunmamaktadır"},
    ]
    return {
        "reference": "182009",
        "title": "Araç İçi Risk Değerlendirme Sistemi",
        "report_date": "31.08.2026",
        "purpose": "Belirlenen konuda araştırmanın gerçekleştirilmesi",
        "scope": "Global (İlan edilmiş olan patent başvuruları)",
        "keywords": ["vehicle risk assessment", "driver monitoring", "vehicle sensor fusion", "emergency detection", "event logging", "driver safety", "risk scoring", "vehicle alert system", "sensor based monitoring", "automotive safety analytics"],
        "ipc_cpc": [
            {"code": "B60R 25/00", "description": "Fittings or systems for preventing or indicating unauthorised use or theft of vehicles"},
        ],
        "evaluation_intro": "Bu alan Word üreticisi tarafından deterministik oluşturulur.",
        "documents": [
            {
                "label": "D1", "number": "CN120088929A", "alternate_number": "", "title": "Driver emergency calling method", "date": "03.06.2025", "source_url": "", "figure_reference": "Fig. 1", "figure_image_url": "",
                "description": ["D1 dokümanı araç içi acil durum tespiti ile ilgilidir. Dokümanda sensör verilerine dayalı bir acil durum değerlendirmesi açıklanmaktadır."],
                "abstract": "A vehicle emergency method is disclosed for detecting an emergency and sending information.",
                "figure_caption": "D1- Şekil 1",
                "comparison_rows": common_rows,
                "novelty_assessment": ["Bu nedenle araştırma konusu, D1 dokümanı karşısında bütün teknik özellikleri ve aralarındaki ilişki bakımından doğrudan ve açık biçimde açıklanmadığından yenilik kriterini sağlamaktadır."],
            },
            {
                "label": "D2", "number": "GB2608795A", "alternate_number": "", "title": "Vehicle monitoring system", "date": "18.01.2023", "source_url": "", "figure_reference": "Fig. 1", "figure_image_url": "",
                "description": ["D2 dokümanı araç izleme sistemi ile ilgilidir. Dokümanda sensörler ve uyarı düzeni açıklanmaktadır."],
                "abstract": "A vehicle monitoring system includes sensors and an alerting arrangement.",
                "figure_caption": "D2- Şekil 1",
                "comparison_rows": common_rows,
                "novelty_assessment": ["D2 teknik olarak yakın bir araç izleme yapısı açıklamaktadır. D2 dokümanında olay kaydının oluşturulması ile ilgili bir emareye rastlanmamıştır. Bu kapsamda araştırma konusu buluşun D2 dokümanı varlığında yeni olduğu düşünülmektedir."],
            },
        ],
        "inventive_step_paragraphs": [
            "D1, araç içi sensör verilerinden acil durum tespiti yapılmasını temel alan en yakın başlangıç noktasını oluşturmaktadır. Araştırma konusu ise bu yapıya çoklu veri kaynaklarının birlikte değerlendirilmesi ve olay kaydı oluşturulması yönlerini eklemektedir. Bu farklar, sürüş sırasında risk durumunun yalnız tekil bir alarm olarak değil, izlenebilir bir olay dizisi içinde değerlendirilmesi teknik problemini ortaya koymaktadır.",
            "D2, araç izleme sensörleri ve uyarı düzeni bakımından D1'i tamamlayan bir öğretim sunmaktadır. İlgili alandaki uzman kişi D1'in acil durum tespitini geliştirmek istediğinde D2'deki izleme ve uyarı öğretisini dikkate alabilir. Bununla birlikte D2'nin olay kaydını ve çoklu kaynak risk birleştirmesini açıkça öğretip öğretmediği kombinasyonun sonucunu belirleyen esas noktadır.",
            "Kalan özellikler, bilinen sensör, haberleşme ve kayıt bileşenlerinin olağan kullanımının ötesinde teknik bir etki oluşturup oluşturmadığı yönünden ele alınmalıdır. Özelliklerin yalnız yan yana getirilmesi yeterli olmayıp aralarındaki işlevsel ilişkinin beklenmeyen veya sinerjik bir sonuç üretmesi gerekir. Mevcut değerlendirmede bu ilişkinin araştırma konusu buluşun buluş basamağı kriterini sağladığı yönündeki ön kanaati desteklediği düşünülmektedir."
        ],
        "conclusion_paragraphs": ["Araştırma konusu yenilik kriterini sağladığı ve buluş basamağı kriterini sağladığı düşünülmektedir."],
        "warnings": ["Patent başvurusu yapılmasına karar verildiği taktirde:", "Tarifname yazımı için sistem unsurları arasındaki veri akışının netleştirilmesini rica ederiz."],
        "attachments": ["Benzer Dokümanlar", "Ön İnceleme Raporu", "Makine Tercümeleri"],
    }


def test_direct_language_gate_rejects_categorical_patentability_statement():
    report = _report()
    with pytest.raises(ValueError, match="ön araştırma raporudur"):
        validate_research_report_language(report)


def test_build_softens_result_language_and_bolds_d1_d2_identity():
    report = _report()
    data = build_research_docx(report, figure_fallbacks=[_png_bytes(), _png_bytes()])
    doc = Document(io.BytesIO(data))

    intro = doc.paragraphs[36]
    bold_text = "".join(r.text for r in intro.runs if r.bold is True)
    assert bold_text == "CN120088929A (D1) ve GB2608795A (D2)"
    assert "teknik yakınlığı en yüksek dokümanlar" in intro.text

    d1_assessment = doc.paragraphs[55].text
    assert "yenilik kriterini sağladığı düşünülmektedir" in d1_assessment
    assert "yenilik kriterini sağlamaktadır" not in d1_assessment
