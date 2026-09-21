from pathlib import Path
import pytest
from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES
from gorus_audit import _lead_for_span, _quote_semantic_boundary_ok, validate_gorus_docx_content_flow
from app_core import build_gorus_docx


def test_version_and_binding_rules():
    assert APP_VERSION == "v5.4.72"
    assert RULESET_VERSION == "2026-09-18.v64"
    assert "Tarifnamede sayfa X, satır Y-Z’de bu durum şu şekilde belirtilmiştir:" in GORUS_RULES
    assert "Tarifnamedeki dayanak şöyledir:" in GORUS_RULES and "YASAKTIR" in GORUS_RULES
    assert "anlamlı bir cümle" in GORUS_RULES
    assert "en az 2200 kelime" in GORUS_RULES


def test_turkish_lead_is_single_canonical_phrase():
    assert _lead_for_span(12, 18, 12, 31, "Türkçe") == "Tarifnamede sayfa 12, satır 18-31’de bu durum şu şekilde belirtilmiştir:"
    assert _lead_for_span(2, 4, 3, 7, "Türkçe") == "Tarifnamede sayfa 2, satır 4 ile sayfa 3, satır 7 arasında bu durum şu şekilde belirtilmiştir:"


def test_semantic_boundary_rejects_clipped_predicate_and_accepts_sentence_or_bullet():
    idx = [
        {"text":"enerji dengesini optimize edememesi, uzun soluklu görevler için en büyük engeli"},
        {"text":"oluşturmaktadır. Bu buluş, söz konusu teknik problemi çözmektedir."},
        {"text":"• Bahsedilen insansız hava aracı üzerinde tandem güneş hücresi bulunmaktadır."},
    ]
    assert not _quote_semantic_boundary_ok("oluşturmaktadır. Bu buluş, söz konusu teknik problemi çözmektedir.", idx)
    assert _quote_semantic_boundary_ok("Bu buluş, söz konusu teknik problemi çözmektedir.", idx)
    assert _quote_semantic_boundary_ok("Bahsedilen insansız hava aracı üzerinde tandem güneş hücresi bulunmaktadır.", idx)


def test_word_content_flow_rejects_legacy_two_stage_phrase():
    opinion={"application_no":"1","applicant":"A","reference":"R","intro":"01.01.2026 tarihli araştırma raporunda, 1 numaralı istemin D1 varlığında buluş basamağı kriterini sağlamadığı belirtilmiştir. Başvuru sahibinin görüşleri aşağıda dikkatinize sunulmaktadır. Araştırma raporunda 1 numaralı istem bakımından gösterilen benzer dokümanlar aşağıdadır:","cited_documents":[{"label":"D1","number":"X","category":"A"}],"sections":[{"label":"D1","heading":"D1 dokümanı:","use_figure":False,"blocks":[{"type":"paragraph","text":"Teknik açıklama yeterince uzun bir paragraftır ve yalnız test amacıyla oluşturulmuştur."}],"novelty_paragraphs":[],"inventive_step_paragraphs":[]}],"combined_assessment":{},"conclusion":[],"signoff":"Saygılarımızla,\nDESTEK PATENT A.Ş."}
    data=build_gorus_docx(opinion)
    from docx import Document
    import io
    doc=Document(io.BytesIO(data))
    p=doc.add_paragraph("Teknik fark vardır. Tarifnamedeki dayanak şöyledir: Sayfa 12, satır 18-31’de bu durum şu şekilde belirtilmiştir: “örnek”")
    out=io.BytesIO(); doc.save(out)
    with pytest.raises(ValueError, match="dayanak ifade kapısı"):
        validate_gorus_docx_content_flow(out.getvalue())
