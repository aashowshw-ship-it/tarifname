"""Belge kararına karşı itiraz — kaynak-kilitli, yeniden kullanılabilir işlem.

Bu modül 182885'in teknik içeriğini başka dosyaya taşımaz; yalnız sabit hukuk
metinleri ve biçim örneği uyarlanır. Hukuki aşama ve kanıtlar dosyaya özeldir.
"""
from __future__ import annotations
import io, re
from copy import deepcopy
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

TEMPLATE = Path(__file__).with_name('Belge_Kararina_Itiraz_182885_Ornek.docx')
HEADING_CRITERIA = 'BULUŞUN PATENTLENEBİLİRLİK KRİTERLERİNE GÖRE DEĞERLENDİRMESİ'
HEADING_SCOPE = 'HAKKIN KAPSAMININ BELİRLENMESİ'
HEADING_EVIDENCE = 'SUNULAN DELİLLERİN LİSTESİ VE AÇIKLAMASI'
REQUIRED = ('source_verification','legal_route','chronology','claims_verbatim','evidence_publication',
            'evidence_quotes','claim_matrix','minus_resolution','inventive_step',
            'dependent_claims','independent_second_review','template_fixed_text','word_visual_render')


def check_stage(patent_kind, status, gazette_date, claim_count, evidence, submitted, audit):
    """FAIL-closed: nobody can mark an absent stage completed with a checkbox alone."""
    if patent_kind not in ('Patent','Faydalı model'):
        raise ValueError('Başvuru türü belirlenemedi (FAIL).')
    if status not in ('Belge kararı verildi, yayımlanmadı','Belge kararı yayımlandı','Belge verildi / diğer aşama'):
        raise ValueError('Karar aşaması seçilmedi (FAIL).')
    if patent_kind == 'Faydalı model' and status == 'Belge kararı yayımlandı':
        raise ValueError('Faydalı modelde SMK 99 itiraz yolu uygulanamaz; hukuki yol yeniden belirlenmelidir (FAIL).')
    if gazette_date and not re.fullmatch(r'\d{2}\.\d{2}\.\d{4}',gazette_date):
        raise ValueError('Bülten tarihi biçimi hatalı (FAIL).')
    if not isinstance(claim_count,int) or claim_count < 1:
        raise ValueError('İstem sayısı ve istem kaynağı gerekli (FAIL).')
    if not evidence or not isinstance(evidence,list):
        raise ValueError('En az bir tarihli gerçek delil eksik (FAIL).')
    if not submitted or not isinstance(submitted,dict):
        raise ValueError('Delil kaynakları ve istem belgeleri eksik (FAIL).')
    if any(not isinstance(x,dict) or not x.get('name') or not x.get('publication_proof') or
           not x.get('quoted_text') or not x.get('source_locator') for x in evidence):
        raise ValueError('Delil yayın/kamuya erişim, aynen alıntı veya konum doğrulaması eksik (FAIL).')
    missing = [key for key in REQUIRED if audit.get(key) is not True]
    if missing:
        raise ValueError('Belge itiraz teslim kapısı FAIL/NOT_RUN: '+', '.join(missing))
    return True


def validate_analysis(result, evidence, claim_text, kind):
    """Structure and self-consistency only; separate source audit remains mandatory."""
    if not result or not isinstance(result,dict): raise ValueError('Analiz boş (FAIL).')
    if result.get('claim_1_verbatim','').strip()!=claim_text.strip():
        raise ValueError('İstem 1 aynen aktarılmadı (FAIL).')
    matrix=result.get('claim_matrix')
    if not isinstance(matrix,list) or len(matrix)<4:
        raise ValueError('İstem özellik matrisi eksik (FAIL).')
    joined=' '.join(x.get('feature','') for x in matrix)
    for x in matrix:
        if x.get('feature','').strip() not in claim_text:
            raise ValueError('İstem unsur tablosu özgün istemin birebir kelimelerini içermiyor (FAIL).')
    if any(x.get('status') not in ('+','−') for x in matrix):
        raise ValueError('İstem tablosu +/- belirsiz (FAIL).')
    for x in matrix:
        if not x.get('feature') or not x.get('d1_basis') or not x.get('legal_consequence'):
            raise ValueError('İstem unsuru veya karşılaştırma gerekçesi eksik (FAIL).')
        if x.get('status')=='−' and not x.get('other_evidence'):
            raise ValueError('D1 eksi unsurunun delil karşılığı eksik (FAIL).')
    if not all(result.get(k) for k in ('chronology','examiner_rebuttal','objective_technical_problem',
                                       'motivation_to_combine','overall_conclusion','evidence_sections')):
        raise ValueError('Kronoloji, uzman gerekçesi veya problem-çözüm analizi eksik (FAIL).')
    listed={x['name'] for x in evidence}
    if set(result['evidence_sections']) != listed:
        raise ValueError('Delil açıklamaları listesi kaynak dosyalarla uyuşmuyor (FAIL).')
    for name,section in result['evidence_sections'].items():
        if len(str(section).split())<230:
            raise ValueError(f'{name} analizi çok kısa (FAIL): en az 230 sözcük.')
    for term in ('müşteri incelemesi','ön taslak','bülten yayımlanınca','dayanak:', 'basılı sayfa', 'pdf sayfası'):
        if term in str(result).lower():
            raise ValueError('Dilekçe içinde müşteri/üretim içi veya yasak kaynak gösterim ifadesi: '+term)
    return True


def _copy_fixed_paragraphs(source):
    paras=list(source.paragraphs)
    heads=[p.text.strip() for p in paras]
    a=heads.index(HEADING_CRITERIA)
    b=heads.index(HEADING_SCOPE)
    c=heads.index(HEADING_EVIDENCE)
    return [deepcopy(p._p) for p in paras[a:c] if p.text.strip() or (p._p is not None)]


def make_docx(reference,app_no,holder,opponent,date_text,analysis,evidence,claim_text,template_path=TEMPLATE, figure_bytes=None, chronology_images=None):
    """New document from reference style, with fixed legal text copied verbatim.

    Never blindly reuses 182885's chronology, merits or party-specific tables.
    """
    source=Document(str(template_path))
    fixed=_copy_fixed_paragraphs(source)
    doc=Document(str(template_path))
    body=doc._element.body
    for el in list(body):
        if el.tag != qn('w:sectPr'): body.remove(el)
    style=doc.styles['Normal']; style.font.name='Arial'; style.font.size=Pt(11)
    style.paragraph_format.space_after=Pt(9)
    def p(text='',bold=False):
        q=doc.add_paragraph();r=q.add_run(str(text));r.bold=bold;return q
    p('TÜRK PATENT VE MARKA KURUMU',True).alignment=WD_ALIGN_PARAGRAPH.CENTER
    p('YENİDEN İNCELEME VE DEĞERLENDİRME DAİRESİ BAŞKANLIĞI’NA',True).alignment=WD_ALIGN_PARAGRAPH.CENTER
    p(date_text).alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p('BAŞVURU NO: '+app_no,True);p('BAŞVURU SAHİBİ: '+holder,True)
    p('İTİRAZ SAHİBİ: '+opponent,True);p('Sayın Yetkili,')
    p(f'{app_no} sayılı patent belge kararına karşı, istemlerde tanımlanan teknik özelliklerin önceki teknik karşısında buluş basamağı içermediği gerekçeleri aşağıda sunulmaktadır.')
    p('İTİRAZ BİLDİRİLEN BAŞVURUNUN KRONOLOJİSİ',True)
    for line in analysis['chronology']:p(line)
    for name, raw in (chronology_images or []):
        q=doc.add_paragraph();q.add_run().add_picture(io.BytesIO(raw),width=Cm(15.5));p('TÜRKPATENT kayıt görüntüsü: '+name)
    # Fixed official legal text from the example, including inline styles, not paraphrased.
    for el in fixed: body.insert(len(body)-1 if body[-1].tag == qn('w:sectPr') else len(body),deepcopy(el))
    p(HEADING_EVIDENCE,True)
    table=doc.add_table(rows=1, cols=3);table.style='Table Grid'
    for c,t in zip(table.rows[0].cells,['Delil No','Tarih','Delil Açıklaması']):c.text=t
    for idx,e in enumerate(evidence,1):
        cells=table.add_row().cells;cells[0].text=f'Ek-{idx}';cells[1].text=e.get('date','');cells[2].text=e['name']
    p('DELİLLER VE HAKKIN KAPSAMI IŞIĞINDA BAŞVURUNUN DEĞERLENDİRİLMESİ',True)
    p('İTİRAZ KONUSU BAĞIMSIZ İSTEM',True);p(claim_text)
    p('İSTEM 1 İÇİN D1 UNSUR KARŞILAŞTIRMA TABLOSU',True)
    m=doc.add_table(rows=1,cols=4);m.style='Table Grid'
    for c,t in zip(m.rows[0].cells,['İstem 1 teknik unsurunun aynen alıntısı','D1','D1 dayanağı / ek delil','Buluş basamağı itirazı']):c.text=t
    for entry in analysis['claim_matrix']:
        row=m.add_row();texts=[entry['feature'],entry['status'],entry['d1_basis']+'\n'+str(entry.get('other_evidence','')),entry['legal_consequence']]
        for c,t in zip(row.cells,texts):c.text=t
    p('İNCELEME RAPORU GEREKÇELERİNE KARŞI İTİRAZLAR',True)
    for t in analysis['examiner_rebuttal']:p(t)
    for idx,e in enumerate(evidence,1):
        p(f'EK-{idx} — {e["name"].upper()}',True)
        p('Özgün kaynak bölümü: '+e['source_locator'])
        p('“'+e['quoted_text']+'”')
        p(analysis['evidence_sections'][e['name']])
        if figure_bytes and e['name'] in figure_bytes:
            q=doc.add_paragraph();q.add_run().add_picture(io.BytesIO(figure_bytes[e['name']]),width=Cm(15))
            p(f'Ek-{idx}: {e["name"]} — özgün PDF şekil/sayfa görüntüsü')
    p('OBJEKTİF TEKNİK PROBLEM VE BULUŞ BASAMAĞI',True)
    p(analysis['objective_technical_problem']);p(analysis['motivation_to_combine'])
    for t in analysis.get('dependent_claims',[]):p(t)
    p('NETİCE VE TALEP',True);p(analysis['overall_conclusion'])
    b=io.BytesIO();doc.save(b);return b.getvalue()
