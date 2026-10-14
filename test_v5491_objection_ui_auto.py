"""İtiraz veri toplama: resmî belge esas, elle istem/kapsam sayısı yok."""
from pathlib import Path
import pytest
from docx import Document
from belge_itiraz import (extract_claims_from_application, verify_evidence_proposal,
                           missing_evidence_fields)

APP=Path(__file__).with_name('app.py').read_text()

def test_separate_workflow_not_research_update():
    assert 'elif work_type == "Araştırma güncelleme - Tip 3":' in APP
    assert '    # ARAŞTIRMA GÜNCELLEME\n    else:' not in APP
    assert 'if work_type == "Belge kararına karşı itiraz":' in APP

def test_manual_claim_and_count_fields_removed():
    assert 'st.text_area("Bağımsız İstem 1 — özgün başvurudan birebir yapıştır' not in APP
    assert 'st.number_input("Toplam istem sayısı"' not in APP
    assert 'extract_claims_from_application(official_spec.name,official_spec.getvalue())' in APP

def test_sources_separate_and_missing_optional_recognized():
    for key in ['obj_official_spec','obj_official_decision','obj_official_exam',
                'obj_official_search','obj_previous','obj_counter']:
        assert 'key="'+key+'"' in APP
    assert 'obj_evidence_dates' not in APP
    assert "missing_evidence_fields(prep['evidence'])" in APP

def test_pdf_claims_automatically_extracted():
    p=Path('/mnt/data/2025_012373.pdf')
    if not p.exists(): pytest.skip('Fixture patent pdf yok')
    cs=extract_claims_from_application(p.name,p.read_bytes())
    assert len(cs)==5
    assert cs[0].startswith('Bir ışınım sensörü yapılanması olup')
    assert 'fotovoltaik algılama elemanı (2)' in cs[0]
    assert cs[0].endswith('içermesidir.')

def test_docx_claims_automatically_extracted():
    from io import BytesIO
    doc=Document();doc.add_paragraph('İSTEMLER')
    doc.add_paragraph('1. Bir fotovoltaik sensör olup, en az bir sıcaklık ölçer ve işlem birimi içermesidir.')
    doc.add_paragraph('2. İstem 1’e uygun sensör olup, en az bir veri kaydedici modülü içermesidir.')
    doc.add_paragraph('TARİFNAME')
    out=BytesIO();doc.save(out)
    cs=extract_claims_from_application('istem.docx',out.getvalue())
    assert len(cs)==2
    assert cs[0].startswith('Bir fotovoltaik sensör olup')

def test_missing_claim_section_fails():
    from io import BytesIO
    doc=Document();doc.add_paragraph('TARİFNAME')
    out=BytesIO();doc.save(out)
    with pytest.raises(ValueError, match='İSTEMLER'):
        extract_claims_from_application('wrong.docx',out.getvalue())

def test_unproved_metadata_stays_missing():
    txt={'doc.pdf':'2008 yılında yayımlanan TC1047A sıcaklık ölçüm devresi, alçak geçiren filtre.'}
    rows=verify_evidence_proposal([dict(name='doc.pdf',date='2008',publication_proof='',
        source_locator='Bölüm 2.2.2',quoted_text='TC1047A sıcaklık ölçüm devresi, alçak geçiren filtre.',figure_page=1)],txt)
    assert missing_evidence_fields(rows)=={'doc.pdf':['publication_proof']}
    with pytest.raises(ValueError,match='kamuya erişim'):
        verify_evidence_proposal(rows,txt,require_complete=True)

def test_quote_must_be_source_exact():
    with pytest.raises(ValueError,match='birebir'):
        verify_evidence_proposal([dict(name='d.pdf',quoted_text='Bu teknik özellik kaynakta bulunmamaktadır.',
           date='2008',publication_proof='url',source_locator='s. 5')],{'d.pdf':'Başka bir teknik içerik'})

def test_wrong_evidence_name_fails():
    with pytest.raises(ValueError,match='eşleşmiyor'):
        verify_evidence_proposal([{'name':'missing.pdf'}],{'real.pdf':'Metin'})


def test_nonpublic_copyright_is_not_proof():
    txt={'x.pdf':'Microchip 2008 Copyright 2008 Technical notes TC1047A analog low pass filter.'}
    rows=verify_evidence_proposal([dict(name='x.pdf',date='2008',publication_proof='© 2008',
      source_locator='Bölüm 2',quoted_text='TC1047A analog low pass filter.',figure_page=0)],txt)
    assert 'publication_proof' in missing_evidence_fields(rows)['x.pdf']

def test_false_figure_number_is_rejected():
    txt={'x.pdf':'Temperature circuit with Figure 3, thermistor and analog low pass filter.'}
    with pytest.raises(ValueError, match='kaynakta geçmeyen şekil'):
        verify_evidence_proposal([dict(name='x.pdf',date='',publication_proof='',
         source_locator='Figure 19',quoted_text='thermistor and analog low pass filter.',figure_page=2)],txt)


def test_figures_appear_on_exact_selected_pdf_page():
    from belge_itiraz import validate_pdf_figure_page
    assert validate_pdf_figure_page('Bölüm 2.5, Figure 18','Figure 18. Dedicated sensor PCB')
    with pytest.raises(ValueError,match='Seçilen PDF sayfası'):
        validate_pdf_figure_page('Bölüm 2.5, Figure 18','Figure 19. Flex PCB')
