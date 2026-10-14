from pathlib import Path
import pytest
from docx import Document
from belge_itiraz import check_stage, validate_analysis, make_docx, REQUIRED, HEADING_CRITERIA, HEADING_SCOPE

CLAIM='Üzerine düşen Güneş ışığını elektriksel sinyale çevirerek gerçek zamanlı ışınım ölçümü yapılmasını sağlayan en az bir fotovoltaik algılama elemanı (2), sıcaklık sensörü (S), elektronik filtre içeren sıcaklık sensörü devre kartı (3) içermesidir.'
FEATURES=[CLAIM[:58],CLAIM[58:100],CLAIM[100:145],CLAIM[145:]]
EVID=[dict(name='2008.pdf',date='2008',publication_proof='2008 yayım',source_locator='Bölüm 2.2.2',quoted_text='sıcaklık sensörü filtresi')]
BASE={'claim_1_verbatim':CLAIM,'chronology':['29.08.2025 başvuru'],
 'claim_matrix':[{'feature':f,'status':'+' if i<2 else '−','d1_basis':'D1 istem 1','other_evidence':'Ek-1','legal_consequence':'aşikardır'} for i,f in enumerate(FEATURES)],
 'examiner_rebuttal':['uzman değerlendirmesine itiraz edilmektedir'],
 'objective_technical_problem':'Mevcut analog ölçüm sinyalinin güvenilirliğinin geliştirilmesidir.',
 'motivation_to_combine':'Kaynakta açıklanan filtre sinyale uygulanmaktadır.',
 'overall_conclusion':'İstemlerin buluş basamağı içermediği açıklanmaktadır.',
 'evidence_sections':{'2008.pdf':'Teknik belgenin içeriği. '*250},'dependent_claims':['İstem 2 buluş basamağı içermemektedir.']}

def test_positive_document_style_and_source_sections(tmp_path):
 assert validate_analysis(BASE,EVID,CLAIM,'Patent')
 b=make_docx('182885','2025/012373','VSİAS','ARGESİM','14.10.2026',BASE,EVID,CLAIM)
 dest=tmp_path/'x.docx';dest.write_bytes(b);d=Document(dest)
 assert any(p.text==HEADING_CRITERIA for p in d.paragraphs)
 assert any(p.text==HEADING_SCOPE for p in d.paragraphs)
 assert d.tables[-1].rows[1].cells[0].text==FEATURES[0]
 assert any("graphic" in header._element.xml for section in d.sections for header in (section.header,section.first_page_header))  # header logo preserved

def test_missing_one_check_fails():
 flags={k:True for k in REQUIRED};flags['word_visual_render']=False
 with pytest.raises(ValueError, match='word_visual_render'):
  check_stage('Patent','Belge kararı verildi, yayımlanmadı','',5,EVID,{'case':'source'},flags)

def test_unpublished_is_valid_preparation_stage():
 flags={k:True for k in REQUIRED}
 assert check_stage('Patent','Belge kararı verildi, yayımlanmadı','',5,EVID,{'case':'source'},flags)

def test_fm_post_grant_route_fails():
 flags={k:True for k in REQUIRED}
 with pytest.raises(ValueError,match='Faydalı model'):
  check_stage('Faydalı model','Belge kararı yayımlandı','',5,EVID,{'case':'source'},flags)

def test_missing_publication_proof_fails():
 flags={k:True for k in REQUIRED}
 with pytest.raises(ValueError,match='yayın'):
  check_stage('Patent','Belge kararı yayımlandı','14.10.2026',5,[dict(EVID[0],publication_proof='')],{'case':'source'},flags)

def test_claim_is_exact():
 a={**BASE,'claim_1_verbatim':CLAIM+' rastgele'}
 with pytest.raises(ValueError,match='aynen'):
  validate_analysis(a,EVID,CLAIM,'Patent')

def test_plus_minus_evidence_missing():
 a={**BASE,'claim_matrix':[dict(x) for x in BASE['claim_matrix']]}
 a['claim_matrix'][2]['other_evidence']=''
 with pytest.raises(ValueError,match='eksi'):
  validate_analysis(a,EVID,CLAIM,'Patent')

def test_claim_feature_not_from_source():
 a={**BASE,'claim_matrix':[dict(x) for x in BASE['claim_matrix']]}
 a['claim_matrix'][0]['feature']='Hayalî bir devre'
 with pytest.raises(ValueError,match='birebir'):
  validate_analysis(a,EVID,CLAIM,'Patent')

def test_unproven_evidence_cannot_be_listed():
 a={**BASE,'evidence_sections':{'2008.pdf':'kısa'}}
 with pytest.raises(ValueError,match='çok kısa'):
  validate_analysis(a,EVID,CLAIM,'Patent')

def test_no_legacy_customer_wording():
 a={**BASE,'overall_conclusion':'Bu belge müşteri incelemesi içindir'}
 with pytest.raises(ValueError,match='müşteri'):
  validate_analysis(a,EVID,CLAIM,'Patent')
