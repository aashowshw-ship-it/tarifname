from pathlib import Path
import pytest

ROOT=Path(__file__).parent
from rules import APP_VERSION, RULESET_VERSION
import app_core


def base_draft():
    return {
        'title':'Radyo Frekansı Duyarlılığı Tabanlı Dijital İkiz Sistemi ve Yöntemi',
        'technical_field':'Buluş, radyo frekansı duyarlılığı tabanlı dijital ikiz sistemi ve yöntemi ile ilgilidir.\n\nBuluş, özellikle kapalı alan radyo ortamlarının izlenmesine yönelik kullanılmaktadır.',
        'prior_art_general_paragraphs':['Mevcut sistemler ağ performansını ölçmektedir ve fiziksel ortamı doğrudan görememektedir.','Fiziksel engeller sinyal düşüşüne ve yanlış alarma neden olabilmektedir.','Kök neden belirsizliği teknik ve operasyonel sonuçlar doğurmaktadır.'],
        'literature_paragraphs':[],
        'short_description_intro':'Buluş teknik çözüm ile ilgilidir.',
        'objectives':['bir teknik amaç sağlamaktır.'],
        'unumbered_invention_definition':'Bir sistem olup, özelliği;',
        'unumbered_invention_features':['bir modem içermesidir.'],
        'figure_descriptions':['Şekil 1, sistemi göstermektedir.'],
        'elements':[{'number':'100','name':'FWA CPE modemi','description':'modem'},{'number':'101','name':'Ev içi dijital ikiz simülatörü','description':'simülatör'}],
        'method_steps':[],
        'detailed_paragraphs':['FWA CPE modemi (100) ile ev içi dijital ikiz simülatörü (101) ilişkilidir.'],
        'formulas':[], 'tables':[], 'experimental_results':[],
        'alternatives':['Buluş depo, stadyum ve hastane ortamlarında uygulanabilmektedir.'],
        'working_principle':'FWA CPE modemi (100) ile ev içi dijital ikiz simülatörü (101) birlikte çalışmaktadır.',
        'system_claim':{'preamble':'elektronik işlem birimi üzerinde koşturulan yazılım vasıtasıyla çalışan bir sistem','elements':['genişbant bağlantıyı sonlandıran FWA CPE modemi (100),','FWA CPE modemi (100) tarafından üretilen veriyi işleyen ev içi dijital ikiz simülatörü (101)'],'closing':'içermesidir.'},
        'dependent_system_claims':[], 'method_claim':None, 'dependent_method_claims':[],
        'abstract':'Buluş, radyo ortamını izleyen bir sistem ile ilgilidir.',
        'source_coverage_map':[{'fact_id':'T001','covered':True,'sections':['ÖNCEKİ TEKNİK'],'evidence':'Mevcut sistemler ağ performansını ölçmektedir ve fiziksel ortamı doğrudan görememektedir.'},{'fact_id':'T002','covered':True,'sections':['ÖNCEKİ TEKNİK'],'evidence':'Fiziksel engeller sinyal düşüşüne ve yanlış alarma neden olabilmektedir.'},{'fact_id':'T003','covered':True,'sections':['ÖNCEKİ TEKNİK'],'evidence':'Kök neden belirsizliği teknik ve operasyonel sonuçlar doğurmaktadır.'},{'fact_id':'T004','covered':True,'sections':['ÖNCEKİ TEKNİK'],'evidence':'Mevcut sistemler ağ performansını ölçmektedir ve fiziksel ortamı doğrudan görememektedir.'}],
        'coverage_audit':{'prior_art_complete':True,'reference_table_complete':True,'claims_consistent':True,'reference_names_clear':True,'reference_order_valid':True,'how_test_passed':True,'core_difference_present':True,'scope_not_overlimited':True,'dependent_claims_non_redundant':True,'dependent_claim_dependencies_valid':True,'example_dimensions_not_claim_limited':True,'product_claim_language_valid':True,'abstract_single_paragraph_sentence':True,'source_attribution_removed':True,'all_technical_facts_covered':True,'software_carrier_valid':True,'detail_intro_sentence_case':True},
    }


def extracted():
    return {'technical_facts':[
        {'id':'T001','category':'önceki_teknik','statement':'a','mandatory':True},
        {'id':'T002','category':'problem','statement':'b','mandatory':True},
        {'id':'T003','category':'problem','statement':'c','mandatory':True},
        {'id':'T004','category':'problem','statement':'d','mandatory':True},
    ]}


def test_versions():
    assert APP_VERSION=='v5.4.43'
    assert RULESET_VERSION=='2026-09-01.v33'


def test_sentence_case_reference_names():
    d=base_draft()
    app_core._validate_turkish_reference_sentence_case(d,'Türkçe')
    d['elements'][1]['name']='Ev İçi Dijital İkiz Simülatörü'
    with pytest.raises(ValueError):
        app_core._validate_turkish_reference_sentence_case(d,'Türkçe')


def test_title_parentheses_rejected():
    d=base_draft(); d['title']='Radyo Frekansı Duyarlılığı (RF Sensing) Sistemi ve Yöntemi'
    with pytest.raises(ValueError):
        app_core._validate_turkish_title_style(d,'Türkçe')


def test_related_alternatives_single_paragraph():
    d=base_draft(); d['alternatives']=['Buluş depoda uygulanır.','Buluş hastanede uygulanır.']
    with pytest.raises(ValueError):
        app_core._validate_related_alternative_paragraphs(d,'Türkçe')


def test_prior_art_facts_must_live_in_prior_art():
    d=base_draft(); e=extracted()
    app_core._validate_prior_art_source_placement(d,e,'Türkçe')
    d['source_coverage_map'][1]['sections']=['BULUŞUN DETAYLI AÇIKLAMASI']
    with pytest.raises(ValueError):
        app_core._validate_prior_art_source_placement(d,e,'Türkçe')


def test_turkish_inline_title_has_no_combining_dot():
    assert app_core._inline_invention_title('Dijital İkiz Sistemi') == 'dijital ikiz sistemi'


def test_house_style_normalizes_title_case_elements_and_merges_alternatives():
    d=base_draft()
    d['elements'][0]['name']='FWA CPE Modemi'
    d['elements'][1]['name']='Ev İçi Dijital İkiz Simülatörü'
    d['detailed_paragraphs']=['FWA CPE Modemi (100), Ev İçi Dijital İkiz Simülatörü (101) ile veri alışverişi yapmaktadır.']
    d['system_claim']['elements']=['FWA CPE Modemi (100),','Ev İçi Dijital İkiz Simülatörü (101)']
    d['alternatives']=['Buluş depoda uygulanabilmektedir.','Buluş hastanede uygulanabilmektedir.']
    out=app_core.apply_tarifname_house_style(d,'Sistem ve yöntem',[],'Türkçe')
    assert [e['name'] for e in out['elements']] == ['FWA CPE modemi','Ev içi dijital ikiz simülatörü']
    body=' '.join(out['detailed_paragraphs'])+' '+ ' '.join(out['system_claim']['elements'])
    assert 'FWA CPE Modemi' not in body
    assert 'Ev İçi Dijital İkiz Simülatörü' not in body
    assert 'ev içi dijital ikiz simülatörü' in body
    assert len(out['alternatives']) == 1
    assert 'depoda' in out['alternatives'][0] and 'hastanede' in out['alternatives'][0]
