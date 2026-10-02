import io
import zipfile
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pytest

from rules import APP_VERSION, RULESET_VERSION, TARIFNAME_DUZENLEME_RULES
from tarifname_update import (
    derive_markup_output_name,
    extract_docx_review_context,
    prepare_review_baseline_docx,
    validate_update_plan,
    build_updated_spec_docx,
    validate_update_result,
    document_text,
)


def _base_docx() -> bytes:
    d = Document()
    d.add_paragraph('SPECIFICATION')
    d.add_paragraph('The codec is a transparent gzip codec used in the response path.')
    d.add_paragraph('CLAIMS')
    d.add_paragraph('A system comprising the transparent gzip codec.')
    b = io.BytesIO(); d.save(b); return b.getvalue()


def _plan(basis_source='existing_spec'):
    quote = 'transparent gzip codec' if basis_source == 'existing_spec' else 'Please generalize gzip to content coding.'
    return {
        'coverage_complete': True,
        'requests': [{
            'id':'R1','customer_request':'Generalize gzip wording','category':'technical_text',
            'decision':'apply','reason':'broader supported terminology','answer_for_customer':'Updated.'
        }],
        'operations': [{
            'request_id':'R1','type':'replace_text','section':'DETAILED DESCRIPTION',
            'locator_text':'The codec is a transparent gzip codec used in the response path.',
            'anchor_text':'','old_text':'gzip','new_text':'content-coding',
            'basis_source':basis_source,'basis_quote':quote,'reason':'minimum terminology revision'
        }],
        'comments': [], 'figure_actions': [], 'blocking_clarifications': [], 'open_procedural_items': []
    }


def test_versions_and_rules():
    assert APP_VERSION == 'v5.4.81'
    assert RULESET_VERSION == '2026-10-02.v73'
    assert 'EN AZ DEĞİŞİKLİK' in TARIFNAME_DUZENLEME_RULES
    assert 'Track Changes' in TARIFNAME_DUZENLEME_RULES
    assert 'Müşteriye gönderilecek mail zorunlu çıktıdır' in TARIFNAME_DUZENLEME_RULES
    assert 'İSTEM AİLESİ BLOK BÜTÜNLÜĞÜ' in TARIFNAME_DUZENLEME_RULES
    assert 'Her `insert_paragraph_after` / `insert_paragraph_before` işlemi TAM OLARAK BİR gerçek Word paragrafı' in TARIFNAME_DUZENLEME_RULES
    assert 'gerçek OMML denklem nesnesine' in TARIFNAME_DUZENLEME_RULES


def test_markup_filename_removes_browser_duplicate_suffix():
    assert derive_markup_output_name('Description_181569(1).docx') == 'Description_181569_markup.docx'
    assert derive_markup_output_name('Description_181569.docx') == 'Description_181569_markup.docx'


def test_minimum_markup_and_format_gate():
    src = _base_docx(); plan = _plan()
    validate_update_plan(plan, src, 'Please generalize gzip to content coding.', 'Henüz başvuru yapılmadı')
    markup = build_updated_spec_docx(src, plan, track_changes=True, add_comments=False)
    clean = build_updated_spec_docx(src, plan, track_changes=False, add_comments=False)
    validate_update_result(src, markup, clean, plan)
    assert 'content-coding codec used in the response path' in document_text(clean)
    with zipfile.ZipFile(io.BytesIO(markup)) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    assert '<w:delText xml:space="preserve">gzip</w:delText>' in xml
    assert '>content-coding</w:t>' in xml
    assert 'transparent gzip codec used in the response path' not in document_text(clean)


def test_existing_customer_track_changes_are_read_then_rejected_for_baseline():
    d = Document(); p = d.add_paragraph(); r = p.add_run('Original ')
    dele = OxmlElement('w:del'); dele.set(qn('w:id'),'1'); dr=OxmlElement('w:r'); dt=OxmlElement('w:delText'); dt.text='word'; dr.append(dt); dele.append(dr); p._p.append(dele)
    ins = OxmlElement('w:ins'); ins.set(qn('w:id'),'2'); ir=OxmlElement('w:r'); it=OxmlElement('w:t'); it.text='replacement'; ir.append(it); ins.append(ir); p._p.append(ins)
    d.add_comment(runs=[r], text='Please review this change', author='Client', initials='C')
    b=io.BytesIO(); d.save(b); raw=b.getvalue()
    ctx=extract_docx_review_context(raw)
    assert 'TRACK DELETE: word' in ctx and 'TRACK INSERT: replacement' in ctx and 'WORD COMMENT:' in ctx
    baseline=prepare_review_baseline_docx(raw)
    assert 'Original word' in document_text(baseline)
    assert 'replacement' not in document_text(baseline)
    with zipfile.ZipFile(io.BytesIO(baseline)) as z:
        assert 'word/comments.xml' not in z.namelist()


def test_post_filing_customer_only_new_matter_is_blocked():
    src=_base_docx(); plan=_plan('customer_request')
    with pytest.raises(ValueError, match='new-matter'):
        validate_update_plan(plan, src, 'Please generalize gzip to content coding.', 'Başvuru yapıldı')


def test_comments_are_written_for_explanation():
    src=_base_docx(); plan=_plan(); plan['comments']=[{
        'request_id':'R1','anchor_text':'The codec is a transparent','text':'This wording was generalized while preserving the existing embodiment.'
    }]
    markup=build_updated_spec_docx(src, plan, track_changes=True, add_comments=True)
    clean=build_updated_spec_docx(src, plan, track_changes=False, add_comments=False)
    validate_update_result(src, markup, clean, plan)
    with zipfile.ZipFile(io.BytesIO(markup)) as z:
        assert 'word/comments.xml' in z.namelist()


def _family_claim_docx() -> bytes:
    d = Document()
    d.add_paragraph('TARİFNAME')
    d.add_paragraph('Mevcut teknik açıklama ve dayanak metni.')
    d.add_paragraph('İSTEMLER')
    d.add_paragraph('1. Bir ses analiz sistemi olup, özelliği; bir analiz birimi içermesidir.')
    d.add_paragraph('2. İstem 1’e uygun sistem olup, özelliği; bir bellek içermesidir.')
    d.add_paragraph('3. Bir ses analiz yöntemi olup, özelliği; analiz işlem adımını içermesidir.')
    d.add_paragraph('4. İstem 3’e uygun yöntem olup, özelliği; sınıflandırma işlem adımını içermesidir.')
    d.add_paragraph('ÖZET')
    b = io.BytesIO(); d.save(b); return b.getvalue()


def test_insert_operation_must_be_one_real_paragraph():
    src = _base_docx()
    plan = _plan()
    plan['operations'] = [{
        'request_id':'R1','type':'insert_paragraph_after','section':'DETAILED DESCRIPTION',
        'locator_text':'','anchor_text':'The codec is a transparent gzip codec used in the response path.',
        'old_text':'','new_text':'Birinci yeni paragraf.\nİkinci yeni paragraf.',
        'basis_source':'customer_request','basis_quote':'Please generalize gzip to content coding.','reason':'test'
    }]
    with pytest.raises(ValueError, match='paragraf-yapısı'):
        validate_update_plan(plan, src, 'Please generalize gzip to content coding.', 'Henüz başvuru yapılmadı')


def test_claim_family_cannot_return_to_system_after_method_block():
    src = _family_claim_docx()
    plan = {
        'coverage_complete': True,
        'requests': [{'id':'R1','customer_request':'Add system fallback','category':'claim_scope','decision':'apply','reason':'fallback','answer_for_customer':'Added.'}],
        'operations': [{
            'request_id':'R1','type':'insert_paragraph_after','section':'İSTEMLER','locator_text':'',
            'anchor_text':'4. İstem 3’e uygun yöntem', 'old_text':'',
            'new_text':'5. İstem 2’ye uygun sistem olup, özelliği; bir doğrulama birimi içermesidir.',
            'basis_source':'customer_request','basis_quote':'Add system fallback','reason':'fallback'
        }],
        'comments': [], 'figure_actions': [], 'blocking_clarifications': [], 'open_procedural_items': []
    }
    validate_update_plan(plan, src, 'Add system fallback', 'Henüz başvuru yapılmadı')
    markup = build_updated_spec_docx(src, plan, track_changes=True)
    clean = build_updated_spec_docx(src, plan, track_changes=False)
    with pytest.raises(ValueError, match='istem-ailesi'):
        validate_update_result(src, markup, clean, plan)


def test_claim_family_inserted_inside_own_block_with_renumbering_passes():
    src = _family_claim_docx()
    plan = {
        'coverage_complete': True,
        'requests': [{'id':'R1','customer_request':'Add system fallback','category':'claim_scope','decision':'apply','reason':'fallback','answer_for_customer':'Added.'}],
        'operations': [
            {
                'request_id':'R1','type':'insert_paragraph_before','section':'İSTEMLER','locator_text':'',
                'anchor_text':'3. Bir ses analiz yöntemi', 'old_text':'',
                'new_text':'3. İstem 2’ye uygun sistem olup, özelliği; bir doğrulama birimi içermesidir.',
                'basis_source':'customer_request','basis_quote':'Add system fallback','reason':'fallback'
            },
            {
                'request_id':'R1','type':'replace_text','section':'İSTEMLER','locator_text':'3. Bir ses analiz yöntemi',
                'anchor_text':'','old_text':'3.','new_text':'4.',
                'basis_source':'existing_spec','basis_quote':'3. Bir ses analiz yöntemi','reason':'renumber'
            },
            {
                'request_id':'R1','type':'replace_text','section':'İSTEMLER','locator_text':'4. İstem 3’e uygun yöntem',
                'anchor_text':'','old_text':'4. İstem 3’e','new_text':'5. İstem 4’e',
                'basis_source':'existing_spec','basis_quote':'4. İstem 3’e uygun yöntem','reason':'renumber dependency'
            }
        ],
        'comments': [], 'figure_actions': [], 'blocking_clarifications': [], 'open_procedural_items': []
    }
    validate_update_plan(plan, src, 'Add system fallback', 'Henüz başvuru yapılmadı')
    markup = build_updated_spec_docx(src, plan, track_changes=True)
    clean = build_updated_spec_docx(src, plan, track_changes=False)
    validate_update_result(src, markup, clean, plan)
    assert '3. İstem 2’ye uygun sistem' in document_text(clean)
    assert '4. Bir ses analiz yöntemi' in document_text(clean)
    assert '5. İstem 4’e uygun yöntem' in document_text(clean)


def test_update_formula_marker_becomes_real_omml():
    src = _base_docx()
    plan = {
        'coverage_complete': True,
        'requests': [{'id':'R1','customer_request':'Add formula C = w EERd + (1-w) EERk','category':'technical_text','decision':'apply','reason':'formula','answer_for_customer':'Added.'}],
        'operations': [{
            'request_id':'R1','type':'insert_paragraph_after','section':'DETAILED DESCRIPTION','locator_text':'',
            'anchor_text':'The codec is a transparent gzip codec used in the response path.', 'old_text':'',
            'new_text':'[[EQ: C = w EER_d + (1 - w) EER_k]]',
            'basis_source':'customer_request','basis_quote':'C = w EERd + (1-w) EERk','reason':'formula'
        }],
        'comments': [], 'figure_actions': [], 'blocking_clarifications': [], 'open_procedural_items': []
    }
    validate_update_plan(plan, src, 'Add formula C = w EERd + (1-w) EERk', 'Henüz başvuru yapılmadı')
    markup = build_updated_spec_docx(src, plan, track_changes=True)
    clean = build_updated_spec_docx(src, plan, track_changes=False)
    validate_update_result(src, markup, clean, plan)
    with zipfile.ZipFile(io.BytesIO(clean)) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    assert '<m:oMath' in xml
    assert '[[EQ:' not in document_text(clean)
