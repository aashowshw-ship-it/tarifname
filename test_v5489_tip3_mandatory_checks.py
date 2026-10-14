"""Negative-path tests: missing evidence, superficial PSA, grammatical errors all block."""
import copy
import pytest
from docx import Document
from tip3_hard_gates import (
    check_turkish_forms, check_problem_solution, validate_second_review,
    seal_four_rounds, require_four_rounds, check_docx_tip3,
)
from rules import FINAL_COMPLIANCE_REQUIRED_CHECKS
from app_core import build_and_gate_tip3_report, tip3_independent_second_review_prompt


def sample_report(d2=True):
    fragments = {
        'd1_distinguishing_features':'D1 karşısında farklılık olarak EVC donanımında Linux temel doğrulaması ve gerçek zamanlı emniyet testlerinin ayrıştırılması görülmektedir',
        'technical_effect':'teknik etki olarak donanım arızasının gerçek zamanlı yazılım arızasından ayrıştırılması ve gereksiz RTOS testinin azaltılması gözlenmektedir',
        'objective_technical_problem':'objektif teknik problem, emniyet kritik bir test düzeninde donanım kaynaklı arızaların zamanlama doğrulamasıyla karışmasının önlenmesidir',
        'd2_teaching':'D2 dokümanında ise alt seviye donanım testi ile farklı işletim sistemi ortamlarında sonraki testlerin otomatik düzenlenmesi öğretilmektedir',
        'combination_motivation':'D1 ile D2 birleşimine teknik motivasyon test çevrimini yönetmek ve farklı platformlarda ayrı test türlerini kullanmak olabilir',
        'remaining_features':'kalan özellikler emniyet kritik düğüm senkronizasyonu ile watchdog ve deterministik giriş çıkış zamanlamasının birlikte değerlendirilmesidir',
        'inventive_step_conclusion':'buluş basamağı kriterini sağlamadığı düşünülmektedir çünkü teknik kombinasyonun belirtilen somut test işlevleriyle uyarlanması öngörülebilirdir',
    }
    pieces = [fragments[k] for k in fragments if d2 or k not in ('d2_teaching','combination_motivation')]
    conclusion = '. '.join(pieces) + '. '
    return {'conclusion_paragraphs':[conclusion], 'documents':[{},{}] if d2 else [{}], 'problem_solution':fragments}


def fake_rounds():
    recall = {'candidate_documents':[{'publication_number':f'CN{i}A'} for i in range(25)],
              'search_axes':[f'axis {i} patents' for i in range(9)], 'technical_problem':'P',
              'technical_features':['A'],'technical_effects':['B']}
    gap={'gap_targets':[{'feature_or_relation':'x','why_uncovered':'absent','search_queries':['x search']}],
         'gap_search_note':'six queries run'}
    neighbour={'provisional_shortlist':[{'publication_number':f'CN{i}A'} for i in range(8)],
               'provisional_d1':'CN1A', 'family_neighbour_note':'IPC/FAMILY trace'}
    final={'documents':[{'publication_number':f'CN{i}A'} for i in range(10)],
           'proposed_d1':'CN1A','proposed_d2':'CN2A',
           'research_quality_audit':{'challenge_note':'challenged D1 and D2 alternative'}}
    return recall,gap,neighbour,final


def second_review():
    return {**{k: {'pass':True,'note':'D1/D2 özgün teknik öğretisi ve kaynak dayanağı yeniden değerlendirildi.'}
               for k in ('source_support','d1_closest_prior_art','distinguishing_features','objective_problem_no_solution_hindsight',
                         'd2_motivation','remaining_features_and_effect','result_consistency','language_and_terminology')},
            'unresolved_issues':[], 'overall_pass':True}


def test_psa_requires_real_conclusion_phrases_and_a_distinct_objective_problem():
    report=sample_report()
    check_problem_solution(report)
    bad=copy.deepcopy(report)
    bad['problem_solution']['objective_technical_problem']='objektif teknik problem bir veya daha fazla test ortamını kullanmaktır'
    with pytest.raises(ValueError,match='objective_technical_problem'):
        check_problem_solution(bad)
    bad=copy.deepcopy(report)
    bad['conclusion_paragraphs']=['D1 ve D2 bilinir, dolayısıyla buluş basamağı yoktur.']
    with pytest.raises(ValueError,match='problem-çözüm'):
        check_problem_solution(bad)


def test_single_d1_does_not_need_fake_d2():
    report=sample_report(d2=False)
    check_problem_solution(report)

@pytest.mark.parametrize('bad', ['testinı', 'testinın','testının','testya','edilmesinı'])
def test_language_denies_broken_turkish_in_final_report(bad):
    report=sample_report()
    report['conclusion_paragraphs'] += [f'Bu {bad} ölçümünü.']
    with pytest.raises(ValueError,match='dil kapısı'):
        check_turkish_forms(report)


def test_self_declared_rounds_without_actual_rounds_fail():
    with pytest.raises(ValueError,match='dört tur'):
        require_four_rounds({'research_quality_audit':{'global_recall_completed':True,'gap_search_completed':True,
             'family_neighbour_search_completed':True,'d1_challenge_completed':True}})


def test_four_round_outputs_bind_first_ten_and_d1d2():
    r,g,n,f=fake_rounds()
    f['_research_round_evidence']=seal_four_rounds(r,g,n,f)
    require_four_rounds(f)
    b=copy.deepcopy(f);b['documents'][0]['publication_number']='US999B2'
    with pytest.raises(ValueError,match='değişmiş'):
        require_four_rounds(b)

@pytest.mark.parametrize('phase', ['recall','gap','neighbour','challenge'])
def test_phase_evidence_cannot_be_issued_when_a_phase_is_empty(phase):
    r,g,n,f=fake_rounds()
    if phase=='recall':r['search_axes']=[]
    elif phase=='gap':g['gap_targets']=[]
    elif phase=='neighbour':n['family_neighbour_note']=''
    else:f['research_quality_audit']['challenge_note']=''
    with pytest.raises(ValueError,match='FAIL'):
        seal_four_rounds(r,g,n,f)


def test_separate_second_review_requires_all_detailed_pass():
    audit=second_review()
    validate_second_review(audit)
    bad=copy.deepcopy(audit);bad['objective_problem_no_solution_hindsight']['pass']=False
    with pytest.raises(ValueError,match='bağımsız ikinci uzman'):
        validate_second_review(bad)
    bad=copy.deepcopy(audit);bad['unresolved_issues']=['source unclear']
    with pytest.raises(ValueError,match='çözülmemiş'):
        validate_second_review(bad)


def test_final_release_requires_all_hard_gates_for_both_tip3_flows():
    for kind in ('tip3','tip3_update'):
        required=FINAL_COMPLIANCE_REQUIRED_CHECKS[kind]
        for key in ('research_4_rounds','problem_solution','turkish_language','independent_second_review'):
            assert key in required


def test_builder_rejects_without_search_evidence_or_second_audit():
    with pytest.raises(ValueError,match='dört tur'):
        build_and_gate_tip3_report(sample_report(),{}, {},research_record=None,second_review=None)
    r,g,n,f=fake_rounds(); f['_research_round_evidence']=seal_four_rounds(r,g,n,f)
    with pytest.raises(ValueError,match='ikinci uzman'):
        build_and_gate_tip3_report(sample_report(),{}, {},research_record=f,second_review=None)


def test_prompt_has_real_independent_second_review():
    prompt=tip3_independent_second_review_prompt('bbf', sample_report(), {'d1':{}},'Otomatik belirle',source_documents='özgün patent metni ' * 35)
    assert 'BAĞIMSIZ' in prompt and 'objective_problem_no_solution_hindsight' in prompt


def test_round_lookalike_hashes_cannot_pass_without_raw_outputs():
    r,g,n,f=fake_rounds()
    f['_research_round_evidence']=seal_four_rounds(r,g,n,f)
    clone=copy.deepcopy(f)
    del clone['_research_round_evidence']['raw_rounds']
    with pytest.raises(ValueError,match='ham tur'):
        require_four_rounds(clone)
    clone=copy.deepcopy(f)
    clone['_research_round_evidence']['raw_rounds']['GAP_SEARCH']['gap_targets'][0]['search_queries']=['rewritten']
    with pytest.raises(ValueError,match='değiştirilmiş'):
        require_four_rounds(clone)
