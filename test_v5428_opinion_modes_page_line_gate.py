from pathlib import Path
from rules import APP_VERSION, RULESET_VERSION
from gorus_audit import build_gorus_quality_report

ROOT=Path(__file__).resolve().parent

def test_version_and_readme_are_synced():
    assert APP_VERSION == "v5.4.38"
    assert RULESET_VERSION == "2026-08-27.v28"
    assert (ROOT/'README.md').read_text(encoding='utf-8').startswith('# Patent Atölyesi v5.4.38')

def test_exact_four_opinion_modes_are_visible_in_order():
    src=(ROOT/'app.py').read_text(encoding='utf-8')
    seq=[
        '"Araştırma raporuna karşı"',
        '"İnceleme raporuna karşı"',
        '"EP araştırma raporu veya ofis aksiyon"',
        '"Yurtdışı ofis aksiyon"',
    ]
    positions=[src.index(x, src.index('opinion_modes = [')) for x in seq]
    assert positions == sorted(positions)
    assert 'opinion_modes = [' in src

def test_research_xy_and_office_reasoned_scope_are_separate():
    src=(ROOT/'app.py').read_text(encoding='utf-8')
    assert 'opinion_case_mode == "Araştırma raporuna karşı"' in src
    assert 'detect_ep_xy_documents(report_text_scope) if xy_scope else detect_examiner_reasoned_documents(report_text_scope)' in src
    assert 'allowed_documents=source_state.get("required_docs")' in src

def test_final_markup_is_authority_for_page_line_citations():
    src=(ROOT/'app.py').read_text(encoding='utf-8')
    assert 'st.session_state.gorus_markup_data' in src
    assert 'final_spec_name = "son_markup_tarifname.docx"' in src
    assert 'validate_quote_locations_against_spec' in src
    # clean data must not be selected as the page/line authority in the final citation block
    block=src[src.index('# Sayfa/satır numaraları modelden alınmaz'):src.index('figure_images =', src.index('# Sayfa/satır numaraları modelden alınmaz'))]
    assert 'gorus_clean_data' not in block

def test_quality_report_surfaces_final_markup_page_line_gate():
    names=[x['name'] for x in build_gorus_quality_report()['checks']]
    assert any('SON MARKUP' in x and 'sayfa/satır' in x for x in names)
