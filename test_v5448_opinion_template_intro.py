from rules import APP_VERSION, RULESET_VERSION, GORUS_RULES
from gorus_audit import validate_opinion_narrative_rules
from app_core import gorus_prompt, gorus_quality_audit_prompt

REPORT = '''T.C.\nTÜRK PATENT VE MARKA KURUMU\nARAŞTIRMA RAPORU\nBaşvuru Numarası: 2026/009893\n22.06.2026\nX D1 istem 1-2\n'''
SPEC = 'İSTEMLER\n1. Test istemi\n2. Test istemi'


def _opinion(intro: str):
    return {
        'intro': intro,
        'sections': [],
        'cited_documents': [],
        'combined_assessment': {'heading':'','paragraphs':[]},
        'conclusion': [],
        'signoff': 'Saygılarımızla,\nDESTEK PATENT A.Ş.',
    }


def test_version_bumped():
    assert APP_VERSION == 'v5.4.51'
    assert RULESET_VERSION == '2026-09-04.v41'


def test_rules_bind_template_intro_and_forbid_institutional_lead():
    assert 'Türk Patent ve Marka Kurumu tarafından' in GORUS_RULES
    assert 'BAŞLAMAZ' in GORUS_RULES
    assert 'Başvuru sahibinin görüşleri aşağıda dikkatinize sunulmaktadır.' in GORUS_RULES
    assert 'gösterilen benzer dokümanlar aşağıdadır:' in GORUS_RULES


def test_turkish_research_intro_accepts_binding_template_shape():
    intro = ('22.06.2026 tarihli araştırma raporunda, 1-2 numaralı istemlerin D1 dokümanı varlığında ayrı ayrı yenilik ve buluş basamağı kriterlerini sağlamadığı belirtilmiştir. '
             'Başvuru sahibinin görüşleri aşağıda dikkatinize sunulmaktadır. '
             'Araştırma raporunda 1-2 numaralı istemler bakımından gösterilen benzer dokümanlar aşağıdadır:')
    try:
        validate_opinion_narrative_rules(_opinion(intro), REPORT, SPEC)
    except ValueError as e:
        assert 'Görüş giriş kapısı' not in str(e), str(e)


def test_turkish_research_intro_rejects_institutional_free_lead():
    intro = ('Türk Patent ve Marka Kurumu tarafından 22.06.2026 tarihinde düzenlenen araştırma raporunda 1-2 numaralı istemlerin yenilik kriterini sağlamadığı belirtilmiştir. '
             'Başvuru sahibinin görüşleri aşağıda dikkatinize sunulmaktadır. Araştırma raporunda 1-2 numaralı istemler bakımından gösterilen benzer dokümanlar aşağıdadır:')
    try:
        validate_opinion_narrative_rules(_opinion(intro), REPORT, SPEC)
    except ValueError as e:
        assert 'Türk Patent ve Marka Kurumu tarafından' in str(e)
    else:
        raise AssertionError('institutional free lead should fail')


def test_generation_and_audit_prompts_include_template_intro_gate():
    p = gorus_prompt(REPORT, SPEC, '', '', '', {}, 'Türkiye araştırma raporu', '180688', 'revizyon yok', 'Türkçe', '')
    assert 'Şablon girişine birebir yapısal sadakat' in p
    assert 'Türk Patent ve Marka Kurumu tarafından' in p
    q = gorus_quality_audit_prompt(REPORT, SPEC, '', '', '', {}, _opinion(''))
    assert 'template_intro_fidelity' in q
