import hashlib
import json
import re
import pytest
import app_core
from source_guards import validate_final_raw_source_audit
from rules import APP_VERSION, RULESET_VERSION, TARIFNAME_RULES, EXTRA_CONTROL_GATE_KEYS


def _fp(reg):
    payload = "\n".join(
        f"{r['passage_id']}|{r['source']}|{re.sub(r'\s+', ' ', r['text']).strip()}"
        for r in reg
    )
    return hashlib.sha256(payload.encode()).hexdigest()


def _df(text):
    return hashlib.sha256(re.sub(r"\s+", " ", text).strip().encode()).hexdigest()


def test_versions_and_rules():
    assert APP_VERSION == 'v5.4.52' and RULESET_VERSION == '2026-09-01.v34'
    assert 'independent_raw_second_read' in EXTRA_CONTROL_GATE_KEYS
    assert 'önceki `source_passage_audit`' in TARIFNAME_RULES


def test_all_caps_title_is_lowered_except_real_acronym():
    assert app_core._inline_invention_title(
        'AYARLANABİLİR SPEKTRAL DAĞILIMA SAHİP LED TABANLI SOLAR SİMÜLATÖR SİSTEMİ'
    ) == 'ayarlanabilir spektral dağılıma sahip LED tabanlı solar simülatör sistemi'


def test_element_sentence_start_restored():
    d = {
        'elements': [
            {'number': '1', 'name': 'Solar Spektrum Kafası'},
            {'number': '2', 'name': 'Çok Bantlı LED Dizisi'},
        ],
        'detailed_paragraphs': [
            'solar spektrum kafası (1) çalışır. çok bantlı LED dizisi (2) ışık üretir.'
        ],
        'method_steps': [],
    }
    out = app_core._normalize_turkish_element_case_in_draft(d)
    assert out['elements'][0]['name'] == 'Solar spektrum kafası'
    assert out['detailed_paragraphs'][0].startswith('Solar spektrum kafası')
    assert '. Çok bantlı LED dizisi' in out['detailed_paragraphs'][0]
    app_core._validate_turkish_reference_sentence_case(out, 'Türkçe')


def test_raw_second_read_requires_provenance_and_every_raw_passage():
    reg = [
        {
            'passage_id': 'B0001',
            'source': 'BBF',
            'text': 'LED grupları bağımsız sürülerek ayarlanabilir spektral ışık üretmektedir.',
        },
        {
            'passage_id': 'B0002',
            'source': 'BBF',
            'text': 'Başvuru sahibi bilgileri ve imza alanı bulunmaktadır.',
        },
    ]
    extracted = {
        'source_passage_audit': [
            {'passage_id': 'B0001', 'classification': 'technical', 'fact_ids': ['T001'], 'reason': 'teknik yapı açıklandığı için'},
            {'passage_id': 'B0002', 'classification': 'nontechnical', 'fact_ids': [], 'reason': 'idari alan olduğu için'},
        ],
        'technical_facts': [
            {'id': 'T001', 'statement': 'LED grupları bağımsız sürülür', 'mandatory': True}
        ],
    }
    draft_text = 'Sistem içerisinde LED grupları bağımsız sürülerek ayarlanabilir spektral ışık üretmektedir.'
    nonce = 'abc123'
    audit = {
        'audit_meta': {
            'audit_mode': 'independent_raw_source_second_read_v2',
            'audit_nonce': nonce,
            'source_fingerprint': _fp(reg),
            'draft_fingerprint': _df(draft_text),
            'independent_second_read': True,
            'prior_classification_used': False,
            'source_coverage_map_used': False,
        },
        'passage_checks': [
            {
                'passage_id': 'B0001',
                'classification': 'technical',
                'classification_reason': 'teknik ışık üretim işlevi açıklandığı için',
                'source_quote': 'LED grupları bağımsız sürülerek ayarlanabilir spektral ışık üretmektedir.',
                'covered': True,
                'evidence': ['LED grupları bağımsız sürülerek ayarlanabilir spektral ışık üretmektedir.'],
                'detail_transfer_required': True,
                'detail_evidence': 'LED grupları bağımsız sürülerek ayarlanabilir spektral ışık üretmektedir.',
                'missing_detail': '',
            },
            {
                'passage_id': 'B0002',
                'classification': 'nontechnical',
                'classification_reason': 'yalnız idari başvuru bilgisi içerdiği için',
                'source_quote': 'Başvuru sahibi bilgileri ve imza alanı bulunmaktadır.',
                'covered': True,
                'evidence': [],
                'detail_transfer_required': False,
                'detail_evidence': '',
                'missing_detail': '',
            },
        ],
        'all_pass': True,
    }
    stats = validate_final_raw_source_audit(
        audit, extracted, reg, draft_text, detail_text=draft_text, expected_audit_nonce=nonce
    )
    assert stats['audited_raw_passages'] == 2
    assert stats['audited_technical_passages'] == 1

    bad = json.loads(json.dumps(audit))
    bad['audit_meta']['prior_classification_used'] = True
    with pytest.raises(ValueError, match='bağımsızlık'):
        validate_final_raw_source_audit(
            bad, extracted, reg, draft_text, detail_text=draft_text, expected_audit_nonce=nonce
        )

    missing = json.loads(json.dumps(audit))
    missing['passage_checks'] = missing['passage_checks'][:1]
    with pytest.raises(ValueError, match='eksik pasaj'):
        validate_final_raw_source_audit(
            missing, extracted, reg, draft_text, detail_text=draft_text, expected_audit_nonce=nonce
        )
