from pathlib import Path
from gorus_audit import detect_ep_xy_documents, detect_defense_documents
from rules import APP_VERSION, GORUS_RULES

ROOT=Path(__file__).resolve().parent

def _sample():
    return """
SUPPLEMENTARY EUROPEAN SEARCH REPORT
DOCUMENTS CONSIDERED TO BE RELEVANT
X BANDARA ET AL
XP033966287
-----
X US 2020/145229 A1
-----
A US 2021/273931 A1
-----
A XP047637201
CATEGORY OF CITED DOCUMENTS
D1: XP033966287
D2: US 2020/145229 A1
D3: US 2021/273931 A1
D4: XP047637201
Document D1 discloses all features.
Document D2 is also mentioned.
D3 is known for a dependent feature.
"""

def test_ep_xy_scope_excludes_a_even_if_reasoned():
    got=detect_ep_xy_documents(_sample())
    assert [x['label'] for x in got]==['D1','D2']
    assert [x['label'] for x in detect_defense_documents(_sample())]==['D1','D2']

def test_rules_and_ui_ep_mode_present():
    assert APP_VERSION=='v5.4.46'
    low=GORUS_RULES.casefold()
    assert 'yalnız araştırma raporunda x veya y' in low
    assert 'd1`, `d2` gibi inceleme-etiketleri kullanılmaz' in low
    src=(ROOT/'app.py').read_text(encoding='utf-8')
    assert 'EP araştırma raporu veya ofis aksiyon' in src
    assert 'detect_ep_xy_documents' in src
