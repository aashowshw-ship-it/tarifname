from copy import deepcopy

from docx import Document
from docx.oxml.ns import qn

from app_core import ARASTIRMA_TEMPLATE, _apply_tip3_page2_layout
from rules import APP_VERSION, RULESET_VERSION


def test_version_is_v5453_rev2():
    assert APP_VERSION == "v5.4.59"
    assert RULESET_VERSION == "2026-09-11.v52"


def test_tip3_page2_layout_is_inline_and_auto_height():
    doc = Document(str(ARASTIRMA_TEMPLATE))
    _apply_tip3_page2_layout(doc)
    for table_index in (0, 1):
        assert doc.tables[table_index]._tbl.tblPr.find(qn("w:tblpPr")) is None
    for row_index in (3, 4):
        tr_pr = doc.tables[0].rows[row_index]._tr.trPr
        assert tr_pr is None or tr_pr.find(qn("w:trHeight")) is None


def test_tip3_keyword_and_ipc_rows_do_not_have_fixed_height():
    doc = Document(str(ARASTIRMA_TEMPLATE))
    _apply_tip3_page2_layout(doc)
    for row_index in (3, 4):
        tr_pr = doc.tables[0].rows[row_index]._tr.trPr
        assert tr_pr is None or tr_pr.find(qn("w:trHeight")) is None
