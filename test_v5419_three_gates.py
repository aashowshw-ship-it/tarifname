from copy import deepcopy
from validators import validate_draft
from test_v5418_rules import base_draft


def errors(d):
    return [x.get('message','') for x in validate_draft(d) if x.get('level') == 'Hata']


def test_passive_database_rejected_from_common_software_carrier():
    d = base_draft()
    d['elements'].append({'number':'4','name':'TCP profil veritabanı','description':'Her slice için TCP profillerini tutar.'})
    d['system_claim']['elements'][1]['subelements'].append('TCP profillerini tutan TCP profil veritabanı (4),')
    assert any('pasif/veri taşıyan' in m for m in errors(d))


def test_dependent_system_claim_bulunmasidir_rejected():
    d = base_draft()
    d['dependent_system_claims'] = ["İstem 1’e uygun sistem olup, özelliği; TCP profilinin veritabanında bulunmasıdır."]
    assert any('yanlış eylem/işlem sonuyla' in m or 'olmasıdır' in m for m in errors(d))


def test_missing_reference_in_method_step_is_rejected():
    d = base_draft()
    d['method_steps'] = [{'number':'1001','text':'İnsansız hava aracından alınan verilerin toplanması'}]
    d['method_claim'] = {'preamble':'Bir yöntem','steps':['İnsansız hava aracından alınan verilerin toplanması (1001)'], 'closing':'işlem adımlarını içermesidir.'}
    assert any('referansını taşımıyor' in m for m in errors(d))


def test_reference_present_in_method_step_is_allowed_for_presence_gate():
    d = base_draft()
    d['method_steps'] = [{'number':'1001','text':'İnsansız hava aracından (1) alınan verilerin toplanması'}]
    d['method_claim'] = {'preamble':'Bir yöntem','steps':['İnsansız hava aracından (1) alınan verilerin toplanması (1001)'], 'closing':'işlem adımlarını içermesidir.'}
    msgs = errors(d)
    assert not any('referansını taşımıyor' in m for m in msgs)
