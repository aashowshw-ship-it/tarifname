from copy import deepcopy
from validators import validate_draft
from test_v5418_rules import base_draft


def errors(d):
    return [x.get('message','') for x in validate_draft(d) if x.get('level') == 'Hata']


def test_generic_unsur_in_claim_is_rejected():
    d=base_draft()
    d['dependent_system_claims']=["İstem 1’e uygun sistem olup, özelliği; antenlerin (1), bir unsur olmasıdır."]
    assert any("belirsiz 'unsur'" in m for m in errors(d))


def test_method_step_noun_ending_is_rejected():
    d=base_draft()
    d['method_steps']=[{'number':'1001','text':'İnsansız hava aracının (1) takibi'}]
    d['method_claim']={'preamble':'Bir yöntem','steps':['İnsansız hava aracının (1) takibi (1001)'],'closing':'işlem adımlarını içermesidir.'}
    assert any('gerçek bir işlem fiilimsisiyle bitmiyor' in m for m in errors(d))


def test_method_step_action_ending_is_allowed():
    d=base_draft()
    d['method_steps']=[{'number':'1001','text':'İnsansız hava aracının (1) takibinin yapılması'}]
    d['method_claim']={'preamble':'Bir yöntem','steps':['İnsansız hava aracının (1) takibinin yapılması (1001)'],'closing':'işlem adımlarını içermesidir.'}
    assert not any('gerçek bir işlem fiilimsisiyle bitmiyor' in m for m in errors(d))
