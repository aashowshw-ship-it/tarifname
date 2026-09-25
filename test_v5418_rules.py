from validators import validate_draft


def base_draft():
    return {
        "technical_field":"Buluş, mobil haberleşme ile ilgilidir.\n\nBuluş, özellikle ağ dilimleme ile ilgilidir.",
        "prior_art_paragraphs":[], "short_description_intro":"", "objectives":[],
        "unumbered_system_definition":"", "unumbered_system_elements":[],
        "elements":[
            {"number":"1","name":"İnsansız hava aracı"},
            {"number":"2","name":"İHA parametre toplama modülü"},
            {"number":"3","name":"Ağ parametre izleme modülü"},
        ],
        "method_steps":[], "detailed_paragraphs":["insansız hava aracı (1) ile ilişki kurulmaktadır."],
        "dependent_system_claims":[], "method_claim":None, "dependent_method_claims":[],
        "system_claim":{
            "preamble":"Bir sistem",
            "elements":[
                "insansız hava aracı (1),",
                {"lead":"bir elektronik işlem birimi üzerinde koşturulan yazılım vasıtasıyla çalışan ve;",
                 "subelements":[
                     "insansız hava aracından (1) veri alan İHA parametre toplama modülü (2),",
                     "İHA parametre toplama modülünün (2) çıktısını alan ağ parametre izleme modülü (3),",
                 ]},
            ],
            "closing":"içermesidir."
        },
        "abstract":"Tek cümleli özet metnidir."
    }


def test_grouped_carrier_is_allowed():
    findings=validate_draft(base_draft())
    assert not [x for x in findings if x.get("level")=="Hata" and "ortak taşıyıcı" in x.get("message","").lower()]


def test_reference_alias_is_rejected():
    d=base_draft()
    d["detailed_paragraphs"]=["İHA (1) ile ilişki kurulmaktadır."]
    findings=validate_draft(d)
    assert any(x.get("level")=="Hata" and "Referans (1)" in x.get("message","") for x in findings)

# ---- merged verbatim test logic from test_v5419_three_gates.py ----
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

# ---- merged verbatim test logic from test_v5420_five_gates.py ----
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
