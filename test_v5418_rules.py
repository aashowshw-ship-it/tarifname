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
