from __future__ import annotations

import json

from rules import TARIFNAME_RULES

CORE_RULES = TARIFNAME_RULES


def extraction_prompt(source_text: str) -> str:
    return f"""{CORE_RULES}

Aşağıdaki BBF'yi yalnızca yapılandırılmış veri çıkarmak için incele. Teknik metni yeniden icat etme.
JSON dışında hiçbir şey yazma.

ÇIKTI ŞEMASI:
{{
  "title": "",
  "technical_field": "",
  "prior_art": [""],
  "advantages": [""],
  "elements": [{{"number":"10","name":"","function":""}}],
  "method_steps": [{{"number":"1001","text":""}}],
  "working_principle": [""],
  "keywords": [""],
  "has_method_basis": true,
  "method_basis_reason": "",
  "figures": ["Şekil 1, ... gösterimidir."],
  "uncertainties": [""]
}}

BBF:
---
{source_text}
---
"""


def drafting_prompt(extracted: dict, claim_mode: str, selected_literature: list[dict]) -> str:
    literature = selected_literature or []
    return f"""{CORE_RULES}

Aşağıdaki BBF verilerinden Türk patent tarifnamesi oluştur. Kaynakta olmayan bilgi ekleme.
İstem türü tercihi: {claim_mode}
Her buluş için zorunlu teknik çekirdeği ayrıca analiz et; paralel tekrarları ana istemde kapsayıcı yaz, ayrıntılarını gerekirse tek bağımlı istemde topla.
Onaylanan literatür dokümanları: {json.dumps(literature, ensure_ascii=False)}

JSON dışında hiçbir şey yazma.

ÇIKTI ŞEMASI:
{{
  "title":"",
  "technical_field":"",
  "prior_art_paragraphs":[""],
  "short_description_intro":"",
  "objectives":[""],
  "unnumbered_system_definition":"",
  "unumbered_system_elements":[""],
  "figure_descriptions":[""],
  "elements":[{{"number":"10","name":"","description":""}}],
  "method_steps":[{{"number":"1001","text":""}}],
  "detailed_paragraphs":[""],
  "system_claim":{{"preamble":"","elements":[""],"closing":"içermesidir."}},
  "dependent_system_claims":[""],
  "method_claim":null,
  "dependent_method_claims":[""],
  "abstract":"",
  "quality_notes":[""]
}}

NOTLAR:
- 'unumbered_system_definition' ve 'unumbered_system_elements' REFERANS NUMARALARI bölümünden önce kullanılacağı için (1), (2), (1001) veya benzeri referans işareti içermemeli, tamamen numarasız olmalıdır.
- elements alanında kaynaktaki gerçek referans işareti aynen korunur. Kaynakta UW, UW_F, UW_PL, UW_R, UW_M gibi sembolik referanslar kullanılıyorsa number alanına bu sembol yazılabilir. Kaynakta gerçek unsur referansı olmayan 21-37 gibi şekil içi/geçici numaralar yeni referans numarası olarak üretilmez. Unsur adı teknik anlamı değiştirilmeden referans listesi için yalnızca ilk kelimenin baş harfi büyük olacak cümle düzenine çevrilir, standart teknik kısaltmalar korunur.
- method_steps alanında numara ayrı, metin numarasızdır. Metin mastar isimle biter: toplanması, analiz edilmesi vb.
- method_claim varsa her adım metninin sonunda '(1001)' gibi numara yer alsın ve bağımsız yöntem istemindeki her madde virgül ile bitsin.
- Sistem istemi elemanları bağımsız bir alışveriş listesi gibi yazılmasın. Teknik olarak ilişkili her sonraki unsur, mümkün olduğunda daha önce tanımlanan unsurun çıktısı/girdisi veya bağlantısı üzerinden kurulmalı; kaynakta olmayan yapay ilişki eklenmemelidir.
- İstemlerde İngilizce kısaltmalar ilk kullanımda Türkçesiyle bir kez açıklansın; devamında Türkçe kullanılsın.
- Kullanıcıya sunulan tarifname metninde BBF veya buluş bildirim formu ifadeleri kullanılmasın.
- Sistem ve yöntem birlikteyse başlık “... Sistemi ve Yöntemi” biçiminde olsun.
- “Buluşun bir gerçekleştirilmesinde” yerine “Buluşun bir yapılanmasında” kullanılsın.
- TEKNİK ALAN ilk cümlesi “Buluş, ... ile ilgilidir.” olsun. Teknik alan ya tek bir bütün paragrafta tamamlanmalı ve bu paragraf içinde yeniden “Buluş, özellikle ...” diye başlanmamalı; ya da ek açıklama “Buluş, özellikle ...” ile verilecekse bu ifade yeni ve ayrı bir paragrafta başlamalıdır. Ayrı paragraf gerekiyorsa technical_field içinde iki paragrafı \n\n ile ayır.
- Detaylı açıklamada referanslı unsurlar tek sürekli paragrafta anlatılsın.
- Literatür paragraflarında İngilizce başlık ve Türkçe karşılığı birlikte yazılsın.
- objectives alanındaki her amaç tam cümle yüklemiyle bitsin: “... karşılaştırmaktır.”, “... sağlamaktır.” gibi. “... karşılaştırmak.” veya “... sağlamak.” biçiminde çıplak mastar bırakma.
- Açıklama bölümlerinde noktalı virgülü gereksiz kullanma; virgül veya nokta tercih et. İstemlerdeki standart “olup, özelliği;” kalıbı korunabilir.
- İstemlerde HPU_W, FW_min, UW, UW_F, FE, TE gibi semboller anlamı açıklanmadan çıplak teknik unsur olarak kullanılmasın. Önce “kuru hibrit güç ünitesi ağırlığı (HPU_W)”, “asgari görev yakıtı ağırlığı (FW_min)”, “ilave yakıt tahsisi (UW_F)” gibi açılımı yaz, matematiksel ifadede sembolü koru.

YAPILANDIRILMIŞ BBF VERİSİ:
{json.dumps(extracted, ensure_ascii=False, indent=2)}
"""


def literature_prompt(extracted: dict, max_docs: int) -> str:
    return f"""Aşağıdaki buluş için patent literatüründe benzer doküman adayları araştır.
En fazla {max_docs} aday ver. Sonuçları doğrulanabilir yayın numarası, başlık, tarih, kısa teknik benzerlik ve kaynak URL ile ver.
Doküman uydurma. Doğrulanamayan numarayı ekleme. Bu aşamada tarifnameye metin yazma.
JSON dışında hiçbir şey yazma.

ŞEMA:
{{"documents":[{{"publication_number":"","title_en":"","title_tr":"","publication_date":"","similarity":"","url":""}}]}}

BULUŞ VERİSİ:
{json.dumps(extracted, ensure_ascii=False, indent=2)}
"""
