from __future__ import annotations

import json

CORE_RULES = r"""
TÜRK PATENT TARİFNAME KURALLARI
1. Kaynak BBF'ye birebir sadık kal. BBF'de bulunmayan teknik unsur, algoritma, değer, bağlantı veya uygulama ekleme.
2. İngilizce teknik terimi ilk geçtiği yerde Türkçe karşılığıyla bir kez açıkla. Örnek: handover (hücre geçişi). Sonraki kullanımlarda yalnızca Türkçe karşılığı kullan.
3. Unsur adları normal cümle yazımıyla yazılır; her kelime büyük harfle başlamaz. AI yerine yapay zekâ yaz.
4. BBF'deki unsur numaralandırmasını aynen koru. 1,2,3 ise aynı; 10,20,30 ise aynı. Sisteme veya yönteme BBF'de unsur olarak verilmemiş ek numara verme.
5. REFERANS NUMARALARI bölümünden önce parantez içinde referans numarası kullanma.
6. Referans listesinde önce unsurlar, ardından varsa yöntem işlem adımları bulunur. İşlem adımları '1001. ... toplanması' biçiminde yazılır.
7. Detaylı açıklamada ve yöntem isteminde işlem adımı referansı işlem ifadesinin sonunda yer alır: '... toplanması (1001)'. Numara cümlenin başına alınmaz.
8. Sistem isteminde unsurları sırayla kur. Bir unsur tanımlanırken yalnızca daha önce tanımlanmış unsurlarla ilişki kur; henüz tanımlanmamış sonraki unsuru kullanma.
9. Ana istemde unsurlar birbirinden bağımsız liste gibi kalmamalı; veri, sinyal, kontrol veya işlem ilişkileri kurulmalı.
10. Alt istemler ana istemi tekrar etmemeli, kısa olmalı ve sistem alt istemleri 'bir modül olmasıdır' veya 'içermesidir' şeklinde bitmelidir. 'yapmasıdır/etmesidir/belirlemesidir' şeklinde bitirme.
11. İnsan veya soyut aktör yerine teknik araç kullan. Örnek: 'operatöre gönderen' değil, 'elektronik cihaz üzerinden operatöre ileten'.
12. BBF açık, sıralı ve teknik işlem adımları içeriyorsa sistem istemine ek olarak bağımsız yöntem istemi hazırla. BBF yöntem akışını desteklemiyorsa yöntem istemi oluşturma.
13. Yöntem istemindeki işlem adımları BBF'deki sıra, metin ve numaralara mümkün olduğunca birebir sadık olmalı.
14. Şablonun giriş ve İSTEMLER altındaki kırmızı/mavi bilgilendirme paragrafları korunacaktır; bu metinleri çıktı JSON'una ekleme.
15. Literatür araştırması kullanıcı tarafından seçilmediyse doküman uydurma veya önceki tekniğe patent numarası ekleme.
"""


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
- 'unumbered_system_definition' ve 'unumbered_system_elements' REFERANS NUMARALARI bölümünden önce kullanılacağı için numarasız olmalıdır.
- elements alanında BBF unsur adı ve numarası aynen korunur.
- method_steps alanında numara ayrı, metin numarasızdır. Metin mastar isimle biter: toplanması, analiz edilmesi vb.
- method_claim varsa her adım metninin sonunda '(1001)' gibi numara yer alsın.
- Sistem istemi elemanları bir önceki tanımlanmış unsurla teknik ilişki kuracak biçimde sıralansın.
- İstemlerde İngilizce kısaltmalar ilk kullanımda Türkçesiyle bir kez açıklansın; devamında Türkçe kullanılsın.

YAPILANDIRILMIŞ BBF VERİSİ:
{json.dumps(extracted, ensure_ascii=False, indent=2)}
"""


def literature_prompt(extracted: dict, max_docs: int) -> str:
    return f"""Aşağıdaki buluş için patent literatüründe benzer doküman adayları araştır.
En fazla {max_docs} aday ver. Sonuçları doğrulanabilir yayın numarası, başlık, tarih, kısa teknik benzerlik ve kaynak URL ile ver.
Doküman uydurma. Doğrulanamayan numarayı ekleme. Bu aşamada tarifnameye metin yazma.
JSON dışında hiçbir şey yazma.

ŞEMA:
{{"documents":[{{"publication_number":"","title":"","publication_date":"","similarity":"","url":""}}]}}

BULUŞ VERİSİ:
{json.dumps(extracted, ensure_ascii=False, indent=2)}
"""
