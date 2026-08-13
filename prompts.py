from __future__ import annotations

import json

from rules import TARIFNAME_RULES

CORE_RULES = TARIFNAME_RULES


def extraction_prompt(source_text: str) -> str:
    return f"""{CORE_RULES}

Aşağıdaki BBF'yi yalnızca yapılandırılmış veri çıkarmak için incele. Teknik metni yeniden icat etme. BBF'deki HER teknik bilgiyi atomik `technical_facts` maddelerine ayır; teknik avantajlar ve ayırt edici sonuçlar da ayrı maddeler olmalıdır. Kişi/sicil/ödül/imza, form talimatı, boş idari alan ve yalnız araştırma anahtar kelimelerini `excluded_nontechnical_items` altında ayır.
JSON dışında hiçbir şey yazma.

ÇIKTI ŞEMASI:
{{
  "title": "",
  "technical_field": "",
  "prior_art": [""],
  "advantages": [""],
  "technical_facts": [{{"id":"T001","category":"alan/problem/çözüm/unsur/işlev/akış/avantaj/alternatif/kullanım/ayırt_edici_yön/görsel","statement":"","mandatory":true}}],
  "excluded_nontechnical_items": [""],
  "elements": [{{"number":"1","name":"","function":""}}],
  "method_steps": [{{"number":"1001","text":""}}],
  "working_principle": [""],
  "keywords": [""],
  "has_method_basis": true,
  "method_basis_reason": "",
  "figures": ["Şekil 1, ... gösterimidir."],
  "figure_reference_audit": [{{"figure":"Şekil 1","reference_marks":["1","2"],"method_marks":[],"symbolic_reference_marks":[],"temporary_marks":[],"notes":""}}],
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

Aşağıdaki BBF verilerinden Türk patent tarifnamesi oluştur. Kaynakta olmayan bilgi ekleme. `technical_facts` içindeki HER mandatory teknik bilgi tarifnamede uygun bölümde korunmalı ve `source_coverage_map` ile fact_id bazında kanıtlanmalıdır; isteme uygun olmayan bilgi isteme zorla taşınmamalıdır.
İstem türü tercihi: {claim_mode}
Her buluş için zorunlu teknik çekirdeği ayrıca analiz et; paralel tekrarları ana istemde kapsayıcı yaz, ayrıntılarını gerekirse tek bağımlı istemde topla.
Buluş yazılım/algoritma/modül ağırlıklıysa bağımsız istemi soyut yazılım olarak bırakma; kaynakta özel donanım zorunlu değilse geniş donanımsal taşıyıcı olarak elektronik cihaz üzerinde koşturulan yazılım veya elektronik işlem birimi dili kullan. Kaynak özel bir taşıyıcı veriyorsa bu taşıyıcıyı kaybetme. Yalnız “işlemci/donanım” kelimesi yeterli değildir; modül/yazılımın teknik taşıyıcı üzerinde çalıştığı/koşturulduğu açık ilişki kurulmalıdır. Gereksiz sunucu/cep telefonu/bilgisayar daraltması yapma ve özel donanım uydurma.
Bağımsız istemleri yalnız sonuç/fonksiyon cümleleriyle bırakma. Tekniğin uzmanının “nasıl gerçekleştiriliyor?” sorusuna cevap verecek şekilde, kaynakta dayanağı bulunan ölçüde işlemi yapan teknik unsur/taşıyıcıyı, kullanılan girdiyi veya önceki unsurdan alınan veriyi, teknik işlem/mekanizmayı ve elde edilen çıktının sonraki unsurla ilişkisini istemde açıkla. Ancak tercihli ayrıntılarla ana istemi gereksiz daraltma.
Bağımlı istemleri her kaynak ayrıntısı için çoğaltma; yalnız gerçek teknik daraltma ve stratejik geri çekilme konumu sağlayan seçilmiş özellikleri kullan.
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
  "system_claim":{{"preamble":"","elements":["",{{"lead":"bir elektronik işlem birimi üzerinde koşturulan yazılım vasıtasıyla çalışan ve;","subelements":["","" ]}}],"closing":"içermesidir."}},
  "dependent_system_claims":[""],
  "method_claim":null,
  "dependent_method_claims":[""],
  "abstract":"",
  "source_coverage_map":[{{"fact_id":"T001","covered":true,"sections":["BULUŞUN DETAYLI AÇIKLAMASI"],"evidence":""}}],
  "quality_notes":[""]
}}

NOTLAR:
- 'unumbered_system_definition' ve 'unumbered_system_elements' REFERANS NUMARALARI bölümünden önce kullanılacağı için (1), (2), (1001) veya benzeri referans işareti içermemeli, tamamen numarasız olmalıdır.
- elements alanında kaynaktaki gerçek sistem/cihaz referans işareti varsa aynen korunur. Açık sistem modülleri bulunmasına rağmen ayrı unsur numarası yoksa kaynak sırasıyla 1, 2, 3... atanır. Kaynakta UW, UW_F, UW_PL, UW_R, UW_M gibi sembolik referanslar kullanılıyorsa number alanına bu sembol yazılabilir. Kaynakta gerçek unsur referansı olmayan 21-37 gibi şekil içi/geçici numaralar yeni referans numarası olarak üretilmez. Unsur adı teknik anlamı değiştirilmeden referans listesi için yalnızca ilk kelimenin baş harfi büyük olacak cümle düzenine çevrilir, standart teknik kısaltmalar korunur.
- method_steps alanında numara ayrı, metin numarasızdır. Kaynakta yöntem işlem adımları için açık referans verilmişse (1, 2, 3..., S101..., M20... vb.) aynen korunur. Kaynakta hiçbir yöntem referansı verilmemişse 1001, 1002, 1003... varsayılanı kullanılır. Kısmen numaralandırılmış kaynakta yalnız eksik adımlara çakışmayacak varsayılan referans atanır. Metin mastar isimle biter: toplanması, analiz edilmesi vb.
- Word çıktısındaki REFERANS NUMARALARI bölümünde yöntem adımı `1001. ...` biçiminde önde yöntem referansıyla yazılır. Bu bölümde yöntem adımının içinde geçen sistem/cihaz unsur işaretleri `(1)`, `(2)` vb. gösterilmez; bunlar BULUŞUN DETAYLI AÇIKLAMASI bölümünden itibaren kullanılır. method_steps içindeki teknik metin, detaylı açıklama ve istem için unsur referanslarını içerebilir; Word üretici referans-listesi görünümünde yalnız bilinen unsur işaretlerini kaldırır.
- method_claim varsa her adım metninin sonunda '(REF)' biçiminde, method_steps alanındaki aynı müşteri/varsayılan referansı yer alsın. Ara işlem adımları virgülle bitsin; son işlem adımının sonunda noktalama olmasın ve ardından ayrı satırda “işlem adımlarını içermesidir.” gelsin. method_claim içindeki işlem metinleri method_steps ile numara hariç birebir aynı olsun.
- Sistem istemi elemanları bağımsız bir alışveriş listesi gibi yazılmasın. Teknik olarak ilişkili her sonraki unsur, mümkün olduğunda daha önce tanımlanan unsurun çıktısı/girdisi veya bağlantısı üzerinden kurulmalı; kaynakta olmayan yapay ilişki eklenmemelidir.
- İstemlerde İngilizce kısaltmalar ilk kullanımda Türkçesiyle bir kez açıklansın; devamında Türkçe kullanılsın.
- Kullanıcıya sunulan tarifname metninde BBF veya buluş bildirim formu ifadeleri kullanılmasın.
- Sistem ve yöntem birlikteyse başlık “... Sistemi ve Yöntemi” biçiminde olsun.
- “Buluşun bir gerçekleştirilmesinde” yerine “Buluşun bir yapılanmasında” kullanılsın.
- TEKNİK ALAN iki paragraf olmalıdır. İlk paragraf yalnız “Buluş, ... ile ilgilidir.” giriş cümlesinden oluşup burada bitmelidir. İkinci paragraf mutlaka “Buluş, özellikle ...” ile başlamalı ve daha ayrıntılı teknik kapsamı vermelidir. “Sistem ve yöntem...” gibi çıplak bir ifadeyle ikinci paragrafa başlama. technical_field içinde iki paragrafı \n\n ile ayır.
- Bağımlı istemlerde “Önceki istemlerden herhangi birine” kalıbını varsayılan olarak kullanma. Ek özellik hangi unsur/işlem adımına dayanıyorsa doğrudan o unsuru ilk tanımlayan en yakın ve gerekli isteme bağla; ana istemde tanımlı bir modülün ayrıntısı için doğrudan ana isteme bağlanmayı tercih et. Gereksiz “İstem X veya Y’ye” zincirleri kurma.
- BULUŞUN DETAYLI AÇIKLAMASI giriş cümlesinde buluş adı başlıktaki Title Case biçimiyle değil cümle içi normal küçük harf düzeninde yazılsın; SIM/eSIM/API gibi teknik kısaltmalar korunsun.
- Detaylı açıklamada referanslı unsurlar tek sürekli paragrafta anlatılsın. “İşlem Adımı / Gerçekleştiren Unsur / Açıklama” türü sistem-yöntem ilişki tablosu oluşturma; bu içeriği müşterinin verdiği sistem ve yöntem referansları arasındaki teknik bağı gösteren doğal paragraf olarak yaz. Kaynakta referans yoksa sistem için 1,2,3... ve yöntem için 1001,1002... varsayılanları kullanılabilir. Yalnız gerçek sayısal/deneysel veri tabloları tablo olarak korunabilir.
- Literatür paragraflarında İngilizce başlık ve Türkçe karşılığı birlikte yazılsın.
- objectives alanındaki her amaç tam cümle yüklemiyle bitsin: “... karşılaştırmaktır.”, “... sağlamaktır.” gibi. “... karşılaştırmak.” veya “... sağlamak.” biçiminde çıplak mastar bırakma.
- Açıklama bölümlerinde noktalı virgülü gereksiz kullanma; virgül veya nokta tercih et. İstemlerdeki standart “olup, özelliği;” kalıbı korunabilir.
- İstemlerde HPU_W, FW_min, UW, UW_F, FE, TE gibi semboller anlamı açıklanmadan çıplak teknik unsur olarak kullanılmasın.
- BBF veya açık teknik müşteri kaynağında kullanılabilir teknik şekil/şema/akış diyagramı varsa bunu yalnız 'tercih' olarak değil ZORUNLU kaynak şekil olarak kabul et; özgün müşteri görseli nihai Şekiller çıktısında yer almadan model üretimi veya yeniden çizilmiş yardımcı şema ile ikame etme.
- Şekilde referans listesinde ayrı olan iki unsur tek `2-3`/tek kutu/tek hedef olarak birleştirilmişse ortak taşıyıcı ve teknik ilişkileri koruyup ayrı kutucuk/çağrı/oklarla AYIR. Şekiller Word istenmişse tüm gerçek sistem referansları şekil setinde en az bir kez; sistem+yöntem akış şekilleri hazırlanıyorsa tüm yöntem adımı referansları da akış şekillerinde en az bir kez gösterilsin. Kaynakta hem sistem görünüşü hem yöntem/akış diyagramı varsa ikisini de koru. Yalnız kullanıcı açıkça hariç bırakırsa veya daha sonra verdiği düzeltilmiş müşteri şekli aynı görselin yerini alıyorsa eski sürümü çıkar. Şekillerde müşteri görselini esas al; şekil üzerindeki gerçek unsur/yöntem/sembolik referansları REFERANS NUMARALARI ile senkron tut. Gömülü grafik, ısı haritası ve diyagramlardaki teknik sonuçları içerik tamlığı açısından incele. Geçici şekil numaralarını yeni unsur numarası olarak uydurma. Eksik/yanlış bir şekil referansı değerlendirilirken referans işareti → unsur adı → detaylı açıklamadaki teknik tanım → şekil üzerindeki fiziksel karşılık sırasıyla doğrulanmalı; kılavuz çizgisi/ok doğrudan ilgili fiziksel unsurda sonlanmalı, alt parçaya ait referans tüm tertibatı göstermemelidir. Her görünür parçayı zorla numaralandırma; yalnız tarifnamede gerçek referansla tanımlı ve şekil üzerindeki yeri güvenilir biçimde belirlenebilen unsuru işaretle. Konum belirsizse uydurma ok/numara üretme. Önce “kuru hibrit güç ünitesi ağırlığı (HPU_W)”, “asgari görev yakıtı ağırlığı (FW_min)”, “ilave yakıt tahsisi (UW_F)” gibi açılımı yaz, matematiksel ifadede sembolü koru.

- BBF/BOM içindeki “Diğer parçalar/Diğer elemanlar” gibi belirsiz başlıkları nihai referans unsuru yapma; gerçek teknik parçaları net adlandır.
- Sistem/cihaz/ürün/tertibat/yapılanma aynı ürün istem dil ailesidir. Yöntem dışındaki alt istemler “olmasıdır.” veya “içermesidir.” ile bitsin; “oluşturulmasıdır/bağlanmasıdır/sağlanmasıdır” kullanma.
- Ana istemde bir referanslı unsuru ilk kez tanımlarken henüz tanımlanmamış sonraki referansları kullanma; kural olarak her bullet tek yeni referanslı unsur tanımlasın.
- Parantezli unsur referansını yalnız REFERANS NUMARALARI listesindeki aynı unsur adıyla veya dilbilgisel çekimiyle kullan. Örn. `1 = İnsansız hava aracı` ise `İHA (1)`/`İHA’dan (1)` YASAK; `insansız hava aracı (1)`/`insansız hava aracından (1)` kullan. Kısaltma numarasız kullanılabilir.
- Yazılım/modül ağırlıklı istemde kaynakta referanssız olan `elektronik işlem birimi` gibi taşıyıcıyı, henüz tanımlanmamış 2/3/4/5 gibi modülleri topluca sayan ayrı bullet yapma. Taşıyıcı ilişkisini her ilgili modül ilk kez tanımlanırken aynı bullet içinde `elektronik işlem birimi üzerinde koşturulan yazılım vasıtasıyla çalışan ... modülü (N)` biçiminde kur; taşıyıcı kaynakta ayrı referanslı unsur ise normal sırayla tanımla.
- Alternatif ve tercih edilebilir yazım stili: aynı referanssız elektronik işlem birimi üzerinde aynı çalışma ilişkisine sahip ARDIŞIK birden fazla modül varsa taşıyıcıyı her modülde tekrar etmek yerine `elements` listesinde `{{"lead":"bir elektronik işlem birimi üzerinde koşturulan yazılım vasıtasıyla çalışan ve;","subelements":["... (2),","... (3),"]}}` biçiminde ortak üst bullet + ayrı gerçek Word alt bullet grubu kullan. Lead referans taşımaz; her subelement tek yeni referans tanımlar ve ilk-tanım sırasına uyar.

- Ürün isteminde işlem isimleştirmesi değil unsur dili kullan: “... bağlanması” değil “... bağlanan ...”.
- Bir zorunlu unsurun kaynakta açık teknik işlevi varsa yalnız konumunu değil işlevini de yaz.
- Aynı olmayan unsurları “ve/veya” ile tek unsur gibi birleştirme.
- Zorunlu olmayan “vidalanan/kaynaklanan/yapıştırılan” ve örnek ölçü/diş/çap ifadeleriyle ana istemi daraltma; teknik farkın kendisiyse koru.
- Bağımlı istemler ana/üst istemi anlam olarak tekrar etmesin; sistem ve yöntem alt istemlerinin HER BİRİ semantik tekrar kontrolünden geçsin. Her biri gerçek ek teknik sınırlama getirsin ve bağımlılık numaraları silme/değişiklik sonrası yeniden kontrol edilsin.
- Bağımlı sistem istemini yalnız `sistemin ... ortamında çalışmaya uygun bir sistem olmasıdır` diye kurma; 5G/6G/hibrit vb. ortam özelliğini ilgili baz istasyonu, arayüz, iletişim birimi veya başka somut unsurun teknik niteliğine bağla. Bağımlı yöntemde de yalnız `yöntemin ... ortamında gerçekleştirilmesidir` demek yasaktır; ortamı gerçek bir işlem adımı/girdi/teknik taşıyıcı ile ilişkilendir.

- Özel uygulama unsur adını gereksiz daraltma; kaynak destekliyorsa O-ring gibi örneği “sızdırmazlık elemanı” altında detaylı açıklamada ver. Her ayrıntıya zorla referans numarası verme.
- Kullanıcıya görünen tarifnamede “müşteri tarafından iletilen teknik çizimde / müşteri bilgilerine göre / ek teknik belgede” gibi kaynak atıfları bulunmasın.
- Özet tek paragraf ve tek cümle olsun.

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
