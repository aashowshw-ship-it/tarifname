# RED KARARINA İTİRAZ — UYGULAMA AKIŞI VE KALİTE KAPILARI
**Kural şablonu**: `700851_SOURCE_OF_TRUTH.md`. **Arayüz**: `INPUT_SCHEMA.json`. Bu iki dosya birlikte uygulanır. Başvuru türü farklıysa tek dokümandan iki farklı teknik savunma üretilir. Varsayılan dil Türkçe (Türkiye).

## Kullanıcıdan tek seferde girdiler
Tüm dosyalar tek sayfada alt alta yüklenir. Ret kararı, son tarifname ve istemler, şekiller, son rapor, önceki görüşler, D belgeleri, üçüncü kişi görüşü, müşteri bilgisi, ana dosya referansı, işlem referansı ve tebliğ günü aynı ekrandadır. Müşteriden görüş yoksa **bağımsız teknik savunma yapılır**. Patent mi FM mi olduğu ret kararından otomatik çıkarılır. Başvuru numarası/sahibi de resmî kaynaktan çıkarılır. Kullanıcıdan tekrar alınmaz.

## 1. Kaynak doğrulama
- Karardaki idari usul, tarih, başvuru no, başvuru türü, ret sebepleri ve ret edilen istemler okunur. Yalnız red kararı YİDD/YİDK yoluna gider, belge kararına itirazı ayrı iş türüdür.
- Son resmî istem kümesi birebir okunur, bağımsız/bağımlı istem kapsamı ve numarası doğrulanır.
- Ret kararının esas aldığı rapor ve daha önce sunulan görüş karşılıklı kontrol edilir. Patentte son inceleme raporu esas olup önceki 2.–4. raporlar **varsa** ek kanıttır. FM araştırma raporu ve düzeltmeleri esas alınır.
- İstem değişikliği geçmişi tartışmalıysa ilk başvuru kaynağı ayrıca istenir. Otomatik Markup/Clean ve yeni istem üretimi yapılmaz.
- D1, D2, ... kaynaklarının kimliği, tam metni, tarihi, özgün figürleri ve atıfları kontrol edilir. Kaynağı kontrol edilmemiş ‘yok’ iddiası FAIL.
- Tebliğ tarihi farklıdır, karar tarihi asla tebliğ tarihi diye kullanılmaz. Tebliğ günü yoksa süre için NOT_VERIFIED, Word çalışması için diğer kanıtlar yeterliyse analiz sürer.

## 2. OBJ bazlı ret kararı inceleme
- Ret kararındaki her bağımsız gerekçe `OBJ-1`, `OBJ-2`... ile kodlanır.
- Her OBJ için `ret_karari_ozet`, `ilgili_istemler`, `ret_dayandigi_D`, `onceki_gorus_cevabi`, `tarifname_dogrudan_alinti`, `itiraz`, `sonuc` alanları doldurulur.
- Uzmanın iddiasını eksiksiz ve tarafsız özetle, ardından **kararlı, kesin ve teknik kaynak dayanaklı** cevap ver. ‘Olabilir, yapılabilir, düşünülmektedir, kısmen’ savunma ve ihtimal dili YASAK.
- Uzmanın belirttiği ‘özellik 1 numaralı istemde yok’ eleştirisiyle yüzleş: gerçekten yoksa bunu inkâr edip uydurma fark gösterme, savunmayı **mevcut istemde bulunan diğer teknik ilişkiler** üzerinden kur. Bilinen özellikleri inkâr etme.
- Uzmanın teknik eşdeğerlik iddiası varsa aynı isim benzerliği ile yetinme. Fiziksel konum, işlev, eleman bağıntısı, veri/enerji aktarımı ve montaj özelliğiyle açıkça cevap ver.

## 3. Faydalı model yolu
- SMK 142 kapsamında yenilik ve teknik katkı; sanayiye uygulanabilirlik olumsuz değilse uzun savunmaya gerek yok.
- D1, D2, ... her kaynak **tek tek** tam bağımsız istem unsur matrisi. Tek kaynağın bütün özellikleri **aynı teknik ilişkilerle doğrudan ve açık** açıklaması gerekir. Yalnız bir unsur bulunmuyor diye değil, teknik bütün bağlamında analiz et.
- SMK 142/2 ‘teknik katkı yok’ gerekçesi için alt istemlerin her unsurunu, mekanik/elektronik işlevini ve gerçek teknik etkisini ayrı savun. Sırf ergonomi diyerek katkısız diye sınıflandırma.
- D1+D2 karışımı yenilik sonucunun kaynağı olamaz. Kararın belgelere ayrı X verdiği ve birlikte sonuç kurduğu hallerde, **tek kaynağa dayalı değerlendirme yükümlülüğünü** açıkça savun.
- **FM resmî Word'de buluş basamağı yok.** En son sözleşmedeki FM kesin kapanış.

## 4. Patent yolu
- Yenilik: Her X belgesi ayrı tam istem matrisi.
- Buluş basamağı: En yakın önceki teknik, farklı teknik unsurlar, doğrudan dayanaklı teknik etki, objektif teknik problem, birleşim için motivasyon ve yönlendirme olup olmadığı, D1+D2/D3 birleşimlerine teknik cevap. Bu bölüm sadece gerçek karar buluş basamağını ret nedeni yapmışsa yazılır.
- Açıklık, destek, yeterli açıklama, buluş bütünlüğü, patentlenebilir konu, kapsam aşımı gibi ret gerekçeleri varsa her biri kendi başlığı ve dayanağı ile ayrıca cevaplanır.
- Patent için **ayrı** kesin kapanış. Yalnız savunulan ret gerekçeleri kapsanır.

## 5. Görsel deliller
- Başvurunun şekilleri ve önceki teknik figürleri kaynağın orijinal sayfasından çıkartılır. Kırmızı daire/ok zorunlu değildir. Doğru parça/tam şekil anlamı açık değilse **hiç işaretleme yapma**.
- Şekillerin altına kaynak doküman + gerçek figür/sayfa ve ayırt edici fark konur. Görsel, istemdeki olmayan özelliğin doğrudan isteme eklenmesi amacıyla kullanılamaz.
- İnceleme ve onaylı metin kaynakları aslına sadıktır. Şekil sırası, boş paragraf düzeni ve öncesi/sonrası boşluklar şablondan kaynaklanır.

## 6. Şablona kesin uyum
- Onaylı örnek `Red Kararına İtiraz_700851.docx` özel Kitaplıkta tutulur ve **SHA-256** ile tanınır. Daha eski 700801/696523 sadece tarihsel arşivdir.
- İki başlangıç boş paragrafı, bilgi satırlarının 2,5 satır aralığı, ‘Referansımız’ sonrası 22 pt, gerçek bölüm öncesi boşluk paragrafları, bölüm başlıklarının her birinin **orijinal** alt çizgi durumu ve 14/7 pt aralıkları korunur.
- Sadece başlıklar, özgün birebir alıntılar ve VAR/YOK unsur öncelikleri kalın. Ana savunma metni normal.
- Word yazım denetimi `tr-TR`. Özellikle hedef dil başka dile düşmüşse FAIL.
- **Noktalı virgül yazma**: Yeni savunma metninde nokta/virgül tercih edilir. Noktalı virgül ancak *birebir kaynak alıntısının parçası* ise korunur. Kaynağa sadakat için mevcut onaylı örnek DOCX'e edit uygulama.
- Uzun E-sayı ve ‘İstemlerde değişiklik yapılmamaktadır’ gibi iç süreç/prosedür meta sözleri Word içinde YASAK.
- Kapanış iki kısa paragraftan uzun olamaz, kesin talep cümlesi `700851_SOURCE_OF_TRUTH.md` içindedir.
- Çıktı adı `Red Kararına İtiraz_{görüş_ref}.docx`, Word içindeki referans `{ana_dosya_ref}`.

## 7. Zorunlu QA nesnesi ve hata durumları
`required_gates`:
- `source_integrity`, `appeal_route`, `protection_type`, `claims_verbatim`, `decision_obj_inventory`, `report_history`, `each_document_novelty`, `claim_matrix_all_features`, `every_obj_response`, `dependent_claims`, `technical_basis_quotes`, `figure_provenance`, `case_metadata`, `approved_template_sha`, `real_blank_paragraphs`, `heading_underline_exact`, `bold_rules`, `turkish_word_language`, `semicolon_policy`, `short_conclusion`, `word_visual_render`, `independent_technical_review`.
- Ek patent kapısı: `inventive_step_when_rejected`. Ek FM kapısı: `no_inventive_step`.
- `FAIL` veya `NOT_RUN` geçersiz. D1/D2'nin her “YOK” unsuru özgün kaynak sayfa/şekille desteklenmezse başarısız. Alıntı kaynağı birebir bulunamazsa başarısız.
- Son DOCX ve kullanılan ham kaynak SHA-256 kontrolüne bağlanır. Word sayfaları render edilip gerçek sayfa görüntülerinde üstbilgi çakışması, şekil taşıması, başlıkların boşluk ve alt çizgi durumu, logo ve sayfa kırılmaları denetlenir.
- Uygulama iş türü tamamlanıp menüye bağlandığında bu kapılar server-side çalışır. Sadece bu metnin GitHub'a kaydı, otomatik iş akışının üretimde çalıştığını kanıtlamaz.
