# Patent Atölyesi v5.4.15

Bu paket, mevcut Render/GitHub tabanlı **Patent Atölyesi** uygulamasının 13.08.2026 tarihli güncel tam sürümüdür.

## GitHub'a yükleme

ZIP'i açın ve içindeki dosyaların tamamını mevcut GitHub deposunun **ana dizinindeki** dosyalarla değiştirin. `Patent_Atolyesi_v5.4.15_GitHub` klasörünü ikinci bir alt klasör olarak yüklemeyin.

Depo kökünde en az şu dosyalar doğrudan görünmelidir:

- `app.py`
- `rules.py`
- `RULES_MEMORY.md`
- `Dockerfile`
- `render.yaml`
- `requirements.txt`
- `packages.txt`
- `Tarifname_181176_template.docx`
- `Gorus_metni_696809_template.docx`
- `On_Arastirma_Raporu_181612_template.docx`

GitHub web arayüzünde **Add file → Upload files** ile bütün dosyaları yükleyin, aynı isimli dosyaların değiştirilmesine izin verin ve commit edin. Render servisi `autoDeploy: true` ise yeni commit otomatik yayınlanır.

## Render ortam değişkenleri

- `OPENAI_API_KEY`
- `OPENAI_MODEL`
- `FIGURE_REFERENCE_CONFIDENCE` (opsiyonel; varsayılan `0.86`)

`render.yaml` varsayılan model değerini `gpt-5.6` olarak taşır. Hesabınızda farklı model adı kullanılıyorsa Render Environment ekranından değiştirin.


## Şekil referans doğrulama ve otomatik düzeltme – v5.4.12

Tarifname oluştururken ayrı `Şekiller` Word çıktısı istenmişse şekiller artık Word'e doğrudan aktarılmadan önce nihai tarifname ile çapraz kontrol edilir.

- Her şekil için `referans işareti → unsur adı → detaylı açıklamadaki teknik tanım → şekil üzerindeki fiziksel karşılık` eşleştirmesi yapılır.
- Mevcut numara/ok doğru varsayılmaz. Örneğin `9 = Travers` ise kılavuz çizgisi traversin kendisine, `1 = Topuz` ise topuzun kendisine yönelmelidir.
- Referans listesindeki bir unsur ilgili şekilde açıkça görünür fakat numarasızsa, fiziksel konumu güvenilir biçimde belirlenebiliyorsa eksik numara/ok eklenir. Her görünür parça zorla numaralandırılmaz.
- Yanlış fiziksel parçaya yönelen mevcut referans oku güvenilir biçimde düzeltilebilir; otomatik görsel düzenleme yalnız numara ve kılavuz çizgisi/oklarla sınırlıdır.
- Düzenleme sonrasında özgün ve revize görsel ikinci kez karşılaştırılır. Teknik geometri değişmişse veya referans hâlâ belirsiz/yanlışsa revize görsel reddedilir ve `Şekiller` Word çıktısı oluşturulmaz. Tarifname Word çıktısı bundan bağımsız olarak verilir.
- Bu kontrol patent literatürü şekilleri için kullanılmaz. Tip 3 raporundaki D1/D2 şekilleri yine yalnız özgün patent kaynaklarından alınır ve yapay zekâ ile yeniden çizilmez.

## Görüş akışı

Görüş bölümü artık doğrudan Word üretmez.

1. Dosyalar yüklenir.
2. **1. Raporu analiz et** çalıştırılır.
3. Sistem rapordaki itirazları, önceki görüşü, tarifnameyi, X/Y dokümanlarını ve varsa müşteri bilgisini birlikte analiz eder.
4. İstem revizyonu gerekmiyorsa mevcut istemlerle **Görüş metnini oluştur** düğmesi açılır.
5. Revizyon gerekiyorsa önce istem bazında öneriler gösterilir: istem no, gerekçe, tarifname dayanağı, mevcut ifade, önerilen ifade.
6. Kullanıcı isterse revizyon önerilerini ek talimatla yeniden analiz ettirir.
7. Kullanıcı revizyonu onaylarsa `.docx` kaynak tarifname üzerinde iki çıktı üretilir:
   - gerçek OOXML Track Changes içeren `Düzenlenen_tarifname_track_changes_<referans>.docx`
   - kabul edilmiş değişiklikleri içeren `Düzenlenen_tarifname_temiz_<referans>.docx`
8. Kullanıcı revize istemleri son kez onayladıktan sonra görüş Word dosyası oluşturulur.
9. Kullanıcı revizyon önerisini gördükten sonra açıkça revizyonsuz devam etmeyi de seçebilir.

Track Changes, bütün istem paragrafını silip yeniden eklemek yerine mümkün olan en küçük ifade/kelime düzeyinde uygulanır.

## Tarifname kuralları

- `Tarifname_181176_template.docx` bağlayıcı şablondur.
- Kaynaktaki bütün teknik bilgi, özellikle önceki teknik, teknik problem/çözüm, unsurlar, yöntem adımları, formüller, tablolar, deneysel sonuçlar, alternatifler, kullanım senaryoları ve teknik etkiler eksiksiz aktarılır. Gömülü teknik şekiller, grafikler, ısı haritaları, eksen/etiket bilgileri ve bu görsellerden anlaşılan teknik sonuçlar da tamlık kontrolünün parçasıdır.
- Kullanıcıya sunulan tarifname metninde `BBF` veya `buluş bildirim formu` gibi kaynak-form atıfları kullanılmaz.
- Tarifname oluşturma ekranında çıktı dili `Türkçe` veya `İngilizce` seçilebilir. Dil değişse de aynı `Tarifname_181176_template.docx` biçimi ve bütün kaynak/istem/referans/şekil kuralları korunur. İngilizce çıktıda başlıklar, açıklama, istemler ve özet İngilizce yazılır.
- Sistem ve yöntem istemleri birlikteyse başlık seçilen dile uygun biçimde `... Sistemi ve Yöntemi` veya `... System and Method` olur.
- `REFERANS NUMARALARI` unsur adlarında yalnızca ilk kelimenin ilk harfi büyük yazılır; teknik kısaltmalar korunabilir.
- Detaylı açıklamadaki yöntem adımları madde işaretli yazılır ve müşteri tarafından verilmiş yöntem referansı adımın sonunda parantez içinde korunur. Kaynakta yöntem referansı yoksa varsayılan `(1001)`, `(1002)`... kullanılır.
- Sistem/yöntem bağımsız istemlerindeki ayrı unsur ve adımlar madde işaretlidir.
- `İSTEMLER` ve `ÖZET` ayrı yeni sayfalardan başlar; başlıklar kalındır. İstemlerde sonradan değişiklik yapılsa da `ÖZET` öncesindeki sayfa geçişi korunur.
- `Buluşun bir gerçekleştirilmesinde` kullanılmaz; `Buluşun bir yapılanmasında` kullanılır.
- Önceki teknik bölümündeki müşteri anlatımı eksiksiz korunur; patent literatürü bunun yerine geçmez.
- Tarifname oluşturma arayüzünde `Mevcut/revize tarifname` alanı yoktur. Mevcut bir tarifnamenin değiştirilmesi ileride ayrı `Tarifname düzenleme` iş akışında ele alınacaktır; yeni tarifname oluşturma akışına karıştırılmaz.
- Şekiller seçimi literatür araştırmasından önce gösterilir.
- `TEKNİK ALAN` **iki paragraf** halinde kurulur. İlk paragraf yalnızca tek giriş cümlesidir ve `Buluş, ... ile ilgilidir.` biçiminde biter. Ardından mutlaka yeni paragraf açılır; ikinci paragraf `Buluş, özellikle ...` ile başlar ve teknik alanın ayrıntısını verir. İkinci paragraf `Sistem ve yöntem...` gibi bir ifadeyle başlatılmaz.
- `ÖNCEKİ TEKNİK` içinde aynı anlatımın devamı olan `Özellikle`, `Bununla birlikte`, `Bu nedenle` gibi cümleler gereksiz yere ayrı paragraf yapılmaz.
- Türkçe tarifnamede patent literatürü paragraflarında doğrulanmış İngilizce başlık ve Türkçe karşılığı birlikte yazılır. İngilizce tarifnamede özgün İngilizce patent başlığı kullanılır.
- `BULUŞUN DETAYLI AÇIKLAMASI` içinde referanslı unsurlar tek tek ayrı paragraf yapılmaz; unsur açıklamaları tek sürekli paragrafta birleştirilir.
- Detaylı açıklamadaki yöntem işlem adımlarında ara maddeler virgülle, son madde noktayla biter. Türkçe bağımsız yöntem isteminde ara işlem adımları virgülle, **son işlem adımı noktalamasız** biter ve ardından `işlem adımlarını içermesidir.` yazılır. İngilizce istemlerde doğal `comprising:` claim yapısı kullanılır.
- `Tarifname_181176_template.docx` fontların yanı sıra boş satır, 1,5 satır aralığı, gerçek Word madde işaretleri/numaralandırması ve istemler arası boşluk bakımından da birebir bağlayıcıdır.
- Amaç cümleleri `... karşılaştırmaktır.`, `... sağlamaktır.` gibi tam yüklemle biter; `... karşılaştırmak.` biçiminde bırakılmaz.
- REFERANS NUMARALARI bölümünden önce `(1)`, `(2)` gibi referans işaretleri kullanılmaz; kısa açıklamadaki ana istem özeti numarasız yazılır.
- Müşterinin sistem/cihaz unsurları ve yöntem işlem adımları için verdiği referans işaretleri aynen korunur; `10, 20...`, `S101...`, `M1...` veya başka bir aile sırf standardizasyon için değiştirilmez. Sistem unsurlarında hiç referans yoksa `1, 2, 3...`; yöntem adımlarında hiç referans yoksa `1001, 1002, 1003...` varsayılanı kullanılır. Kısmen numaralandırılmış kaynakta yalnız boş kalan referanslar çakışmayacak varsayılanlarla tamamlanır.
- `REFERANS NUMARALARI` bölümünde önce sistem modülleri art arda yazılır, ardından **tek bir boş paragraf** bırakılarak yöntem işlem adımları müşterinin verdiği referans ailesiyle; kaynakta referans yoksa `1001...` varsayılanıyla art arda yazılır.
- Sistem ile yöntem arasındaki `İşlem Adımı / Gerçekleştiren Unsur / Açıklama` ilişkisi tarifname gövdesinde açıklama tablosu olarak verilmez. Bu ilişki, modüllerin hangi işlem adımını hangi veri/çıktıyı kullanarak gerçekleştirdiğini açıklayan doğal teknik paragraf halinde yazılır. Yalnız kaynağın gerçek sayısal/deneysel veri tabloları gerektiğinde tablo olarak korunabilir.
- `ŞEKİLLERİN KISA AÇIKLAMASI` bölümünde `Şekil 1...`, `Şekil 2...`, `Şekil 3...` açıklamaları aralarında boş paragraf olmadan doğrudan alt alta yazılır.
- Kaynakta `UW`, `UW_F`, `UW_PL`, `UW_R`, `UW_M` gibi sembolik referanslar varsa sayısal unsur listesinden sonra bir boşlukla `UW. Kullanılabilir ağırlık` biçiminde yazılabilir; kaynakta gerçek referans olmayan `21-37` gibi geçici şekil numaraları uydurulmaz.
- Aynı numaralı yöntem adımının teknik metni `REFERANS NUMARALARI`, detaylı açıklamadaki yöntem listesi ve bağımsız yöntem isteminde birebir senkron tutulur. Bir yerde değişirse üçü birlikte güncellenir. Bağımsız yöntem isteminde son işlem adımı virgülsüz/noktalamasız biter.
- Bağımlı istemlerde `Önceki istemlerden herhangi birine` kalıbı varsayılan olarak kullanılmaz; ek özelliğin dayandığı en doğrudan istem numarası seçilir. Ana istemde tanımlı bir modülün ayrıntısı çoğunlukla doğrudan ana isteme bağlanır.
- İstemlerde teknik olarak ilişkili unsurlar birbirinin girdisi/çıktısı veya bağlantısı üzerinden yazılır; bağımsız unsur listesi gibi bırakılmaz.
- Buluş ağırlıklı olarak yazılım/algoritma/modül/birimlerden oluşuyorsa bağımsız sistem ve yöntem istemleri soyut yazılım diliyle bırakılmaz. Kaynakta özel donanım gerekmiyorsa geniş taşıyıcı olarak `bir elektronik cihaz üzerinde koşturulan yazılım vasıtasıyla ...` veya eşdeğer elektronik cihaz/işlem birimi dili kullanılır; gereksiz sunucu, cep telefonu veya bilgisayar daraltması yapılmaz.
- Bağımlı istemlerde her kaynak ayrıntısını ayrı isteme dönüştürmek zorunlu değildir. Yalnız gerçek teknik daraltma ve stratejik geri çekilme konumu sağlayan seçilmiş özellikler kullanılır; ana istemdeki elektronik cihaz/yazılım dayanağı alt istemlerde gereksiz yere tekrar edilmez.
- İstemlerde `HPU_W`, `FW_min`, `UW_F` gibi sembollerin teknik açılımı önce yazılır ve sembol parantez içinde gösterilir; formüllerde semboller korunur.
- Teknik açıklama metninde noktalı virgül gereksiz kullanılmaz; standart `olup, özelliği;` istem kalıbı istisnadır.

## Tip 3 ön araştırma

- Akış adım atlamadan yürür: BBF analizi → DP referansı → global tam 10 doküman + TotalPatent sorgusu + önerilen D1/D2 → kullanıcı benzer dokümanları → kullanıcı dokümanlarından yalnız en ilgili birkaç `10+` belge → D1/D2 değişim durumu → sistemin `Bence buluş basamağı var/yok` kanaati → kullanıcı sonuç modu → Word raporu.
- İlk 10 doğrulanmış doküman sonradan kullanıcı belgesi geldi diye yeniden yazılmaz. Kullanıcı belgeleri ayrı `10+ XX... or YY...` satırında gösterilir.
- Kullanıcı dokümanlarından sonra D1 ve D2'nin değişip değişmediği eski/yeni numaralarla açıkça belirtilir.
- Sonuç seçimi, sistem kanaatinden önce gösterilmez. Seçenekler `Buluş basamağı var / Buluş basamağı yok / Otomatik belirle`dir.
- D1 ve D2 tablolarında sol teknik özellik listesi birebir aynıdır. Sağ hücre yalnız `+` / `-` değildir; işaretin ardından özelliğin dokümanda geçtiği somut yer (`Özet`, `İstem`, `Şekil`, paragraf/sütun vb.) yazılır.
- D1/D2 şekilleri model tarafından çizilmez. Yalnız özgün patent şekli patent kaynağından veya kullanıcı patent dosyasından alınır; temin edilemezse yapay şekil üretmek yerine hata/uyarı verilir.
- Rapor gövdesinde `BBF` ifadesi, ok zincirleri veya `özellik + özellik` türü yapay kısaltmalar kullanılmaz.
- `On_Arastirma_Raporu_181612_template.docx` yalnız görünüm örneği değil, doğrudan doldurulan bağlayıcı Word şablonudur; gövdesi silinip yeniden kurulmaz.
- `Anahtar Kelimeler` hücresi yalnız İngilizce teknik anahtar kelimeler içerir ve şablondaki 5x2 iç tablo aynen korunur; en fazla 10 ifade kullanılır.
- `IPC Kodu` alanındaki sınıflandırma açıklamaları İngilizce yazılır. Her satırda IPC/CPC kodu şablondaki gibi **kalın**, açıklaması normal yazıdır.
- `Kapsam` hücresi şablondaki sabit `Global (İlan edilmiş olan patent başvuruları)` metniyle korunur; araştırma kesim tarihi bu hücreye eklenmez.
- D1/D2 `Özet` alanına model özeti veya Türkçe çeviri yazılmaz. İlgili patentin doğrulanmış **özgün İngilizce Abstract** metni doğrudan kullanılır; kullanıcı orijinal patent dosyası sağladıysa öncelik o dosyadadır.
- Özgün İngilizce Abstract bulunamazsa model yeni abstract üretmez; dosya oluşturulmadan önce eksik kaynak uyarısı/hatası verilir.
- Uyarı hücresi şablondaki dört ayrı paragraf yapısında kalır. Word üretiminden sonra section/marj, ana tablo sayısı, 5x2 keyword tablosu, sabit Kapsam metni, IPC bold/normal run yapısı, İngilizce D1/D2 abstract ve uyarı paragraf sayısı otomatik biçim kontrolünden geçirilir.

## Araştırma güncelleme - Tip 3

Arayüzde yeni bir iş türü olarak **Araştırma güncelleme - Tip 3** bulunur. Kurallar kullanıcı ekranında uzun bir kural listesi olarak gösterilmez; `rules.py` arka planda bağlayıcıdır.

Akış:

1. `İlk BBF` yüklenir.
2. `Revize BBF` yüklenir.
3. `İlk Ön Araştırma Raporu` yüklenir.
4. DP referans numarası, çıktı dosya adı ve araştırma kesim tarihi girilir.
5. **Farkları ve teknik katkıyı analiz et** aşamasında ilk ve revize araştırma konusu arasındaki gerçek teknik farklar, teknik etkiler ve ilk rapordaki D1/D2 karşısındaki etkileri ekranda gösterilir. Bu analiz Word raporuna ayrı bölüm olarak taşınmaz.
6. **Revize konu için yeni patent araştırmasını yap** aşamasında revize edilen ayırt edici özellikler ve ilk rapordaki D1/D2 başlangıç noktası alınarak global araştırma yapılır; toplam 10 doğrulanmış doküman değerlendirilir. İlk raporda olmayan yeni yakın dokümanlar arayüzde ayrıca gösterilir.
7. Sistem yenilik ve buluş basamağı için kendi teknik kanaatini açıkça gösterir. Ardından kullanıcı `Buluş basamağı sağlanıyor` veya `Buluş basamağı sağlanmıyor` sonucunu seçer.
8. **Ön Araştırma Raporunu oluştur** aşamasında çıktı yine standart `On_Arastirma_Raporu_181612_template.docx` formatında hazırlanır; `Revizyon farkları` gibi yeni bölüm eklenmez.
9. Yeni bulunan belge D1/D2'den daha güçlü ise D1/D2 değişebilir. Yardımcı belge ise ayrı D3 başlığı açılmadan buluş basamağı değerlendirmesinde doğal paragraf olarak kullanılır.
10. İlk rapordaki D1/D2 korunuyorsa özgün patent şekilleri tekrar kullanılır; yeni D1/D2 seçilmişse yeni dokümanın özgün patent şekli kullanılır.

## Bağlayıcı şablonlar

- Tarifname: `Tarifname_181176_template.docx`
- Görüş: `Gorus_metni_696809_template.docx`
- Tip 3: `On_Arastirma_Raporu_181612_template.docx`

## Kural sürümü

Uygulamadaki kuralların tek yürütme kaynağı `rules.py` dosyasıdır. İnsan tarafından okunabilir kayıt `RULES_MEMORY.md` içindedir.

Kural sürümü: `2026-08-13.v5`

## Yerel çalıştırma

```bash
pip install -r requirements.txt
export OPENAI_API_KEY="..."
streamlit run app.py
```

Windows PowerShell:

```powershell
pip install -r requirements.txt
$env:OPENAI_API_KEY="..."
streamlit run app.py
```


### İstem teknik mekanizma kontrolü
Tarifname oluşturma sırasında bağımsız istemler yalnız hedeflenen fonksiyonu veya sonucu söylemekle bırakılmaz. Kaynakta dayanak bulunduğu ölçüde teknikte uzman kişinin “nasıl gerçekleştiriliyor?” sorusuna cevap verecek teknik unsur/taşıyıcı, girdi-veri ilişkisi, teknik işlem mekanizması ve çıktı ilişkisi istemde gösterilir. Yazılım ağırlıklı istemlerde elektronik cihaz/işlemci dayanağına ek olarak yazılımın bu taşıyıcı üzerinde hangi teknik işlem yoluyla sonucu ürettiği de kontrol edilir; buna karşılık tercihli uygulama ayrıntılarıyla ana istem gereksiz daraltılmaz.

### Şekiller - tarifname oluşturma
Şekil çıktısı müşteri görsellerini esas alır. Görseller `ŞEKİL 1`, `ŞEKİL 2` şeklinde sıralanır, başlık görselin altında bulunur ve her sayfanın üstünde `sayfa / toplam sayfa` göstergesi kullanılır. Sayfa sayısı sabit değildir. Şekildeki gerçek referans işaretleri tarifnamedeki `REFERANS NUMARALARI` ile senkron olmak zorundadır. Ayrıntılı şekil kuralları `rules.py` içinde arka planda uygulanır.


## v5.4.12 — Tek tuş tarifname kalite kapısı

Bu sürüm mevcut tarifname, görüş, Tip 3 ve araştırma güncelleme akışlarını korur. Tarifname oluşturma tarafına bağlayıcı kalite kapısı eklenmiştir. Ana istemde henüz tanımlanmamış referans kullanımı, belirsiz “Diğer parçalar” unsuru, ürün/yapılanma isteminde yöntem dili, gereksiz kapsam daraltma, bağımlı istem tekrarı ve örnek ölçülerin isteme taşınması kalite turunda özellikle denetlenir. Taslak yerel doğrulamadan geçmezse aynı tıklama içinde doğrulama hatası modele geri beslenerek en fazla iki ek otomatik düzeltme yapılır. Word üretiminden sonra şablon paragraf yapısı ve LibreOffice render smoke-test doğrulanır. v5.4.11 şekil referans/ok kuralları aynen korunmuştur.


## v5.4.14 — BBF %100 teknik kapsam ve yazılım taşıyıcı kalite kapısı

Bu sürümde tarifname oluşturma akışının en üst kuralı olan **BBF'deki tüm teknik bilgilerin uygun yerde eksiksiz kullanılması** yalnız prompt talimatı olmaktan çıkarılarak atomik kapsam kontrolüne bağlanmıştır. BBF ve ek teknik kaynaklar `technical_facts` maddelerine ayrılır; kişi/sicil/ödül/imza, form talimatları, boş idari alanlar ve yalnız araştırma anahtar kelimeleri teknik kapsam dışına alınır. Her zorunlu teknik fact nihai taslakta `source_coverage_map` ile bölüm ve kanıt metnine eşleştirilmeden Word üretimi yapılamaz.

Yazılım/modül ağırlıklı buluşlarda artık yalnız “işlemci” veya “donanım” kelimesinin istemde bulunması yeterli değildir. Modül/yazılım ile kaynakta dayanaklı teknik taşıyıcı arasında `üzerinde çalışan/koşturulan`, `vasıtasıyla` veya eşdeğer açık yürütme ilişkisi aranır. Kaynak SIM/eSIM üzerindeki güvenli işlemci, bellek ve izole çalışma ortamı gibi özel bir taşıyıcı veriyorsa bu taşıyıcı genel elektronik cihaz ifadesiyle kaybedilmez.

BULUŞUN DETAYLI AÇIKLAMASI giriş cümlesindeki buluş adı artık başlık Title Case biçiminde tekrar edilmez; cümle içinde normal küçük harf düzeni kullanılır ve SIM/eSIM/API/NFC gibi teknik kısaltmalar korunur. Arayüzde Word üretiminden önce `BBF teknik bilgi kapsam kontrolü` paneli her zorunlu fact'in tarifnamede hangi bölümde karşılandığını gösterir.

## v5.4.15 — Referans listesi ve alt istem tekrar kontrolü

Bu sürümde REFERANS NUMARALARI içindeki yöntem satırları `1001. ...` biçimine alınmıştır. Referans listesi içinde yöntem metninde geçen sistem/cihaz `(1)`, `(2)` gibi parantezli unsur işaretleri otomatik olarak kaldırılır; bu işaretler BULUŞUN DETAYLI AÇIKLAMASI bölümünden itibaren kullanılır. Word yapısal kalite kapısı bu ayrımı doğrudan denetler.

Ayrıca yerel semantik tekrar kalite kapısı sistem alt istemleriyle sınırlı olmaktan çıkarılmış, yöntem alt istemlerine de uygulanmıştır. Her alt istem ana/üst isteme gerçek teknik sınırlama eklemelidir; yalnız farklı kelimelerle tekrar eden istemler çıktı üretimini durdurur.


## 2026-08-13.v5 — Referans listesi ve alt istem semantik kalite kapısı
- `İSTEMLER` ve `ÖZET` DOCX yapısında `page_break_before=True` ile ayrı sayfadan başlatılır ve doğrulayıcı her iki başlık için bunu zorunlu olarak kontrol eder.
- `REFERANS NUMARALARI` altındaki yöntem işlem adımları `1001. ...`, `1002. ...` biçiminde önde yöntem numarasıyla yazılır. Bu satırlarda sistem/cihaz unsur işaretleri `(1)`, `(2)` vb. gösterilmez. Parantezli unsur referansları `BULUŞUN DETAYLI AÇIKLAMASI` bölümünden itibaren başlar.
- Aynı yöntem adımının teknik kelime dizisi üç yerde senkron tutulur; referans listesindeki sistem/cihaz parantez işaretlerinin bilinçli olarak kaldırılması senkron ihlali sayılmaz.
- Sistem ve yöntem bağımlı istemlerinin tamamı ana/üst istemlere ve önceki bağımlı istemlere karşı semantik tekrar kontrolünden geçirilir. Farklı kelimelerle aynı teknik sınırlamayı tekrar eden alt istem Word üretimini durdurur.
- Sistem şekillerinde görünüşte temsil edilen zorunlu ana taşıyıcı referans atlanmaz. Ok uçları küçük ve tutarlı tutulur. Yöntem akışına kaynak açıkça işlem-adımı döngüsü vermedikçe son adımdan önceki adıma geri dönüş oku eklenmez.
