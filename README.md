# Patent Atölyesi v5.4.52

## v5.4.52 / 2026-09-02 — Süreçler / Tarayıcı WebGPU AI

- Render Free üzerinde 512 MB RAM / 0.1 CPU nedeniyle sunucuda GGUF model çalıştırma kaldırıldı; Docker imajı artık yerel LLM indirmez.
- `Süreçler → Patent / Faydalı Model Başvurusu` önce kurallı parser ile ön kontrolü üretir, ardından desteklenen Chrome/Edge tarayıcısında **Qwen2.5-0.5B-Instruct q4 + WebGPU** ile alanları ikinci kez yapılandırır.
- Tarayıcı AI için OpenAI/Gemini/API anahtarı veya kredi gerekmez. Model ilk kullanımda tarayıcıya indirilir ve tarayıcı önbelleğinde tutulur.
- Başvuru kaynak metni harici bir AI API'sine gönderilmez; model kullanıcının tarayıcısında çalışır.
- AI çıktısı yine sunucudaki kaynak metinle doğrulanır. Kaynakta bulunmayan şirket adı, kişi adı, kimlik, e-posta, telefon vb. kabul edilmez.
- WebGPU kullanılamazsa sistem bunu açıkça gösterir ve kurallı ön kontrolle devam eder.
- Buluş başlığı daima Tarifname'den, DP referansı Tarifname dosya adından gelir.


## v5.4.51 / 2026-09-02 — Süreçler / Kredisiz Yerel AI Hibrit Çıkarım

- `Süreçler → Patent / Faydalı Model Başvurusu` artık **kurallı parser + yerel küçük Qwen modeli** ile hibrit çalışır. OpenAI/API çağrısı ve token kredisi kullanılmaz.
- Yerel model: Apache-2.0 lisanslı Qwen2.5-0.5B-Instruct tabanlı IQ2_XS GGUF; `llama.cpp` ile yalnız analiz düğmesine basıldığında tek seferlik çalıştırılır.
- Yerel AI; OCR/tablo/mail düzeni bozulduğunda şirket unvanı, buluşçu adı, e-posta, telefon, kimlik, adres vb. alanları yeniden yapılandırır. Modelin ürettiği değer kaynakta tekrar doğrulanamazsa kabul edilmez.
- Buluş başlığı daima Tarifname'den, DP referansı daima Tarifname dosya adından alınmaya devam eder; yerel AI bunları değiştiremez.
- `İmza ... e-posta ... İmza` gibi kirli alanlar ayrıca temizlenir. Yerel AI çalışamazsa ekran bunu görünür biçimde bildirir ve kurallı çıkarımla devam eder.
- Docker imajı `llama.cpp` çalıştırıcısını ve yaklaşık 325 MB'lık Qwen IQ2_XS modelini build sırasında indirir; ilk kullanımda ayrıca model indirilmez.


## v5.4.50 / 2026-09-02 — Süreçler / Koordinatlı OCR + Mail Başvuru Tercihleri

- Resim bilgi kaynaklarında düz OCR yerine Tesseract kelime koordinatlarından satır/sütun ilişkisi yeniden kurulur; büyük yatay boşluklar TAB olarak korunur.
- Form başlıkları ve açıklama cümleleri artık hak sahibi/buluş sahibi adı, TCKN/VKN veya adres gibi alanlara yanlış atanmaz; `Cinsiyet`, `Doğum Tarihi`, `Sahip Türü` gibi başlıklar veri sanılmaz.
- EML/MSG içindeki HTML tablo hücreleri korunur; Outlook mesajında yapılandırılmış HTML gövde varsa düz metin yerine bu yapı tercih edilir.
- Mail/yazı içinden `Buluşçu bilgileri gizlensin mi?`, `Kamu destekli proje kapsamında mı?` ve `Erken yayın talep ediliyor mu?` cevapları AI kullanmadan çıkarılır ve ön kontrolde ayrı gösterilir.
- Kamu destekli proje cevabı EVET ise kurum ve proje numarası zorunlu ön kontrole girer.
- Cevap bulunmazsa beyan formundaki belirtilen varsayım uygulanır: üç tercih için HAYIR; ön kontrolde `varsayılan` olarak açıkça gösterilir.
- Buluş başlığının tek otoritesi her zaman yüklenen Tarifname'dir; beyan/mail başlığı EPATS başlığını değiştirmez.

## v5.4.49 / 2026-09-02 — Süreçler / Gerçek Beyan Formu + Çoklu Dosya Desteği

- Başlık daima yüklenen Tarifname kaynağından alınır; beyan formu/e-posta başlığı çatışma oluşturmaz.
- Beyan formu tablo parser'ı rol bazlı çalışır; açıklama/not satırları kişi adı kabul edilmez, TCKN/VKN/adres/il-ilçe/telefon/e-posta aynı kişiye bağlanır.
- Bilgi kaynağı olarak PNG/JPG/JPEG/WEBP/TIF/TIFF/BMP yerel Tesseract OCR ile desteklenir; OpenAI/API kullanılmaz.
- Tarifname ve Şekiller alanları DOC/DOCX/PDF kabul eder.
- EPATS PDF üretiminde Word otomatik satır numaraları kaldırılır; font/stil/kenar boşluğu/header-footer ve teknik içerik korunur.
- İstem sayımı farklı Word numId'lerinin aynı abstract numbering şemasına bağlı olduğu dosyalarda da doğru sayım yapar.

## v5.4.47 / 2026-09-01 — Süreçler / AI'sız Patent-Faydalı Model Ön Kontrolü

- `Süreçler > Patent / Faydalı Model Başvurusu` artık **OpenAI/API çağrısı yapmaz**; `OPENAI_API_KEY` tanımlı olmasa da bu bölüm çalışır.
- Beyan formu, DOC/DOCX/PDF/TXT/MD, EML, MSG ve yapıştırılan metindeki başvuru verileri etiket/tablo/bağlam kurallarıyla yerel olarak çıkarılır.
- Hak sahibi, buluş sahibi, VKN/TCKN, ülke, il, adres, başvuru türü, DP/dosya referansı, buluş başlığı ve rüçhan bilgileri yalnız açık kaynak verisinden alınır; bulunmayan bilgi tahmin edilmez.
- Bir Word tablosunda aynı satırda birden fazla `etiket | değer` çifti bulunması desteklenir. Serbest e-postada yalnız açık `hak/başvuru sahibi` ve `buluş sahibi` ifadeleri kabul edilir.
- Tarifname yalnız buluş başlığını teyit etmek/eksikse başlığı almak için kullanılır; tarifnameden hak sahibi veya buluş sahibi çıkarılmaz.
- Eksik veya çelişkili zorunlu bilgi varsa EPATS geçiş kilidi kapalı kalır. T/İ/Ö/Ş belge kontrolleri ve PDF üretimi tamamen yerel yazılımla devam eder.
- Üretim modüllerindeki tarifname/görüş/araştırma yapay zekâ akışları değiştirilmemiştir.


## v5.4.46 / 2026-09-01 — Süreçler / Otomatik Başvuru Ön Kontrolü

- `Süreçler > Patent / Faydalı Model Başvurusu` ekranı elle veri girişi yerine belge/e-posta/yazı kaynaklarından otomatik bilgi çıkaracak şekilde yenilendi.
- Beyan formu, DOC/DOCX/PDF/TXT/MD ile EML ve Outlook MSG e-posta dosyaları desteklenir; ayrıca e-posta/yazı metni doğrudan yapıştırılabilir.
- Hak sahibi, buluş sahibi, buluş başlığı, adres, rüçhan ve bulunan diğer bilgiler kaynak adıyla birlikte zorunlu ön kontrolde gösterilir; eksik veya çelişkili bilgi varsa EPATS geçişi kilitli kalır.
- Tarifname Word dosyasındaki kırmızı/mavi şablon açıklamaları otomatik kaldırılır. Temiz belge Tarifname/İstemler/Özet PDF'lerine ayrılır.
- Ön kontrolde Tarifname gerçek PDF sayfa sayısıyla `T-n`, istemler gerçek istem adediyle `İ-n`, özet `Ö`, şekiller ise `Ş-n` olarak gösterilir.
- Hazırlanan her PDF ayrı ayrı kontrol için indirilebilir; eksikler tamamlanmadan nihai EPATS paketi verilmez.

## v5.4.45 / 2026-09-01 — Süreçler / Patent-Faydalı Model Başvurusu MVP

- Mevcut üretim seçenekleri korunarak üst çalışma seçimine **Süreçler** bölümü eklendi.
- İlk süreç olarak **Patent / Faydalı Model Başvurusu** veri giriş ekranı eklendi.
- Hak sahibi, buluş sahibi, rüçhan ve başvuru bilgileri yapılandırılmış olarak tutulur.
- Oluşturulan tarifname DOCX dosyası `İSTEMLER` ve `ÖZET` başlıklarına göre ayrılarak EPATS için Tarifname / İstemler / Özet PDF seti hazırlanır.
- Şekiller DOCX/PDF ise EPATS paketine dahil edilir.
- EPATS otomatik tarayıcı aktarımı sonraki aşama için ayrılmıştır; bu sürüm başvuru verisi ve belge paketleme altyapısını hazırlar.


Bu paket, mevcut Render/GitHub tabanlı **Patent Atölyesi** uygulamasının 01.09.2026 tarihli güncel tam sürümüdür.

## GitHub'a yükleme — ESKİ DOSYALARIN KALMAMASI ÖNEMLİ

ZIP'i açın ve içindeki dosyaların tamamını mevcut GitHub deposunun **ana dizinine** koyun. ZIP içeriğini doğrudan depo köküne yükleyin; ZIP içinde dış/üst klasör yoktur.

**GitHub web arayüzündeki `Add file → Upload files` işlemi yalnız aynı adlı dosyaları günceller; yeni pakette artık bulunmayan eski dosyaları SİLMEZ.** Bu nedenle önceki sürümde daha fazla dosya varsa, doğrudan upload sonrasında bu fazla dosyaların depoda kalması normaldir; GitHub bunları kendiliğinden silmez.

Bu paket **düz (flat) GitHub paketi** olarak hazırlanmıştır: test dosyaları da doğrudan ana dizindedir; `.gitignore`, `download`, `download (1)` veya `tests/` alt klasörü yoktur.

Bu pakette `REPO_FILE_MANIFEST.txt` bulunur. Güncellemeden sonra uygulamaya ait depo kökünü bu manifest ile karşılaştırın; manifestte olmayan eski Patent Atölyesi dosyalarını silin. Kendi özel `.github/` workflow dosyalarınız varsa bunlar manifest dışında olsa bile bilinçli olarak korunabilir.

**Yalnız GitHub web kullanıyorsanız:** yeni dosyaları ana dizine yükleyip commit edin; ardından manifestte bulunmayan eski uygulama dosyalarını GitHub'da `Delete file` ile kaldırıp ikinci commit yapın.

**Git kullanabiliyorsanız en temiz yöntem:** depoyu klonlayın, depo içinde `git rm -r .` çalıştırın, bu ZIP'in içindeki temiz dosya setini repo köküne kopyalayın, ardından `git add -A`, `git commit` ve `git push` yapın. Böylece önceki sürümden artık kullanılmayan takip edilen dosyalar kesin olarak silinir.

Depo kökünde en az şu dosyalar doğrudan görünmelidir:

- `app.py`
- `rules.py`
- `tarifname_update.py`
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

Görüş hazırlama başlangıcında tam dört çalışma türü vardır ve sıra sabittir: **Araştırma raporuna karşı**, **İnceleme raporuna karşı**, **EP araştırma raporu veya ofis aksiyon**, **Yurtdışı ofis aksiyon**. İlk iki tür Türkiye dosyaları içindir. Türkiye araştırma raporlarında ve EP araştırma raporlarında savunma kapsamına yalnız **X ve Y** kategorisindeki dokümanlar alınır; **A** kategorisi teknik arka plan kabul edilir. İnceleme ve ofis aksiyonlarında ise yalnız uzmanın gerekçede fiilen kullandığı dokümanlar esas alınır.

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
- Tarifname oluşturma arayüzünde `Mevcut/revize tarifname` alanı yoktur. Mevcut bir tarifnamenin değiştirilmesi ayrı `Tarifname düzenleme` iş akışında ele alınır; yeni tarifname oluşturma akışına karıştırılmaz.
- Şekiller seçimi literatür araştırmasından önce gösterilir.
- `TEKNİK ALAN` **iki paragraf** halinde kurulur. İlk paragraf yalnızca tek giriş cümlesidir. Sistem+yöntem istem yapısında ilk paragraf `Buluş, ... sistemi ve yöntemi ile ilgilidir.` şeklinde; yalnız sistemde `... sistemi ile ilgilidir.`, yalnız yöntemde `... yöntemi ile ilgilidir.` şeklinde biter. Ardından mutlaka yeni paragraf açılır; ikinci paragraf `Buluş, özellikle ...` ile başlar ve teknik alanın ayrıntısını verir. İkinci paragraf `Sistem ve yöntem...` gibi bir ifadeyle başlatılmaz.
- `ÖNCEKİ TEKNİK` içinde aynı anlatımın devamı olan `Özellikle`, `Bununla birlikte`, `Bu nedenle` gibi cümleler gereksiz yere ayrı paragraf yapılmaz.
- Türkçe tarifnamede patent literatürü paragraflarında doğrulanmış İngilizce başlık ve Türkçe karşılığı birlikte yazılır. Paragraf bağlayıcı taslak dilinde `Literatürde yapılan araştırmalar sonucu ... rastlanmıştır. Söz konusu başvuru/doküman ... ile ilgilidir. Ancak bahsedilen başvuruda/dokümanda ... ile ilgili bir emareye rastlanmamıştır.` yapısını izler; `Buluşta ise ...` biçiminde görüş/savunma dili kullanılmaz. İngilizce tarifnamede özgün İngilizce patent başlığı kullanılır.
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

Kural sürümü: `2026-09-01.v34`

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


## v5.4.16 — BBF kaynak şekilleri zorunlu kullanım kalite kapısı

BBF veya açık teknik müşteri kaynağında kullanılabilir teknik şekil, blok diyagramı veya yöntem akış diyagramı bulunuyorsa bu görseller artık yalnız “öncelikli” kabul edilmez; **nihai Şekiller Word çıktısında zorunlu kaynak şekiller** olarak tutulur. Model tarafından yeniden çizilen bir şema, özgün BBF şeklinin yerine geçemez. Kaynakta sistem şekli ve yöntem/akış şekli birlikte bulunuyorsa ikisi de kullanılır. Ayrıca yüklenen müşteri şekilleri BBF içindeki teknik şekilleri otomatik olarak devre dışı bırakmaz; kaynak görsellerle birlikte değerlendirilir ve bayt düzeyinde mükerrerler tekilleştirilir. Kullanıcı açıkça bir şekli hariç bırakmadıkça veya daha yeni/düzeltilmiş müşteri şekli aynı görselin yerini almadıkça kaynak teknik şekil atlanırsa Şekiller çıktısı başarısız sayılır.

## 2026-08-13.v6 — Kaynak şekil envanteri

Şekiller üretiminden önce kullanılabilir BBF teknik görselleri `source_figure_inventory` altında envantere alınır. Şekiller Word oluşturulmadan önce bu kaynak görsellerin seçim listesinde bulunduğu doğrulanır. Böylece kaynakta şekil varken yalnız yardımcı/model üretimi bir çizimin teslim edilmesi engellenir.


## v5.4.17 — Referans kimliği ve ana istem tanım sırası sert kalite kapısı

- `(N)` referansı artık yalnız referans listesindeki aynı unsur adı/çekimiyle kullanılabilir; kısaltma veya eş anlamlı ad numarayı taşıyamaz.
- Ana sistem isteminde her bullet için ilk-tanım sırası kod seviyesinde kontrol edilir; bir bullet birden fazla yeni/henüz tanımlanmamış referans kullanırsa çıktı engellenir.
- Referanssız elektronik işlem birimi gibi yazılım taşıyıcıları, henüz tanımlanmamış modülleri sayan ayrı bullet yapılmaz; ilgili modül tanımına gömülür.
- Salt `... ortamında çalışmaya uygun sistem` ve `yöntemin ... ortamında gerçekleştirilmesi` biçimindeki bağlam-only alt istemler engellenir.
- Sistem ve yöntem bağımlı istemlerinin semantik tekrar kalite kapısı korunur.

## v5.4.18 — Hiyerarşik yazılım taşıyıcısı ve şekil referans-seti tamlığı

- 181284 dosyasındaki kullanıcı düzeltmesi kural haline getirildi: aynı elektronik işlem birimi üzerinde aynı çalışma ilişkisine sahip ardışık yazılım modülleri, taşıyıcı ifadesi tekrar edilmeden ortak bir üst madde altında gerçek Word alt madde işaretleriyle sıralanabilir. Üst madde referans taşımaz; her alt madde bir yeni referanslı unsuru ilk-tanım sırasıyla tanımlar. `... çalışan ve;` yalnız bu hiyerarşik yapı için izin verilen noktalı virgül istisnasıdır.
- Şekillerde kaynak BBF/müşteri görseli esas olmaya devam eder. Ancak referans listesinde ayrı olan unsurlar aynı tek kutu/tek hedef altında `2-3` gibi birleştirilemez; aynı taşıyıcı içinde dahi ayrı kutucuk/callout/ok ile ayırt edilebilir gösterilir.
- Ayrı Şekiller çıktısı istenmişse REFERANS NUMARALARI bölümündeki tüm gerçek sistem unsurları nihai şekil setinde en az bir kez bulunmak zorundadır. Sistem+yöntem şekilleri hazırlanıyorsa yöntem referanslarının tamamı da akış şekillerinde kapsanır. Set bazında eksik referans kalite kapısını durdurur.
- v5.4.17'de bazı yeni istem kontrolleri yalnız yardımcı `validators.py` tarafında kalabildiği için aktif Streamlit kalite kapısıyla eşzamanlılık riski vardı. v5.4.18'de aynı kontroller `app.py` aktif üretim/Word kapısına da bağlandı ve hiyerarşik istem yapısı hem üretici hem doğrulayıcı tarafından destekleniyor.

## 2026-08-14.v9 — Çıktı Sonrası Üçlü Tarifname Kalite Kapısı

Tarifname oluşturma akışında Word üretimi artık işlemin sonu değildir. İndirme düğmesi açılmadan önce üç zorunlu kontrol yeniden çalışır:

1. **BBF/Kaynak tamlık kontrolü:** mandatory `technical_facts` maddelerinin tamamının karşılığı ve kanıt metni nihai Word içinde doğrulanır. Tek bir teknik bilgi eksikse çıktı geri çevrilir.
2. **Ana istem + alt istem kontrolü:** ürün/yöntem dili, ortak yazılım taşıyıcısı, ilk-tanım sırası, bağımlılık, semantik tekrar, gerçek ek sınırlama ve gereksiz alt istemler yeniden denetlenir. Sistem/cihaz alt istemlerinde `bulunmasıdır` ve diğer eylem/sonuç sonlandırmaları kabul edilmez.
3. **Referans kullanım kontrolü:** REFERANS NUMARALARI bölümündeki unsur adları BULUŞUN DETAYLI AÇIKLAMASI ve İSTEMLER içinde geçtiği her yerde doğru `(N)` işaretiyle kullanılmalıdır. Yöntem istemlerindeki gNodeB/modül/veritabanı/yığın/cihaz gibi unsur kullanımları da bu kontrole dahildir. Referans listesi yöntem satırlarında ise önceki kural gereği sistem `(1)`, `(2)` işaretleri gösterilmez.

Ayrıca ortak `elektronik işlem birimi üzerinde koşturulan yazılım` üst bullet'ı yalnız gerçekten yürütülebilir yazılım/modül/kontrolör/arayüz/yığınlar için kullanılır. Veritabanı, bellek veya salt veri yapısı kaynak açıkça yürütülebilir bir yazılım birimi olduğunu söylemiyorsa bu ortak grubun altına alınmaz.


## 2026-08-14.v10 — Beşli Son Kalite Kapısı

Tarifname Word üretildikten sonra indirme sunulmadan önce altı zorunlu kontrol çalışır: (1) BBF/kaynak teknik bilgi tamlığı, (2) buluşa ait teknik kaynak bilgisinin BULUŞUN DETAYLI AÇIKLAMASI bölümüne tam ve kaynak-sadık aktarımı, (3) ana istem ve alt istem kalite/tekrar/gereklilik kontrolü, (4) detaylı açıklama ve istemlerde referans numarası tamlığı, (5) `Tarifname_181176_template.docx` ile boşluk-hiza-sayfa kırılımı-numaralandırma-kapanış girintisi uygunluğu, (6) istemlerde belirsiz `unsur` placeholder'ı ve yöntem adımlarında salt isimle bitiş kontrolü. Yöntem adımları `... yapılması/edilmesi/aktarılması/belirlenmesi` gibi gerçek işlem fiilimsileriyle biter.

## 2026-08-14.v11 — Tam Tarifname Şablon Sadakati

Tarifname Word çıktısındaki 4. kalite kapısı artık seçili başlıkları kontrol eden dar bir kontrol değildir. Nihai `.docx`, `Tarifname_181176_template.docx` ile deterministik olarak karşılaştırılır. Kontrol; section/sayfa geometrisini, header ve footer parçalarını, PAGE alanlarının yalnız şablondaki üst konumda kalmasını, ana başlıkların paragraf biçimini, bölüm geçişlerindeki boşlukları, `BULUŞUN KISA AÇIKLAMASI` öncesindeki şablon `space-after` boşluğunu, `ŞEKİLLERİN KISA AÇIKLAMASI` öncesini, son şekil açıklaması ile `Çizimlerin...` paragrafı arasını, `REFERANS NUMARALARI` içindeki sistem/yöntem ayrımını, `BULUŞUN DETAYLI AÇIKLAMASI` öncesini, `İSTEMLER` öncesindeki iki boşluğu, istem kapanışlarını ve `ÖZET` bölümünün başlık/buluş adı/metin/son boşluk ritmini kapsar.

Tarifname üretiminde header/footer yeniden kurulmaz; bağlayıcı şablondan aynen korunur. Footer'a ayrıca sayfa numarası eklenmesi yasaktır. Sabit/dinamik paragraflar mümkün olduğunca şablondaki paragraf arketiplerinden kopyalanarak üretilir. Bu tam şablon kontrolü başarısızsa çıktı kullanıcıya sunulmaz.

## 2026-08-14.v12 — Ham Kaynak + SVG + İstem Türü Sertleştirmesi

- Otomatik istem seçiminde kaynakta açık sistem/modül ve yöntem dayanağı birlikte bulunuyorsa `Sistem ve yöntem` zorunludur; model önerisi bunlardan birini düşüremez.
- Türkçe bağımlı yöntem istemleri tek ek işlemde `işlem adımını içermesidir.`, birden fazla ek işlemde `işlem adımlarını içermesidir.` ile kapanır; eylem-sonu isim kapanışları kalite kapısından geçmez.
- BBF + ek teknik müşteri belgeleri ham metin pasajlarına deterministik olarak ayrılır. Her pasaj bir kez teknik fact'e bağlanır veya gerekçeli teknik-dışı olarak işaretlenir; teknik pasaj sessizce dışarıda bırakılamaz.
- ZIP/ek teknik belge akışı `.svg` müşteri şekillerini kabul eder. SVG, model/Word için rasterize edilse de özgün müşteri şekli olarak envanterde kalır. Kullanılabilir kaynak şekillerin tamamı nihai Şekiller dosyasına girmeden şekil kapısı geçmez.
- GitHub dağıtımı düz ana-dizin ZIP'idir; testler köktedir ve gereksiz `download`/`.gitignore` dosyaları pakete alınmaz.


## 2026-08-14.v13 — Formül + renkli şablon run + uzman “NASIL?” kapıları

- Kaynakta açık matematiksel bağıntı bulunduğunda tarifname Word çıktısında denklem düz metin olarak bırakılmaz. `formulas[].expression` gerçek Word OMML denklem nesnesi olarak oluşturulur; istem içinde açık bağıntı kullanılacaksa `[[EQ: ...]]` işaretleyicisi inline OMML denkleme çevrilir. Nihai `.docx` içinde beklenen ve gerçek denklem nesnesi sayıları karşılaştırılır.
- `Tarifname_181176_template.docx` içindeki giriş talimatı ile İSTEMLER altındaki üç sabit talimat paragrafının kırmızı/mavi kelime dağılımı run düzeyinde bağlayıcıdır. Sabit metnin tamamını tek kırmızı run'a yazıp mavi run'ları boş bırakmak artık tam şablon kontrolünden geçmez.
- Bağımsız sistem istemindeki yazılım/modül unsurları deterministik uzman “NASIL?” kontrolünden geçer. `X modülü (N), ... yapan bir modül` biçimindeki İngilizce claim sırası reddedilir; Türkçe istemde kaynak destekli teknik işlev/mekanizma önce, unsur adı ve referansı sonra yazılır. Sınıflandırma ve hesaplama modüllerinde kaynak kriter/ilişki açıklıyorsa yalnız sonuç fiili yeterli değildir.
- Bu kontroller yalnız prompt kuralı değildir: taslak validasyonu, Word üretimi ve çıktı sonrası kalite kapısında çalışır.


## v5.4.24 — Görüş şablon ve buluş basamağı çıktı kapısı

Görüş hazırlama akışında dil seçimi arayüzün bağlayıcı girdisidir. Başvuru sahibi raporda/tarifnamede güvenilir biçimde bulunamıyorsa arayüzde ayrıca girilebilir; kullanıcı girdisi aynen korunur ve metadata alanları boşsa Word çıktısı verilmez.

`Gorus_metni_696809_template.docx` görüş çıktısı için bağlayıcıdır. Kurum başlıkları, 3x3 metadata tablosu, metadata sonrası fiziksel boş paragraf, `Sayın Uzman,` girişi, kısa giriş paragrafı ve ardından fiziksel boş paragraf deterministik olarak denetlenir. D1/D2/D3 şekil tablolarından önce taslaktaki iki fiziksel boş paragraf korunur; şekiller modelce yeniden çizilmez, yüklenen özgün patent PDF'lerinden alınır.

Tarifname dayanaklarında model sayfa/satır numarası üretmez. Birebir alıntı önce tarifname metninde doğrulanır, ardından DOC/DOCX/PDF fiziksel sayfası render edilerek basılı satır numaraları üzerinden `Tarifname sayfa X, satır Y-Z’te bu durum şu şekilde belirtilmiştir: “...”` biçiminde deterministik atıf eklenir.

Buluş basamağı itirazında ana ikna bölümü, uzmanın gerekçede fiilen kullandığı doküman kapsamına göre kurulur. Tek D1 gerekçesi varsa tek-doküman genel değerlendirmesi, gerçek bir kombinasyon gerekçesi varsa dokümanların birlikte değerlendirilmesi kullanılır. Çıktı kapısı teknik fark → teknik katkı/teknik etki → objektif teknik problem → motivasyon/yönlendirme → gereken ilave yapısal/işlevsel değişiklikler → geriye dönük değerlendirme riskinin kurulmasını ve yeterli teknik derinliği zorunlu tutar. Raporun gerekçeli değerlendirmesinde fiilen kullanılan savunma dokümanları objektif incelenir; yalnız `ilgili dokümanlar` listesinde bulunup gerekçede kullanılmayan D-dokümanları görüşe taşınmaz. Raporda olmayan X/Y kategorisi uydurulmaz; mevcut istemlerle devam kararı verilmişse görüş içinde yeni istem revizyonu yapılmaz.

İndirme düğmesi ancak şu görüş kapıları geçtikten sonra açılır: metadata/kaynak doğruluğu, birebir tarifname alıntısı, fiziksel sayfa-satır doğrulaması, özgün D-şekilleri, tam görüş şablon sadakati, buluş basamağı birlikte değerlendirme derinliği ve Word→PDF render smoke testi.


## v5.4.25 — Görüş ham-kaynak ve ikinci okuma kalite kapısı

- İnceleme raporunda yalnız listelenen D1/D2/D3 ile uzmanın gerekçeli değerlendirmede fiilen kullandığı dokümanlar ayrılır. Arayüz yalnız savunmada gerekli dokümanı ister.
- Girişte doküman seçimi/usul anlatımı yapılmaz.
- Türkçe görüş anlatımında noktalı virgül kullanılmaz.
- Tarifname dayanağı desteklediği savunmanın aynı paragrafına bağlanır; `Tarifname sayfa ...` ayrı paragraf yapılmaz.
- Teknik fark → teknik katkı/etki → objektif teknik problem → motivasyon/yönlendirme → ilave değişiklik → hindsight zinciri ham kaynaklara karşı doğrulanır.
- Önceki teknik dokümanın gereksiz unsur referans numaraları görüş anlatımına taşınmaz.
- İlk taslak, rapor + tarifname + önceki görüş + savunma dokümanları + müşteri bilgileri + onaylı istem seti karşısında bağımsız ikinci okumadan geçer. Başarısızsa bir kez otomatik düzeltilir ve yeniden denetlenir.
- Word indirme öncesinde şablon, font/punto, 1,5 satır aralığı, fiziksel boşluk ritmi, özgün şekil, inline dayanak, noktalama, doküman kapsamı ve render kapıları görünür kalite raporuyla doğrulanır.


## v5.4.26 — EP görüş sekmesi, X/Y filtresi ve EP markup kalite kapısı

- Görüş başlangıcı `EP Araştırma Raporu` ve `Ofis Aksiyonu / İnceleme Raporu` olarak ayrıldı.
- EP araştırma raporunda yalnız X/Y kategorisi dokümanlar savunma kapsamına alınır. A kategorisi görüşe otomatik taşınmaz.
- EP İngilizce görüş girişi `Dear Sir/Madam` ve kullanıcı tarafından verilen EP giriş kalıbını, sonuç ise `In the light of above explanations and defence...` kalıbını kullanır.
- EP tarifname markup literatür eklerinde D1/D2 etiketleri tarifname gövdesinde kullanılmaz. `As a result of the research on the subject...` önceki teknik formatı ve komşu paragrafın font/punto/spacing özellikleri zorunlu olarak klonlanır.
- Article 84 antecedent düzeltmesinde belirsiz `the actor` ifadesi otomatik çoğullaştırılmaz. Tarifnamede açık referent aranır ve her değişiklik için fiziksel sayfa/satır dayanağı gösterilir.
- Bağımlı istem görüşü artık tüm itirazlı istemleri veya teknik grupları kapsayan teknik katkı kontrolünden geçer.
- Word kalite raporuna EP X/Y kapsamı, EP giriş/sonuç formatı, dependent-claim teknik katkısı, markup literatür etiketi ve font eşleşmesi kontrolleri eklendi.


## v5.4.27 — Minimum Track Changes ve EP önceki teknik fark kapısı

- EP tarifname literatür eklerinde D1/D2 etiketi kullanılmaz. Her X/Y paragrafı mevcut formatta `As a result of the research on the subject...` ile başlar, objektif doküman açıklamasından sonra `However,` ile başvurunun as-filed metninde zaten bulunan teknik farkı açıklar. Yeni özellik veya yeni teknik etki eklenmez.
- Claim markup minimum-fark mantığıyla üretilir: değişmeyen kelime/cümle parçası silinip yeniden eklenmez. `the→a` yalnız artikel, `actor→authenticated actor (8)` yalnız actor tokenı, eksik harf yalnız karakter insertion olarak işaretlenir.
- Word indirme kapısında minimum redline, EP However-fark dayanağı, X/Y kapsamı, Article 123(2) dayanağı ve font/punto eşleşmesi görünür kontrol satırlarıdır.


## v5.4.28 — Dört görüş modu ve son Markup sayfa/satır kapısı

- Ana ekrandaki sürüm, `APP_VERSION` ile birlikte **v5.4.28** olarak görünür; README başlığı da aynı sürümle senkron tutulur.
- Görüş hazırlama ana seçimi dört moddur: Türkiye Araştırma, Türkiye İnceleme, EP Araştırma/Ofis Aksiyonu, Yurtdışı Ofis Aksiyonu.
- Araştırma raporlarında yalnız X/Y dokümanları savunulur; A kategorisi savunma dokümanı değildir.
- Revizyon/Markup varsa bütün fiziksel sayfa-satır dayanakları son Markup dosyasından hesaplanır. Temiz sürüm veya ilk yüklenen tarifname bu iş için kullanılmaz.
- Word üretiminden hemen önce alıntı metni ve sayfa/satır konumu ikinci kez doğrulanır; eşleşme yoksa çıktı bloke edilir.


## v5.4.29 — DP otomatik çıktı adı + son ham-BBF ikinci okuma kapısı (21.08.2026)

- Tarifname oluşturma ekranında **DP referans numarası çıktı adının tek kaynağıdır**. Ayrı `Çıktı dosyasının adı` ve `Şekiller dosyasının adı` soruları kaldırıldı. Örneğin DP `181267` ise çıktılar otomatik `Tarifname_181267.docx` ve şekiller seçilmişse `Şekiller_181267.docx` olur. DP referansı boşsa üretim başlamaz.
- `technical_facts` listesine alınmış teknik bir bilgi artık `mandatory=false` ile kapsam dışına çıkarılamaz. Kaynaktaki bütün teknik facts, örnek senaryo/koşul/avantaj/teknik etki/alternatif dahil, nihai tarifnamede uygun bir bölümde korunmak zorundadır.
- Tarifname taslağı üretildikten sonra **bağımsız son ham kaynak ikinci okuması** çalışır. Bu tur önceki `source_coverage_map` beyanını kanıt saymaz; her technical ham pasajı ve her technical_fact'i kullanıcıya gidecek gerçek taslak metninden birebir alıntıyla yeniden kontrol eder. Eksik tek pasaj/fact varsa aynı tıklama içinde düzeltme turuna dönülür.
- Word üretildikten sonra `ham pasaj → technical_fact → source_coverage_map → nihai Word kanıtı` zinciri deterministik olarak tekrar doğrulanır. Sonrasında 6 kapı yeniden çalışır: **1/6 Ham kaynak/BBF tamlığı, 2/6 Detaylı Açıklama tam kaynak aktarımı, 3/6 Ana+alt istemler, 4/6 Referanslar, 5/6 Tam şablon, 6/6 Unsur/yöntem dili**. Bunlardan herhangi biri geçmezse indirme düğmesi açılmaz.
- Başarılı üretimde arayüz açıkça `Ham veri kontrolü yapıldı` mesajını verir ve kontrol edilen ham pasaj, teknik pasaj ve atomik teknik bilgi sayılarını gösterir. Bu mesaj yalnız gerçek kontroller başarıyla tamamlandığında gösterilir.
- Sistem+yöntem tarifnamesinde TEKNİK ALAN ilk cümlesi `Buluş, ... sistemi ve yöntemi ile ilgilidir.` kalıbıyla bitmek zorundadır.
- Türkçe patent literatürü paragrafı bağlayıcı taslaktaki `Ancak ... ile ilgili bir emareye rastlanmamıştır.` diliyle biter; `Buluşta ise ...` kullanımı kalite kapısında reddedilir.
- Son patent literatürü/önceki teknik paragrafı ile `Sonuçta yukarıda bahsedilen...` paragrafı arasında **tam bir fiziksel boş paragraf** zorunludur; bu boşluk Word tam şablon kapısında ayrıca doğrulanır.




## v5.4.34 — ÖNCEKİ TEKNİK derinlik ve Ekstra Kontroller bildirimi (26.08.2026)

- Tarifname üretiminde kaynakta 4+ önceki-teknik/problem fact varsa müşteri kaynaklı ÖNCEKİ TEKNİK gövdesi patent literatürü hariç en az 3 gelişmiş paragraf ve en az 2400 karakter olarak doğrulanır; son genel paragraf `Yukarıda belirtilen eksiklikler, ...` ile başlar.
- Son kalite durumu artık ham BBF ikinci okumasının yanında `prior_art` ve `draft_quality` PASS bilgisini de taşır.
- Ham BBF ikinci okuması + ÖNCEKİ TEKNİK kaynak/derinlik + tam taslak + nihai Word kapıları + formül/HOW + render kontrollerinin tamamı fiilen geçtiğinde arayüz sonunda **EKSTRA KONTROLLER YAPILDI** uyarısı gösterilir. Herhangi bir kontrol yapılmadıysa veya başarısızsa bu mesaj kesinlikle gösterilmez.
- Arayüz dışı kullanım için aynı şartı uygulayan `tarifname_extra_controls_completed(...)` ortak kapısı eklendi.

## v5.4.33 — Görüş revizyon sırası, dayanak bölümü ve Destek Patent markup yazarı (26.08.2026)

- Görüşte istem değişikliği kararı artık bağlayıcı olarak rapor, tarifname/istemler, gerekli X/Y veya gerekçede kullanılan önceki teknik dokümanlar, varsa önceki görüş ve müşteri bilgisinin birlikte analizinden sonra verilir.
- Kullanıcı revizyonu onayladıysa görüş metninde önce **İstemlerde Yapılan Değişiklikler ve Dayanakları** bölümü üretilir. Değişiklik gerekçeleri ve birebir tarifname dayanakları burada verilir, X/Y/D savunmaları daha sonra başlar.
- Bu değişiklik dayanaklarının sayfa/satır numaraları da revizyonlu dosyada son Markup fiziksel render'ından alınır.
- `bileşen/modül/mekanizma` gibi mevcut fonksiyonel taşıyıcılar sırf açıklık itirazı bulunduğu için otomatik silinmez. Açık dayanak varsa minimum değişiklikle korunur ve yalnız gerektiği kadar somutlaştırılır.
- Yöntem bağımlı istemlerinde sonuç-odaklı kapanışlar, kaynak desteklediğinde `... işlem adımını/adımlarını içermesidir` yöntem diline dönüştürülür. Yeni teknik bilgi eklenmez.
- Görüş ve tarifname düzenleme Track Changes yazarı **Destek Patent** olarak standardize edildi.
- Word görüş üreticisi, revizyon-dayanak bölümünü D1/D2/X/Y bölümlerinden önce yerleştirir ve birebir alıntıları mevcut sayfa/satır kalite kapısına dahil eder.

## v5.4.32 — Tarifname Düzenleme / müşteri revizyonu iş akışı (21.08.2026)

- Arayüze yeni ve bağımsız **Tarifname düzenleme** iş türü eklendi. Yeni tarifname oluşturma akışına karışmaz.
- Ana giriş müşteriye gönderilmiş son `.docx` tarifnamedir. Müşteri revizyon/soruları ayrı PDF/DOCX/DOC/TXT/MD/ZIP olarak yüklenebilir veya aynı Word içindeki **yorumlar ve Track Changes** doğrudan müşteri talebi olarak okunabilir.
- Aynı müşteri Word'ü kaynak olarak kullanıldığında mevcut müşteri Track Changes'i otomatik kabul edilmez: sistem önce yorum/değişiklikleri talep bağlamı olarak çıkarır, ardından müşteri insertions'larını reddedip deletions'larını geri getirerek temiz baz üretir ve kendi değerlendirdiği değişiklikleri yeni bir Destek Patent markup katmanı olarak uygular.
- Başvuru durumu zorunlu girdidir: **henüz başvuru yapılmadı / başvuru yapıldı / rüçhan başvurusu yapıldı; sonraki başvuru hazırlanıyor**. Post-filing veya rüçhan sonrası yalnız yeni müşteri bilgisine dayanan teknik ekleme otomatik uygulanmaz; new-matter/rüçhan etkisi için kullanıcı kararı gerekir.
- Müşterinin bütün talepleri ayrı karar matrisine alınır: `apply`, `partial`, `explain`, `clarification`, `figure_action`, `procedural_action`. İkinci bağımsız okuma `coverage_complete=true` vermeden çıktı üretilemez.
- Word revizyonları gerçek **OOXML Track Changes** ile ve **EN AZ DEĞİŞİKLİK** prensibiyle uygulanır. Değişmeyen ön/son kelimeler yeniden silinip eklenmez; mevcut run biçimleri, sayfa/section ölçüleri ve numaralandırma korunur.
- Yeni paragraf gerçekten gerekliyse mevcut paragraf komple yeniden yazılmaz; komşu paragraf biçimi kopyalanarak Track Changes içinde yeni paragraf eklenir.
- Teknik/hukuki olarak uygulanmayan veya stratejik açıklama gerektiren müşteri talepleri gerektiğinde gerçek **Word comment** ile açıklanabilir. Patent metninin gövdesine müşteri notu eklenmez.
- Şekiller bu modda otomatik yeniden çizilmez. Yüklenmiş mevcut şekiller tarifnameyle teknik uyum, okunabilirlik ve başvuru biçimi bakımından analiz edilir; gerekiyorsa hangi şeklin hangi nedenle güncellenmesi gerektiği ve editable Mermaid/DWG/Visio kaynağı talebi listelenir.
- Müşteriye gönderilecek **mail taslağı zorunlu çıktıdır**. Mail; yapılan ana değişiklikleri, uygulanmayan/kısmen uygulanan taleplerin nedenini, doğrudan soruların cevaplarını, açık stratejik konuları ve şekil taleplerini özetler.
- Kullanıcıya varsayılan Word çıktısı yalnız `<kaynak_adı>_markup.docx` dosyasıdır. Clean/accepted sürüm yalnız iç kalite kontrolünde oluşturulur; ayrıca istenmedikçe indirme olarak verilmez.
- Yeni çekirdek modül: `tarifname_update.py`. Yeni regresyon paketi: `test_v5432_tarifname_update.py`.

## v5.4.31 — Tarifname yazım biçimi ve önceki teknik sert kalite kapıları (21.08.2026)

- Türkçe referans unsur adları sentence-case zorunludur: yalnız ilk normal kelime büyük başlar; teknik kısaltmalar korunur. Word üretiminden önce bu biçim yalnız doğrulanmaz, unsur adının detaylı açıklama/istem/yöntem adımlarındaki eşleşmeleri de deterministik olarak normalize edilir.
- Detaylı açıklama ve istemlerde aynı unsurun Title Case yazımı kalite kapısında reddedilir.
- Türkçe buluş başlığında parantez içi İngilizce karşılık/kısaltma reddedilir; kaynak destekliyorsa daha genel teknik başlık tercih edilir.
- Aynı gruptaki alternatif kullanım örnekleri tek sürekli paragrafta oluşturulur.
- BBF'deki önceki-teknik/problem technical_facts'in tamamı özellikle ÖNCEKİ TEKNİK bölümünde kanıtlanır; kaynak ayrıntılıysa bölüm kısa iki paragrafa sıkıştırılamaz.
- Bu kontroller hem taslakta hem Word indirme öncesi kalite kapısında çalışır.


## v5.4.31 — Tarifname bağlayıcı biçim kapıları
- Türkçe buluş başlığı bağlayıcı Title Case biçimine normalize edilir; teknik kısaltmalar korunur, bağlaçlar küçük bırakılır.
- Patent literatürü `English title (Türkçe başlık)` biçiminde yazılır; `Türkçe karşılığı` meta-dili reddedilir.
- BULUŞUN KISA AÇIKLAMASI içindeki numarasız buluş tanımı ana istemin yalnız referans işaretleri çıkarılmış birebir kopyasıdır.
- Detaylı açıklamadaki bütün sistem unsurlarının temel tanımları tek sürekli paragrafta bulunur; modül zinciri ayrı paragraflara bölünemez.
- `bir gerçekleştirimde / bir gerçekleştirmede / buluşun bir gerçekleştirilmesinde` yasaktır; `Buluşun bir yapılanmasında` kullanılır.
- ÖNCEKİ TEKNİK müşteri problem kümeleri kısa özetlenemez; son genel paragraf `Yukarıda belirtilen eksiklikler, ...` ile bağlanır.
- Şekil kısa açıklamalarında referans/adım numarası aralıkları yazılmaz.
- Ayrı şekiller Word dosyasında üst PAGE / NUMPAGES sayacı Arial 11 ve kalın olmak zorundadır ve indirme öncesi doğrulanır.


## v5.4.36 — Tarifname Düzenleme istem görünürlüğü ve güvenli şekil revizyonu (26.08.2026)

- Müşterinin `istemlerde açıkça vurgulansın/görülsün` talebi, teknik içerik tarifnamede zaten destekleniyorsa `zaten var` diye kapatılmaz. Tam teknik ad + kısaltma görünürlüğü minimum Track Changes ile sağlanır; aynı müşteri maddesinde sayılan her test/işlev ayrı dayanak kontrolünden geçer.
- `figure_actions` kayıtları `safe_auto_edit`, `basis_source`, `basis_quote` ve `edit_instructions` alanlarını destekler.
- Kaynakla desteklenen ve hedef Şekil numarası açık olan sınırlı şekil değişiklikleri otomatik uygulanabilir. Revize şekil özgün şekille ikinci görsel kontrolden geçmezse reddedilir.
- Tarifname Düzenleme arayüzü, güvenli revizyon başarıyla tamamlandığında ayrıca `Revize Şekiller Word dosyasını indir` çıktısı verir; güvenli olmayan değişiklikler yalnız şekil aksiyonu olarak kalır.

- Aynı Word içindeki açık kırmızı müşteri revizyon notları yorum/Track Changes gibi revizyon kaynağı sayılır; baz metinden çıkarılır ve otomatik kabul edilmez. Eski `.doc` ana tarifnameler LibreOffice üzerinden `.docx` tabanına dönüştürülerek gerçek OOXML Markup akışına alınır.

## v5.4.36 / 2026-08-27.v26 — Ayrı yöntem şekli + şekil kalite kapısı + istem satır bütünlüğü

- Sistem/cihaz şekli ile yöntem akış şekli kesin olarak ayrıldı. Yöntem istemi ve yöntem işlem adımları mevcutsa, kaynakta zaten ayrı bir yöntem akış şekli bulunmadığı sürece sistem şekline `1001, 1002...` bindirmek yasaktır; ayrıca yöntem akış şekli oluşturulur.
- Otomatik yöntem şekli siyah-beyaz, boş dolgulu kutu-ok yapısındadır. Her işlem adımının referansı kendi kutusunun içinde yer alır; oklar işlem sırasını gösterir.
- BBF/ek teknik kaynak şekli zaten kullanılabilir siyah-beyaz çizgisel ise özgün şekil korunur. Kaynak teknik şekil renkli veya renk dolgulu ise teknik geometri ve referanslar korunarak siyah-beyaz çizgisel patent stiline dönüştürülür ve özgün/dönüştürülmüş görsel ikinci doğrulamadan geçer.
- Ayrı Şekiller Word dosyasındaki `PAGE / NUMPAGES` sayacı artık ortalı **Arial 11 kalın** biçimindedir; alanların kendisi ve `/` ayıracı bu biçimde deterministik olarak doğrulanır.
- Şekiller kalite kapısı final görsellerde maddi renk bulunmamasını, sistem referanslarının set bazında kapsanmasını, yöntem istemi varsa yöntem referanslarının ayrı yöntem/akış şeklinde bulunmasını ve sayfa sayacı biçimini kontrol eder. Ayrı şekiller seçilmişse bu kapı geçmeden `EKSTRA KONTROLLER YAPILDI` uyarısı verilemez.
- Türkçe numaralı istemlerde `... olup, özelliği;` geçişi Word satır sonunda parçalanmaması için non-breaking boşluklarla birlikte tutulur ve nihai Word kapısında kontrol edilir.
- İstem açıklık kuralı güçlendirildi: aynı veri üzerinde ardışık fiiller kullanılırken ikinci fiilin nesnesi belirsiz bırakılamaz; gerektiğinde `bahsedilen verileri` / `söz konusu verileri` biçiminde açık nesne bağı kurulur.



## v5.4.37 / 2026-08-27.v27 — Gerçek Arial sayaç kapısı + kısa istem son-satır kapısı

- Şekiller üst bilgisindeki `PAGE / NUMPAGES` alanlarında artık yalnız `/` ayırıcı run veya python-docx font özelliği kontrol edilmez. PAGE ve NUMPAGES alan sonuçlarının OOXML `w:rFonts` içindeki `ascii`, `hAnsi`, `eastAsia`, `cs` değerlerinin tamamı literal `Arial`; `w:sz/w:szCs=22`; `w:b/w:bCs=true` olmak zorundadır. Header paragrafı varsayılan run biçimi de Arial 11 kalın olarak açıkça yazılır.
- Şekiller Word dosyası ayrıca LibreOffice/PDF render kapısından geçirilir. Her sayfada görünür `1 / N`, `2 / N`... sayacı üst bölgede, 11 punto ve kalın olarak render edilmeden çıktı verilmez. Linux render ortamında Arial'ın Arimo veya Liberation Sans ile ikame edilmesi yalnız PDF QA için kabul edilir; DOCX içinde font adı yine Arial olmak zorundadır.
- Türkçe istemlerde `olup, özelliği;` yalnız önceki tek kelimeyle bağlanmaz. Geçişten önceki son en az beş kelime de non-breaking kuyruk olarak korunur; böylece `sistemi olup, özelliği;` gibi kısa ikinci satır oluşması engellenir.
- Tarifname PDF render kalite kapısı İSTEMLER bölgesindeki fiziksel satırları inceler; `olup, özelliği;` ile biten 1–4 kelimelik kısa/orphan son satır tespit edilirse Word çıktısı reddedilir.


## v5.4.38 / 2026-08-27.v28 — Bağımsız istem iki-fiziksel-satır preamble kapısı

- Türkçe bağımsız sistem/cihaz/ürün/yöntem istemi yalnız buluş adı + `olup, özelliği;` biçiminde kısa bırakılamaz. Preamble, kaynakta dayanaklı teknik bağlamı ve/veya temel işlevsel ilişkiyi içerir.
- Taslak kalite kapısı açıkça çok kısa bağımsız istem girişini reddeder ve mevcut otomatik düzeltme turuna geri gönderir.
- Nihai Word, LibreOffice ile PDF'e render edildikten sonra İSTEMLER bölümündeki bağımsız istem girişleri fiziksel satır bazında sayılır. `olup, özelliği;` öncesindeki gerçek preamble metni en az iki fiziksel satır oluşturmuyorsa çıktı bloke edilir.
- Manuel satır sonu, ekstra boşluk, gereksiz tekrar veya anlamsız dolgu iki-satır koşulunu sağlamış kabul edilmez. Non-breaking `olup, özelliği;` orphan kapısı ayrıca korunur; iki kontrol birbirinin alternatifi değildir.


## v5.4.39 / 2026-08-31.v29 — Tip 3 D1/D2 kalın kimlik + ön değerlendirme dili kapısı

- Tip 3 `2. DEĞERLENDİRME` girişinde nihai patent kimlikleri şablondaki gibi **`<yayın no> (D1) ve <yayın no> (D2)`** olarak kalın yazılır. Dinamik paragraf doldurma sırasında şablondaki mixed-run bold biçimi artık kaybedilemez.
- Word kalite kapısı giriş paragrafında D1 ve varsa D2 etiketlerinin gerçekten kalın run içinde olduğunu deterministik olarak doğrular.
- Yenilik ve buluş basamağı değerlendirmeleri ön araştırma niteliğine uygun ihtiyatlı dille yazılır: `... kriterini sağladığı düşünülmektedir` / `... kriterini sağlamadığı düşünülmektedir`. Kategorik `sağlamaktadır`, `sağlamamaktadır`, `sağlar/sağlamaz`, `sağlanır/sağlanmaz` sonuç dili değerlendirme alanlarında reddedilir.
- Aynı kurallar normal Tip 3 ve `Araştırma güncelleme - Tip 3` çıktılarına birlikte uygulanır.


## v5.4.40 / 2026-09-01.v30 — Gerçek bağımsız ham-BBF ikinci okuma + cümle içi başlık/unsur sentence-case kapısı

- Tarifname taslağı sonrasında yapılan SON HAM KAYNAK kontrolü artık önceki `source_passage_audit`, `technical_facts`, `source_coverage_map` veya `coverage_audit` verilerini ikinci okuyucuya vermez. Ham kaynak pasajlarının tamamı sıfırdan yeniden `technical/nontechnical` sınıflandırılır.
- Her ikinci-okuma satırında gerçek ham kaynak `source_quote`, sınıflandırma gerekçesi ve teknik pasaj için nihai taslaktan birebir evidence zorunludur. Audit; çalıştırmaya özgü nonce, ham-pasaj envanteri SHA-256 parmak izi ve nihai taslak SHA-256 parmak izi ile doğrulanır. Eksik/uyuşmayan meta, eski audit'in yeniden kullanılması veya önceki sınıflandırmanın kullanıldığını bildiren audit kapıdan geçmez.
- Nihai ekstra-kontrol bildirimi artık ayrı `independent_raw_second_read` boolean kapısına da bağlıdır. Bu kapı gerçekten PASS olmadan `EKSTRA KONTROLLER YAPILDI` yazılamaz.
- BBF'de buluş başlığı tamamı büyük harfli olsa bile detaylı açıklama girişinde normal sözcükler cümle-içi küçük harfe çevrilir; yalnız gerçek teknik kısaltmalar korunur. `AYARLANABİLİR ... SİSTEMİ` gibi kaynak başlığı cümle içine aynen büyük harfle taşınamaz.
- Unsur adları gövde içinde cümle-içi küçük harfle yazılır; fakat paragraf başında veya `.`, `?`, `!` sonrasında yeni cümle unsur adıyla başlıyorsa ilk normal kelime doğal büyük harfle başlamak zorundadır. Hem taslak hem nihai Word kapısı bu hatayı deterministik olarak reddeder.


## v5.4.41 / 2026-09-01.v31 — Bağımlı istem kısa giriş kuralı + Word çıktı kapısı

- Türkçe yöntem dışı bağımlı istemler yalnız `İstem X’e uygun sistem olup, özelliği;` kısa giriş kalıbıyla başlar. Buluş adı, cihaz/sistem alt türü veya başka tanımlayıcı ifade bağımlı istem girişinde tekrar edilmez.
- Türkçe yöntem bağımlı istemleri yalnız `İstem X’e uygun yöntem olup, özelliği;` kısa giriş kalıbıyla başlar.
- X, ek teknik özelliğin gerçekten dayandığı istem numarasıdır; zincir bağımlılık kuralları değişmez.
- Taslak kalite kapısı ve nihai Word kalite kapısı bu başlangıçları ayrı ayrı doğrular. `İstem 1’e uygun ayarlanabilir ... sistemi olup, özelliği;` gibi uzatılmış giriş varsa çıktı kullanıcıya açılmaz.


## v5.4.43 / 2026-09-01.v33 — Detaylı Açıklama tam kaynak aktarımı + gerçek BBF ikinci-okuma kapısı

- BBF ve açık teknik müşteri kaynaklarında **buluşun kendisini açıklayan bütün teknik bilgi**, yalnız tarifnamenin herhangi bir bölümünde bulunmakla yetinemez; `BULUŞUN DETAYLI AÇIKLAMASI` içinde de eksiksiz bulunur. Teknik alan/kullanım, teknik problem, çözüm, bütün unsurlar ve işlevleri, unsur ilişkileri, çalışma prensibi, teknik etkiler/avantajlar, alternatifler, örnekler, ölçü/değer/aralıklar, performans bilgileri ve teknik görselden çıkarılabilen buluş bilgisi bu kapsamdadır. Salt üçüncü kişi önceki-teknik/patent-literatürü bu zorunlu tekrara dahil değildir.
- Kaynak cümle teknik ve dilbilgisel olarak düzgünse **özetlenmez ve gereksiz yere yeniden yazılmaz**; mümkün olan en yüksek ölçüde kaynak cümle yapısı korunur. Yalnız dilbilgisi/noktalama, patent metni geçişi ve kanonik unsur adı/referans normalizasyonu yapılabilir. Kaynakta `eleman (1)` gibi geçici ifade varsa, nihai referans tablosundaki gerçek ad (ör. `solar spektrum kafası (1)`) kullanılır; bu normalizasyon yeni teknik bilgi ekleyemez.
- `AM1.5G`, `365–1000 nm`, `850 nm`, `PWM` gibi teknik literal/değer/kısaltmalar kaynakta varsa Detaylı Açıklama içinde deterministik olarak aranır; birinin düşmesi çıktı kapısını başarısız yapar.
- `source_coverage_map` içindeki her buluş-teknik fact için bölüm listesinde `BULUŞUN DETAYLI AÇIKLAMASI` ve bu bölümde gerçekten bulunan en az 20 karakterlik kanıt zorunludur. Modelin yalnız `covered=true` demesi yeterli değildir.
- Taslak sonrası bağımsız ham-BBF ikinci okuması önceki passage/fact/coverage kararlarını görmez. Buna rağmen ilk ve ikinci okuma teknik/teknik-dışı sınıflandırması çelişirse veya bir buluş-teknik pasaj ikinci okumada `detail_transfer_required=false` işaretlenirse çıktı **FAIL** olur ve kaynak çıkarımı yeniden yapılır. Nonce + kaynak SHA-256 + taslak SHA-256 bağlaması olmadan eski/uydurma audit kabul edilmez.
- Referans tablosunda açıkça tanımlanmış her sistem unsuru, BBF'deki `Yeni/Önceki` işaretinden bağımsız olarak istem setinde en az bir kez bulunur. Ana istem için zorunlu/farklılaştırıcı değilse uygun bir bağımlı istemde geri çekilme pozisyonu olarak kullanılabilir; `Yeni` kutusunun işaretli olmaması unsurun istemlerden sessizce atılmasına gerekçe değildir.
- Nihai çıktı kapıları artık altıdır: `source_completeness + detail_source_transfer + claims + references + template + element_step_language`. Bu altı kapı ve diğer ekstra kontroller gerçekten PASS olmadan `EKSTRA KONTROLLER YAPILDI` gösterilemez.


## v5.4.43 — Detaylı Açıklama yerleşim ve dil kapısı

Yeni tarifname üretiminde mevcut uygulama/önceki teknik ve teknik problem pasajları ÖNCEKİ TEKNİK bölümünde tutulur; Detaylı Açıklamaya zorla tekrar edilmez. Detaylı Açıklama sabit girişinden sonra ilk teknik paragraf bütün referanslı unsurları referans sırasıyla, kanonik unsur adı + numarası ve kaynak tanımıyla tek sürekli paragrafta açıklar. Sonraki paragraflarda kullanım, çözüm, teknik etkiler, unsur ilişkileri, çalışma prensibi, alternatifler, örnekler ve teknik değerler kaynak-sadık biçimde verilir. Gövde düzyazısında `Buluş;`/`Sistem;`/`Yöntem;`, gereksiz noktalı virgül ve Detaylı Açıklamada `uygundur` kullanımı çıktı kapısında reddedilir; buluşu kasteden `Sunulan çözüm/Bu çözüm` öznesi kanonik `Buluş/Sistem/Yöntem` diline çevrilir.


## v5.4.44 / 2026-09-01.v34 — Kayıpsız patent yeniden yazımı ve literal kayıp kapısı

- Detaylı Açıklama artık iki kuralı birlikte uygular: kaynak teknik bilgisinin **eksiksiz korunması** ve eski **patent yazım/çalışma-prensibi katmanının** korunması. Kaynak cümleleri sırayla kopyalamak yeterli değildir.
- İlk teknik paragraf bütün referanslı unsurların tanım paragrafıdır. Devamında unsurların teknik ilişkileri ve çalışma prensibi akıcı patent diliyle açıklanır. Kaynak alternatif/örnek/seçilebilir mod içeriyorsa uygun `Buluşun bir yapılanmasında, ...` dili zorunludur.
- `working_principle` boş/yetersiz olamaz; en az üç referanslı unsurun birlikte çalışma ilişkisi taslakta ve nihai Word'de doğrulanır.
- `AM1.5G`, `365–1000 nm`, `850 nm`, `PWM` gibi standart/kısaltma/değerler yalnız `technical_facts.statement` üzerinden değil doğrudan ham teknik passage üzerinden de çıkarılır. Detaylı Açıklamada bulunmayan literal, coverage beyanından bağımsız olarak indirmeyi bloke eder.
- Böylece BBF'deki teknik ayrıntılar kaybolmadan önceki patent dili, yapılanma ve çalışma-prensibi düzeni korunur.
