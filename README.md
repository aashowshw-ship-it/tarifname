# Patent Atölyesi v5.4.3

Bu paket, mevcut Render/GitHub tabanlı **Patent Atölyesi** uygulamasının 10.08.2026 tarihli güncel tam sürümüdür.

## GitHub'a yükleme

ZIP'i açın ve içindeki dosyaların tamamını mevcut GitHub deposunun **ana dizinindeki** dosyalarla değiştirin. `tarifname-main-v5.4.3-final` klasörünü ikinci bir alt klasör olarak yüklemeyin.

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

`render.yaml` varsayılan model değerini `gpt-5.6` olarak taşır. Hesabınızda farklı model adı kullanılıyorsa Render Environment ekranından değiştirin.

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
- Sistem ve yöntem istemleri birlikteyse başlık da uygun biçimde `... Sistemi ve Yöntemi` olur.
- `REFERANS NUMARALARI` unsur adlarında yalnızca ilk kelimenin ilk harfi büyük yazılır; teknik kısaltmalar korunabilir.
- Detaylı açıklamadaki yöntem adımları madde işaretli yazılır ve numara adımın sonunda `(1001)` biçiminde yer alır.
- Sistem/yöntem bağımsız istemlerindeki ayrı unsur ve adımlar madde işaretlidir.
- `İSTEMLER` ve `ÖZET` ayrı yeni sayfalardan başlar; başlıklar kalındır.
- `Buluşun bir gerçekleştirilmesinde` kullanılmaz; `Buluşun bir yapılanmasında` kullanılır.
- Önceki teknik bölümündeki müşteri anlatımı eksiksiz korunur; patent literatürü bunun yerine geçmez.
- Tarifname oluşturma arayüzünde `Mevcut/revize tarifname` alanı yoktur. Mevcut bir tarifnamenin değiştirilmesi ileride ayrı `Tarifname düzenleme` iş akışında ele alınacaktır; yeni tarifname oluşturma akışına karıştırılmaz.
- Şekiller seçimi literatür araştırmasından önce gösterilir.
- `TEKNİK ALAN` **iki paragraf** halinde kurulur. İlk paragraf yalnızca tek giriş cümlesidir ve `Buluş, ... ile ilgilidir.` biçiminde biter. Ardından mutlaka yeni paragraf açılır; ikinci paragraf `Buluş, özellikle ...` ile başlar ve teknik alanın ayrıntısını verir. İkinci paragraf `Sistem ve yöntem...` gibi bir ifadeyle başlatılmaz.
- `ÖNCEKİ TEKNİK` içinde aynı anlatımın devamı olan `Özellikle`, `Bununla birlikte`, `Bu nedenle` gibi cümleler gereksiz yere ayrı paragraf yapılmaz.
- Patent literatürü paragraflarında doğrulanmış İngilizce başlık ve Türkçe karşılığı birlikte yazılır.
- `BULUŞUN DETAYLI AÇIKLAMASI` içinde referanslı unsurlar tek tek ayrı paragraf yapılmaz; unsur açıklamaları tek sürekli paragrafta birleştirilir.
- Detaylı açıklamadaki yöntem işlem adımlarında ara maddeler virgülle, son madde noktayla biter. Bağımsız yöntem istemindeki madde işaretli işlem adımlarının **her biri virgülle** biter ve ardından `işlem adımlarını içermesidir.` yazılır; işlem adımları noktalamasız bırakılamaz.
- `Tarifname_181176_template.docx` fontların yanı sıra boş satır, 1,5 satır aralığı, gerçek Word madde işaretleri/numaralandırması ve istemler arası boşluk bakımından da birebir bağlayıcıdır.
- Amaç cümleleri `... karşılaştırmaktır.`, `... sağlamaktır.` gibi tam yüklemle biter; `... karşılaştırmak.` biçiminde bırakılmaz.
- REFERANS NUMARALARI bölümünden önce `(1)`, `(2)` gibi referans işaretleri kullanılmaz; kısa açıklamadaki ana istem özeti numarasız yazılır.
- Kaynakta adlandırılmış sistem modülleri bulunuyor ancak ayrı referans numarası verilmiyorsa, modüller kaynak sırasına göre `1, 2, 3...` olarak numaralandırılır. Yöntem işlem adımları ayrı bir referans ailesidir ve kullanıcı açıkça farklı bir sistem istemedikçe `1001, 1002, 1003...` olarak numaralandırılır; kaynakta işlem adımları `1, 2, 3...` veya `S101...` biçiminde verilmiş olsa dahi tarifname senkronizasyonunda `1001...` ailesine dönüştürülür.
- `REFERANS NUMARALARI` bölümünde önce sistem modülleri art arda yazılır, ardından **tek bir boş paragraf** bırakılarak yöntem işlem adımları `1001...` ailesiyle art arda yazılır.
- Sistem ile yöntem arasındaki `İşlem Adımı / Gerçekleştiren Unsur / Açıklama` ilişkisi tarifname gövdesinde açıklama tablosu olarak verilmez. Bu ilişki, modüllerin hangi işlem adımını hangi veri/çıktıyı kullanarak gerçekleştirdiğini açıklayan doğal teknik paragraf halinde yazılır. Yalnız kaynağın gerçek sayısal/deneysel veri tabloları gerektiğinde tablo olarak korunabilir.
- `ŞEKİLLERİN KISA AÇIKLAMASI` bölümünde `Şekil 1...`, `Şekil 2...`, `Şekil 3...` açıklamaları aralarında boş paragraf olmadan doğrudan alt alta yazılır.
- Kaynakta `UW`, `UW_F`, `UW_PL`, `UW_R`, `UW_M` gibi sembolik referanslar varsa sayısal unsur listesinden sonra bir boşlukla `UW. Kullanılabilir ağırlık` biçiminde yazılabilir; kaynakta gerçek referans olmayan `21-37` gibi geçici şekil numaraları uydurulmaz.
- İstemlerde teknik olarak ilişkili unsurlar birbirinin girdisi/çıktısı veya bağlantısı üzerinden yazılır; bağımsız unsur listesi gibi bırakılmaz.
- İstemlerde `HPU_W`, `FW_min`, `UW_F` gibi sembollerin teknik açılımı önce yazılır ve sembol parantez içinde gösterilir; formüllerde semboller korunur.
- Teknik açıklama metninde noktalı virgül gereksiz kullanılmaz; standart `olup, özelliği;` istem kalıbı istisnadır.

## Tip 3 ön araştırma

- Araştırma kesim tarihi kullanılır.
- Tam 10 doğrulanmış doküman belirlenir ve tek satır `TotalPatent arama sorgusu: ... or ...` üretilir.
- Kullanıcının bulduğu benzer dokümanlar nihai D1/D2 analizine eklenir.
- `Buluş basamağı var / Buluş basamağı yok / Otomatik belirle` seçimi korunur.
- D1 ve D2 tablolarında sol teknik özellik listesi birebir aynıdır. Sağ hücre yalnız `+` / `-` değildir; işaretin ardından özelliğin dokümanda geçtiği somut yer (`Özet`, `İstem`, `Şekil`, paragraf/sütun vb.) yazılır.
- D1/D2 şekilleri model tarafından çizilmez. Yalnız özgün patent şekli patent kaynağından indirilerek rapora eklenir; temin edilemezse yapay şekil üretmek yerine hata/uyarı verilir.
- Rapor gövdesinde `BBF` ifadesi, ok zincirleri veya `özellik + özellik` türü yapay kısaltmalar kullanılmaz.
- `On_Arastirma_Raporu_181612_template.docx` yalnız görünüm örneği değil, doğrudan doldurulan bağlayıcı Word şablonudur; gövdesi silinip yeniden kurulmaz.

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

Kural sürümü: `2026-08-10.v3`

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


### Şekiller - tarifname oluşturma
Şekil çıktısı müşteri görsellerini esas alır. Görseller `ŞEKİL 1`, `ŞEKİL 2` şeklinde sıralanır, başlık görselin altında bulunur ve her sayfanın üstünde `sayfa / toplam sayfa` göstergesi kullanılır. Sayfa sayısı sabit değildir. Şekildeki gerçek referans işaretleri tarifnamedeki `REFERANS NUMARALARI` ile senkron olmak zorundadır. Ayrıntılı şekil kuralları `rules.py` içinde arka planda uygulanır.
