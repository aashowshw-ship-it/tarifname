# Patent Atölyesi – Kayıtlı İş Kuralları

Kural sürümü: **2026-08-11.v2**

**BBF tamlık kontrolü görsel içeriği de kapsar:** gömülü teknik şekiller, grafikler, ısı haritaları, eksen/etiketler ve görsellerden açıkça çıkarılabilen teknik sonuçlar, metinsel içerikle birlikte eksiksiz değerlendirilir.

Bu dosya arayüzde kullanılan kuralların okunabilir özetidir. Uygulamanın çalıştırdığı tam metin `rules.py` içindedir.

## 1. Tarifname için değişmez temel

BBF'de bulunan bütün teknik bilgiler kullanılmalıdır. Önceki teknik açıklamaları, teknik problem, çözüm, unsurlar, işlevler, yöntem akışı, formüller, matematiksel ilişkiler, deneysel sonuçlar, tablolar, alternatif gerçekleştirmeler, kullanım senaryoları, şekil açıklamaları, referans tablosu ve teknik etkiler atlanamaz.

Yeni tarifname oluşturma akışında teknik kaynak BBF ve açıkça teknik kaynak olarak yüklenen müşteri belgeleridir. `Mevcut/revize tarifname` bu ekranda kullanılmaz; mevcut tarifnameyi değiştirme işlemi ileride ayrı `Tarifname düzenleme` iş akışında ele alınacaktır. Önceden hazırlanmış benzer tarifnameler yalnızca unsur ve istem kurgusunu görmek için kullanılır; bunların teknik içeriği yeni buluşa taşınmaz.

## 2. Her buluş aynı istem mantığında değildir

Önce şu ayrım yapılır:

- Zorunlu teknik çekirdek
- Zorunlu işlem sırası
- Paralel veya tekrarlanan işlem kolları
- Belirli gerçekleştirmelere ait ayrıntılar
- Alternatifler
- Çıktılar ve teknik etkiler

Buluş yalnızca yöntem olarak daha doğru korunuyorsa sistem istemi oluşturulmaz. Yalnızca sistem olarak daha doğruysa yöntem istemi oluşturulmaz. Her ikisinin de açık dayanağı varsa sistem ve yöntem istemleri birlikte hazırlanır.


## 2A. Tarifname dil, referans ve istem bağlantı kuralları

BULUŞUN KISA AÇIKLAMASI bölümündeki amaçlar tam yüklemle biter. `... karşılaştırmaktır.`, `... sağlamaktır.` gibi kullanım esastır; çıplak mastar (`... karşılaştırmak.`) bırakılmaz.

REFERANS NUMARALARI bölümünden önce `(1)`, `(2)`, `(1001)` gibi referans işaretleri kullanılmaz. Ana istem mantığı kısa açıklamada özetleniyorsa unsur adları numarasız yazılır. Kaynakta `UW`, `UW_F`, `UW_PL`, `UW_R`, `UW_M` gibi sembolik referanslar gerçekten kullanılıyorsa sayısal unsur listesinden sonra bir boş paragrafla `UW. Kullanılabilir ağırlık`, `UW_F. İlave yakıt tahsisi` biçiminde gösterilir. Kaynakta gerçek unsur referansı olmayan 21-37 gibi geçici şekil numaraları tarifnameye yeni referans numarası olarak taşınmaz.

Tarifname oluştururken müşterinin sistem/cihaz unsurları ve yöntem işlem adımları için verdiği referans işaretleri aynen korunur. Kaynakta `10, 20...`, `S101...`, `M1...` veya başka bir açık referans ailesi varsa sırf standartlaştırmak için değiştirilmez. Sistem modüllerine hiç referans verilmemişse kaynak sırasıyla `1, 2, 3...`; yöntem işlem adımlarına hiç referans verilmemişse `1001, 1002, 1003...` varsayılanı atanır. Kısmen numaralandırılmış kaynakta mevcut müşteri referansları korunur ve yalnız boş kalanlar çakışmayacak varsayılanlarla tamamlanır. `REFERANS NUMARALARI` bölümünde önce sistem modülleri art arda, sonra tek bir boş paragraf, ardından yöntem işlem adımları art arda yazılır.

Sistem ile yöntem arasındaki `İşlem Adımı / Gerçekleştiren Unsur / Açıklama` ilişkisi açıklama tablosu halinde tarifnameye taşınmaz. Bu ilişki doğal teknik paragraf halinde yazılır ve ilişkili modüllerin birbirinden aldığı veri/çıktı ile gerçekleştirdiği yöntem adımı açıkça bağlanır. Yalnız kaynaktaki gerçek sayısal veya deneysel veri tabloları gerektiğinde tablo olarak korunabilir.

`ŞEKİLLERİN KISA AÇIKLAMASI` bölümünde `Şekil 1...`, `Şekil 2...`, `Şekil 3...` açıklamaları arasında boş paragraf kullanılmaz; şekil açıklamaları doğrudan alt alta sıralanır.

Ana sistem istemi bağımsız unsur listesi olarak yazılmaz. Teknik olarak ilişkili sonraki unsur, kaynakta dayanağı bulunduğu ölçüde önceki unsurun verisini, çıktısını, hesap sonucunu veya bağlantısını açıkça kullanır. Kaynakta bulunmayan yapay bağlantı kurulmaz.

İstemlerde sembol tek başına anlam yerine geçirilmez. `kuru hibrit güç ünitesi ağırlığı (HPU_W)`, `asgari görev yakıtı ağırlığı (FW_min)`, `ilave yakıt tahsisi (UW_F)` gibi teknik açılım önce yazılır. Sembol matematiksel ilişkilerde aynen korunur.

Detaylı açıklamadaki yöntem işlem adımlarında ara maddeler virgülle, son madde noktayla biter. Türkçe bağımsız yöntem istemindeki ara işlem adımları virgülle, son işlem adımı ise noktalamasız biter; ardından ayrı satırda `işlem adımlarını içermesidir.` yazılır. Yöntem adımı referansları kaynakta verilmişse aynen korunur; kaynakta hiç referans yoksa `1001, 1002, 1003...` varsayılanı kullanılır. Aynı referanslı yöntem adımının teknik metni REFERANS NUMARALARI, detaylı açıklamadaki yöntem listesi ve bağımsız yöntem isteminde birebir aynı olmalıdır; istemde değişen adım aynı anda diğer iki yerde de değiştirilir.

Teknik açıklamada noktalı virgül gereksiz kullanılmaz. Virgül veya nokta tercih edilir. `olup, özelliği;` istem kalıbı bu kuralın istisnasıdır.


### Tarifname çıktı dili
Tarifname oluşturma ekranında `Türkçe` veya `İngilizce` seçilebilir. Dil seçimi yalnız çıktı dilini değiştirir; BBF/ek belge tamlığı, kaynak sadakati, istem mimarisi, donanımsal taşıyıcı, bağımlı istem stratejisi, referans senkronizasyonu, şekil kuralları ve `Tarifname_181176_template.docx` biçimi aynen korunur. İngilizce çıktıda TECHNICAL FIELD ilk paragrafı yalnız `The invention relates to ... .`, ikinci paragrafı `In particular, the invention relates to ... .` yapısındadır; bağımsız istemlerde `comprising:` ve bağımlı istemlerde `The system according to claim X, wherein...` / `The method according to claim X, wherein...` kalıpları kullanılır.

## 3. Paralel işlem adımları

Birinci, ikinci ve k'ıncı metrik için aynı tür analiz ayrı ayrı yapılıyorsa, ana istemde zorunlu çoklu yapı kapsayıcı biçimde yazılabilir. Ayrı analiz kolları ve bunların ayrı çıktıları aynı alt teknik akışa aitse tek bağımlı istemde birlikte sınırlandırılabilir.

Analiz işlemi ile analiz çıktısı aynı unsur değildir. Bununla birlikte teknik bütünlük gerektiriyorsa aynı bağımlı istem içinde beraber verilebilir.

Eğitim/genel aşamadaki paralel yapı ile test aşamasındaki paralel yapı aynı istem mantığıyla ele alınır, fakat aynı unsur sayılmaz. Aynı hesap farklı veri ve farklı aşama üzerinde yapılıyorsa ayrı referans numaraları ve ayrı işlevler korunur.

## 4. 1001–1006 türü tekrarlar

BBF'de aynı video/veri farklı aşamalarda yeniden giriş olarak gösteriliyorsa iki adım aynı metinle yazılmaz. Metin, ilk adımın tekil analiz girişi; sonraki adımın aynı veriyi çoklu analiz veya test aşamasına aktarması gibi aşama farkını açıkça göstermelidir.

## 5. Referans listesi ve yöntem adımı eşleşmesi

BBF referans tablosu eksiksiz çıkarılır. Her unsurun detaylı açıklamada karşılığı bulunur.

Şu üç bölüm birlikte kontrol edilir:

1. `REFERANS NUMARALARI`
2. `Yöntemin gerçekleştirdiği işlem adımları aşağıdaki gibidir:`
3. Yöntem istemleri

Aynı numara için kullanılan terminoloji ve işlev uyumlu olmalıdır. Kullanıcı birebir eşleşme istediğinde metinler birebir aynı tutulur.

Ana istemde kapsayıcı ve numarasız bir ifade kullanılabilir. Ayrıntılı numaralı adımlar bağımlı isteme taşınabilir. Bu işlem referans tablosundaki numaraların değiştirilmesini gerektirmez.

## 6. Ana ve bağımlı istem

Ana istem buluşun zorunlu teknik çekirdeğini açık, sıralı ve gereksiz tekrarsız kapsar. Unsurlar veri, sinyal, kontrol, işlem veya fiziksel bağlantı ilişkisi içinde kurulur.

Buluş ağırlıklı olarak yazılım, algoritma, modül veya veri işleme birimlerinden oluşuyorsa bağımsız istem soyut yazılım seviyesinde bırakılmaz. Kaynakta özel bir donanım zorunlu değilse geniş bir donanımsal taşıyıcı tercih edilir. Özellikle `bir elektronik cihaz üzerinde koşturulan yazılım vasıtasıyla ...`, `bir elektronik cihaz içerisinde çalışan yazılım tarafından ...` veya kaynakça uygun `elektronik işlem birimi tarafından ...` dili kullanılabilir. Sunucu, cep telefonu veya kişisel bilgisayar gibi daha dar taşıyıcılar ancak teknik kaynak gerçekten gerektiriyorsa yazılır. Özel donanım uydurulmaz. Aynı ilke yazılım ağırlıklı yöntem isteminde de uygulanır.

Bağımlı istemler ana istemi tekrar etmez. Yalnızca BBF'de dayanağı bulunan ve kapsamı gerçek anlamda daraltan teknik ayrıntıları ekler. `Önceki istemlerden herhangi birine uygun` kalıbı varsayılan değildir. Ek özellik ana istemde tanımlı bir modül veya işlem adımının ayrıntısıysa doğrudan ana isteme bağlanır. `İstem X veya Y’ye uygun` zinciri ancak yeni özellik gerçekten her iki bağımlılık yoluna da ihtiyaç duyuyorsa kullanılır; bağımlılık her zaman teknik olarak en doğrudan gerekli isteme kurulmalıdır. Kaynakta geçen her ayrıntı için ayrı bağımlı istem üretmek zorunlu değildir ve tercih edilmez. Yalnız koruma stratejisi bakımından anlamlı geri çekilme konumu sağlayan seçilmiş özellikler bağımlı isteme taşınır. Ana istemde elektronik cihaz/yazılım taşıyıcısı zaten kurulmuşsa alt istemlerde aynı taşıyıcı gereksiz yere tekrar edilmez.

Formüller zorunlu çekirdek değilse ana istemi gereksiz daraltmamak için bağımlı istemlerde kullanılabilir.

## 7. Tip 3 ön araştırma

- BBF'den teknik problem, unsurlar, işlevler, işlem adımları ve teknik etkiler çıkarılır.
- Global araştırma yapılır.
- Tam 10 doküman belirlenir.
- Tek satır halinde `TotalPatent arama sorgusu: ... or ...` üretilir.
- Kullanıcının benzer dokümanları sorulur ve nihai değerlendirmeye eklenir.
- Yeniliği tek başına bozan belge varsa D1 tek başına kullanılabilir.
- Aksi halde D1 ve tamamlayıcı D2 seçilir.
- Kullanıcı `Buluş basamağı var / Buluş basamağı yok / Otomatik belirle` seçimi yapabilir.
- D1 ve D2 tablolarındaki özellik listeleri aynı olmalıdır.
- Yalnızca `+` veya `-` kullanılır.
- Ön Araştırma Raporu_181612 şablonu korunur.

## 8. Son kalite kontrolü

Tarifname tamamlanmadan önce şu başlıklar birlikte kontrol edilir:

- BBF'deki bütün bilgilerin aktarılması
- Kullanıcının revizyonlarının korunması
- Referans tablosunun tamlığı
- Detaylı açıklama, yöntem adımları ve istemlerin uyumu
- Ana istemin buluşu gerçekten yansıtması
- Bağımlı istemlerin gerçek daraltma sağlaması
- Formül, tablo ve deneysel sonuçların korunması
- Aynı metinle yazılmış farklı yöntem adımı numaraları
- Eğitim/genel ve test aşamalarının doğru ayrılması

## 9. Görüş oluşturma – bağlayıcı şablon ve revizyon akışı

Görüş çalışmaları için tek bağlayıcı şablon `Gorus_metni_696809_template.docx` dosyasıdır. Eski görüş dosyaları içerik geçmişi olarak incelenebilir ancak şablon yerine kullanılamaz. Şablonun logo, header/footer, marj, sayfa geometrisi, başlık yerleşimi, font, punto ve paragraf düzeni korunur.

Başvuru bilgi alanlarında **Başvuru No**, **Başvuru Sahibi** ve **Referans** etiketleri kalın; değerleri normal yazıdır. Görüş gövdesindeki normal paragraflar iki yana yaslıdır.

İnceleme veya araştırma raporunda istem revizyonu gerekmiyorsa istemler sırf daha iyi yazılabilir diye değiştirilmez. Revizyon gerekiyorsa akış şöyledir: önce yalnızca gerekli istem değişiklikleri kullanıcıyla netleştirilir, en az değişiklik ve kapsam aşımı yapmama ilkeleriyle mutabakat sağlanır, kullanıcı açıkça onayladıktan sonra görüş hazırlanır. Görüş aşamasında onaylı istemlere kendiliğinden yeni değişiklik eklenmez.

İstem revizyonlarında yalnızca uzmanın belirttiği noktalar ve bunları gidermek için zorunlu ifadeler değiştirilir. Ürün/sistem istemlerinde yapısal unsur dili, yöntem istemlerinde yöntem dili kullanılır; ürün ve yöntem istemleri karıştırılmaz.

## 10. Çıktı dosya adı

İndirilen Word dosyalarının gerçek adı insan tarafından okunabilir biçimde kalmalıdır. `Görüş%20Metni_698891.docx` veya `G%C3%B6r%C3%BC%C5%9F...` gibi URL-kodlu adlar kullanıcıya indirme adı olarak verilmez. Uygulama `%20` ve diğer URL kodlarını çözerek örneğin `Görüş Metni_698891.docx` adını kullanır. Bu normalizasyon Tarifname, Görüş ve Tip 3 indirmelerinin tamamında ortak `safe_output_name()` fonksiyonu üzerinden uygulanır.

## 11. Tarifname dil ve sayfa düzeni – 07.08.2026 son düzeltmeleri

- Sistem ve yöntem istemleri birlikte hazırlanıyorsa buluş başlığı da buna göre **“... Sistemi ve Yöntemi”** biçiminde olmalıdır; yalnızca sistem başlığı bırakılmaz.
- Tarifnamenin kullanıcıya sunulan metninde **“BBF”**, **“buluş bildirim formu”** veya “BBF’de açıklandığı üzere” benzeri kaynak-form atıfları kullanılmaz. Kaynaktaki teknik bilgi doğrudan buluş anlatımı olarak yazılır.
- `REFERANS NUMARALARI` bölümünde unsur adları başlık biçiminde yazılmaz. Yalnızca ilk kelimenin ilk harfi büyük olur; SIM, IMEI, API gibi standart teknik kısaltmalar korunabilir. Aynı unsur adları cümle içinde cümle yapısına uygun küçük harfle kullanılır.
- Detaylı açıklamada `Yöntemin gerçekleştirdiği işlem adımları aşağıdaki gibidir:` ifadesinden sonra adımlar madde işaretli yazılır ve numara başta `1001.` biçiminde değil, adım metninin sonunda `(1001)` biçiminde verilir.
- Bağımsız sistem istemindeki unsurlar ile bağımsız yöntem istemindeki işlem adımları ayrı ayrı madde işaretli yazılır.
- Tarifname ve bölüm başlıkları kalındır. `İSTEMLER` yeni bir sayfadan, `ÖZET` ayrıca mutlaka yeni ve ayrı bir sayfadan başlar. İstem revizyonu/azaltması sırasında ÖZET sayfa geçişi kaybedilemez.
- **“Buluşun bir gerçekleştirilmesinde” kullanılmaz; “Buluşun bir yapılanmasında” kullanılır.**
- Önceki teknik bölümünde müşterinin kaynakta verdiği teknik arka plan, eksiklikler, problem anlatımı ve karşılaştırmalar eksiksiz korunur. Seçilen patent literatürü bunların yerine geçmez, yalnızca ayrı patent paragrafları olarak eklenir.
- Tarifname oluşturma arayüzünde `Mevcut/revize tarifname` alanı bulunmaz. Mevcut tarifnameyi değiştirme işi ayrı `Tarifname düzenleme` akışıdır ve şimdilik bu pakete eklenmemiştir.
- Şekiller seçimi tarifname akışında literatür araştırmasından önce gösterilir.
- `TEKNİK ALAN` **zorunlu olarak iki paragraf** halinde yazılır. İlk paragraf yalnızca tek giriş cümlesinden oluşur ve **“Buluş, ... ile ilgilidir.”** biçiminde biter. Sonra yeni paragraf açılır; ikinci paragraf mutlaka **“Buluş, özellikle ...”** ile başlar ve teknik alanı ayrıntılandırır. İkinci paragraf **“Sistem ve yöntem...”** gibi bir ifadeyle başlatılmaz.
- `ÖNCEKİ TEKNİK`te aynı teknik anlatımın devamı olan **“Özellikle...”**, **“Bununla birlikte...”**, **“Bu nedenle...”** gibi cümleler sırf yeni cümle başladığı için ayrı paragraf yapılmaz.
- Literatür araştırması sonucu eklenen her patent paragrafında **İngilizce başlık + Türkçe başlık karşılığı** birlikte verilir.
- `BULUŞUN DETAYLI AÇIKLAMASI`nda referanslı unsurlar tek tek ayrı paragraflara bölünmez; unsur açıklamaları tek ve sürekli bir paragrafta birleştirilir. Ayrı yapılanma/alternatif, yöntem listesi ve çalışma prensibi ayrı paragraf olabilir.
- `Yöntemin gerçekleştirdiği işlem adımları aşağıdaki gibidir:` sonrasında ara maddeler virgülle, son madde noktayla biter; maddeler noktalamasız bırakılmaz.
- Şekillerin kısa açıklamasında gerekli değilse yöntem adımı numara aralığı tekrarlanmaz; kısa ve işlevsel açıklama kullanılır.
- `Tarifname_181176_template.docx` font ve başlık kadar **boş paragraflar, 1,5 satır aralığı, otomatik istem numaralandırması, gerçek Word madde işaretleri, istemler arası boşluklar ve sayfa geçişleri** bakımından da bağlayıcıdır.

## 12. Görüş akışı – analiz, revizyon mutabakatı ve Markup

Görüş modülü artık tek düğmeyle doğrudan Word üretmez. İlk düğme **`1. Raporu analiz et`** düğmesidir. Bu aşamada rapor, inceleme dosyalarında önceki görüş, varsa müşteri bilgisi, tarifname ve X/Y dokümanları birlikte analiz edilir.

İlk analiz istem revizyonunun gerekli olup olmadığını açıkça belirler. Revizyon gerekmiyorsa bu sonuç arayüzde gösterilir ve mevcut istemlerle **`2. Görüş metnini oluştur`** düğmesi açılır.

Revizyon gerekiyorsa görüş henüz oluşturulmaz. Arayüzde her öneri için istem numarası, gerekçe, tarifname dayanağı, mevcut ifade ve önerilen ifade gösterilir. Kullanıcı isterse ek talimat girerek revizyon önerilerini yeniden analiz ettirebilir.

Kullanıcı revizyonları onaylarsa kaynak tarifname `.docx` olmak zorundadır. Uygulama iki ayrı Word dosyası üretir:

- `Düzenlenen_tarifname_track_changes_<referans>.docx`: gerçek OOXML Track Changes/Markup işaretleri içerir.
- `Düzenlenen_tarifname_temiz_<referans>.docx`: aynı revizyonların kabul edilmiş temiz halidir.

Track Changes değişiklikleri mümkün olan en küçük kelime/ifade düzeyinde yapılır; tüm istem paragrafı topluca silinip yeniden eklenmez. Kullanıcı revize istem setini son kez onaylamadan görüş Word dosyası üretilmez.

Kullanıcı revizyon önerisini gördükten sonra açıkça **mevcut istemlerle revizyonsuz devam etmeyi** de seçebilir. Bu seçim de açık kullanıcı kararı olarak kayda alınır ve görüş mevcut istem seti üzerinden hazırlanır.

## 13. Tip 3 rapor biçimi - 09.08.2026 kesinleştirmeleri

- `On_Arastirma_Raporu_181612_template.docx` doğrudan doldurulan bağlayıcı şablondur; rapor gövdesi sıfırdan yeniden kurulmaz.
- `2. DEĞERLENDİRME` bölümü, D1/D2 yerleşimi, karşılaştırma tabloları, buluş basamağı bölümü, sonuç, uyarılar ve ekler şablondaki sıra ve formatta kalır.
- D1/D2 karşılaştırma tablosunda sağ hücre çıplak `+` veya `-` içermez. İşaretin ardından özelliğin dokümanda nerede bulunduğu `Özet`, `İstem`, `Şekil`, paragraf/sütun/sayfa veya ilgili açıklama bölümüyle belirtilir. D1 ve D2'nin sol teknik özellik listesi birebir aynıdır.
- Patent şekilleri yapay zekâ ile üretilmez veya yeniden çizilmez. Yalnız özgün patent şekilleri resmi/public patent kaynağından ya da kullanıcı tarafından sunulan orijinal patent dosyasından alınır.
- Ön araştırma raporu gövdesinde `BBF`, `ilk BBF`, `ikinci BBF`, `buluş bildirim formu` ifadeleri kullanılmaz. Normal raporda `araştırma konusu`, güncelleme raporunda gerektiğinde `ilk araştırma konusu` ve `revize araştırma konusu` kullanılır.
- Rapor anlatımında `→`, `=>`, ok zinciri veya `özellik + özellik + özellik` gibi yapay zekâ çıktısı izlenimi veren kısa sembolik anlatım kullanılmaz.

## 14. Araştırma güncelleme - Tip 3

Arayüzde ayrı iş türüdür ve üç temel yükleme alanı vardır: **İlk BBF**, **Revize BBF**, **İlk Ön Araştırma Raporu**. DP referans numarası, çıktı dosya adı ve araştırma kesim tarihi ayrıca alınır.

İlk aşamada ilk ve revize araştırma konusu karşılaştırılır. Yalnız kelime farkı değil, gerçek teknik sınırlama, yeni unsur/işlev, veri işleme ilişkisi, teknik etki ve teknik problem değişiklikleri çıkarılır. İlk rapordaki D1/D2 ve olumsuzluk gerekçeleriyle birlikte her farkın teknik katkı oluşturup oluşturmadığı ekranda gösterilir.

İkinci aşamada revize ayırt edici özelliklere odaklı global patent araştırması yapılır. İlk rapordaki D1/D2 başlangıç noktasıdır; ilk raporda bulunmayan yeni yakın dokümanlar ekranda ayrıca gösterilir. Daha güçlü belge D1/D2 olabilir, diğerleri yardımcı doküman olarak kalır.

Yeni araştırmadan sonra sistem kendi teknik kanaatini açıkça yazar. Word hemen oluşturulmaz; kullanıcı **Buluş basamağı sağlanıyor** veya **Buluş basamağı sağlanmıyor** sonucunu seçer. Yenilik sonucu ise kaynaklara göre dürüst şekilde otomatik belirlenir.

Güncelleme Word çıktısı yeni bir rapor türü değildir. Aynı Tip 3 Ön Araştırma Raporu şablonu kullanılır; fark analizi arayüzde kalır, Word'e `Revizyon farkları`, `BBF farkları` gibi ek başlık taşınmaz. Yardımcı yeni doküman için şablonda olmayan ayrı D3/D4 bölümü açılmaz; gerekirse buluş basamağı paragrafında kullanılır.



## 15. Tarifname oluşturma - şekil oluşturma kesin kuralları (10.08.2026)

- Şekillerde **müşterinin sağladığı özgün teknik görsel önceliklidir**. Teknik kurgu, kutu-ok ilişkileri ve anlam sırf estetik için yeniden tasarlanmaz.
- Ayrı şekiller Word dosyasında görseller sırayla **ŞEKİL 1, ŞEKİL 2, ŞEKİL 3...** olarak adlandırılır; bu başlık ilgili görselin **altında**, ortalı ve kalın yer alır.
- Her şekiller sayfasının üstünde **`mevcut sayfa / toplam sayfa`** göstergesi bulunur (örn. `1 / 3`). Toplam sayfa sayısı hiçbir zaman sabitlenmez; şekil adedi, boyutu ve okunabilirliğe göre dinamik oluşur. Bir sayfaya uygun büyüklükte birden fazla şekil yerleştirilebilir.
- Nihai şekiller müşteriden alındıktan sonra şekillerdeki gerçek teknik referans işaretleri ile `REFERANS NUMARALARI` listesi çapraz kontrol edilir. Şekilde bulunan gerçek bir unsur/yöntem referansı tarifnamede karşılıksız bırakılamaz.
- `1`, `2`, `3` ve `S101`, `1001` gibi işaretler şekil, referans listesi, detaylı açıklama ve ilgili istemlerde aynı teknik karşılığı taşır.
- `UW`, `UW_F`, `UW_PL`, `UW_R`, `UW_M` gibi şekil üzerinde gerçek sembolik referans olarak kullanılan işaretler sayıya çevrilmez; referans listesinde `UW. Kullanılabilir ağırlık`, `UW_F. İlave yakıt tahsisi` mantığında gösterilir. Metin/istem içinde teknik ad önce, sembol parantez içinde kullanılır.
- Kaynaktaki 21-37 gibi yalnız geçici şekil numarası olup gerçek tarifname referans sistemine ait olmayan işaretler yeni unsur numarası olarak uydurulmaz. Teknik anlam kaybetmeden kaldırılabiliyorsa kaldırılır.
- Referans listesinde şekille ilişkili gerçek bir unsur bulunup şekil üzerinde işareti eksikse şekil tamamlanmış sayılmaz. Kaynaktan konumu açıksa yalnız eksik referans işareti eklenir; konum belirsizse uydurma yerleştirme yapılmaz.
- Patent şekillerindeki açıklama yazıları mümkün olduğunca azaltılır. Ancak müşterinin özgün şekli içindeki yazı/formül kaldırıldığında teknik anlam veya hesaplama ilişkisi kaybolacaksa müşteri görseli korunabilir. Bu durum bilinçli bir şekil-formalite riski olarak kabul edilir; teknik içerik formalite uğruna değiştirilmez.
- Şekil kalite kontrolünün zorunlu eşleşmeleri: **şekil ↔ REFERANS NUMARALARI ↔ detaylı açıklama ↔ istemler**, ayrıca yöntem şekillerinde **↔ yöntem adımları**.
