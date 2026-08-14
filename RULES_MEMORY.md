# Patent Atölyesi – Kayıtlı İş Kuralları

Kural sürümü: **2026-08-14.v12**

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

## 5A. İstemde “nasıl gerçekleştiriliyor?” açıklığı

İstem, yalnız hedeflenen sonucu veya fonksiyonu söylemekle bırakılmaz. Özellikle bağımsız istemde, kaynakta açık teknik dayanak bulunduğu ölçüde teknikte uzman kişinin “bu sonuç nasıl elde ediliyor?” sorusunun cevabı görülebilmelidir. Bu nedenle zorunlu teknik özellikler yazılırken uygun olduğunda işlemi gerçekleştiren teknik unsur/taşıyıcı, kullanılan girdi veya önceki unsurdan alınan veri, uygulanan teknik işlem/mekanizma ve üretilen teknik çıktı ya da sonraki unsurla ilişki birlikte kurulur.

“tespit eden”, “dönüştüren”, “optimize eden”, “sınıflandıran”, “oluşturan” gibi yalnız sonuç bildiren ifadeler, müşteri kaynağı teknik mekanizmayı açıklıyorsa tek başına bırakılmaz. Ancak bu kural ana isteme bütün uygulama ayrıntılarını doldurmak anlamına gelmez. Buluşun teknik katkısını ve çalışmasını anlamak için gerekli olan kaynak destekli “nasıl” özellikleri ana istemde, tercihli ayrıntılar ise bağımlı istemlerde veya detaylı açıklamada tutulur. Yazılım ağırlıklı buluşlarda elektronik cihaz/işlemci gibi donanımsal taşıyıcıya ek olarak yazılımın bu taşıyıcı üzerinde hangi veri ve teknik yapılar üzerinden hangi işlemi gerçekleştirdiği de istemden anlaşılmalıdır.

## 6. Ana ve bağımlı istem

Ana istem buluşun zorunlu teknik çekirdeğini açık, sıralı ve gereksiz tekrarsız kapsar. Unsurlar veri, sinyal, kontrol, işlem veya fiziksel bağlantı ilişkisi içinde kurulur.

Buluş ağırlıklı olarak yazılım, algoritma, modül veya veri işleme birimlerinden oluşuyorsa bağımsız istem soyut yazılım seviyesinde bırakılmaz. Kaynakta özel bir donanım zorunlu değilse geniş bir donanımsal taşıyıcı tercih edilir. Özellikle `bir elektronik cihaz üzerinde koşturulan yazılım vasıtasıyla ...`, `bir elektronik cihaz içerisinde çalışan yazılım tarafından ...` veya kaynakça uygun `elektronik işlem birimi tarafından ...` dili kullanılabilir. Sunucu, cep telefonu veya kişisel bilgisayar gibi daha dar taşıyıcılar ancak teknik kaynak gerçekten gerektiriyorsa yazılır. Özel donanım uydurulmaz. Aynı ilke yazılım ağırlıklı yöntem isteminde de uygulanır.

Bağımlı istemler ana istemi tekrar etmez. Yalnızca BBF'de dayanağı bulunan ve kapsamı gerçek anlamda daraltan teknik ayrıntıları ekler. `Önceki istemlerden herhangi birine uygun` kalıbı varsayılan değildir. Ek özellik ana istemde tanımlı bir modül veya işlem adımının ayrıntısıysa doğrudan ana isteme bağlanır. `İstem X veya Y’ye uygun` zinciri ancak yeni özellik gerçekten her iki bağımlılık yoluna da ihtiyaç duyuyorsa kullanılır; bağımlılık her zaman teknik olarak en doğrudan gerekli isteme kurulmalıdır. Kaynakta geçen her ayrıntı için ayrı bağımlı istem üretmek zorunlu değildir ve tercih edilmez. Yalnız koruma stratejisi bakımından anlamlı geri çekilme konumu sağlayan seçilmiş özellikler bağımlı isteme taşınır. Ana istemde elektronik cihaz/yazılım taşıyıcısı zaten kurulmuşsa alt istemlerde aynı taşıyıcı gereksiz yere tekrar edilmez.

Formüller zorunlu çekirdek değilse ana istemi gereksiz daraltmamak için bağımlı istemlerde kullanılabilir.

## 7. Tip 3 ön araştırma

- BBF'den teknik problem, unsurlar, işlevler, işlem adımları ve teknik etkiler çıkarılır.
- DP referansı alınır.
- Global araştırmada tam 10 doğrulanmış doküman belirlenir; tek satır `TotalPatent arama sorgusu: ... or ...` verilir ve aynı aşamada önerilen D1/D2 açıkça gösterilir.
- Bu üç çıktı verilmeden `Sizin araştırdığınız benzer dokümanlar var mı?` sorusuna geçilmez.
- Kullanıcı kendi dokümanlarını yüklerse ilk 10 liste değiştirilmez. Kullanıcı dokümanları arasından yalnız en ilgili birkaç belge ayrı `10+ XX... or YY...` satırında gösterilir.
- Kullanıcı dokümanlarının analizi sonrasında nihai D1/D2 belirlenir; `D1 değişti/değişmedi` ve `D2 değişti/değişmedi` açıkça yazılır, değiştiyse eski/yeni doküman numarası belirtilir.
- Yeniliği tek başına bozan belge varsa D1 tek başına kullanılabilir; aksi halde D1 ve tamamlayıcı D2 seçilir.
- Nihai D1/D2'den sonra sistem mutlaka kendi teknik kanaatini `Bence buluş basamağı var.` veya `Bence buluş basamağı yok.` şeklinde verir.
- Ancak sistem kanaatinden sonra kullanıcıdan `Buluş basamağı var / Buluş basamağı yok / Otomatik belirle` sonucu seçmesi istenir. Kullanıcı seçmeden Word raporu üretilmez.
- D1 ve D2 tablolarındaki özellik listeleri birebir aynı sıradadır. Sağ hücre `+` veya `-` ile başlar ve somut doküman dayanağı içerir; `±` kullanılmaz.
- Yalnız özgün patent şekli kullanılır; model/AI şekli üretilmez.
- `On_Arastirma_Raporu_181612_template.docx` bağlayıcıdır ve gövdesi yeniden kurulmaz.
- `Anahtar Kelimeler` yalnız İngilizce yazılır; şablondaki 5x2 iç tablo ve hücre düzeni değiştirilmez, en fazla 10 ifade kullanılır.
- `IPC Kodu` alanında kodlar kalın, açıklamalar normal yazıdır ve açıklamaların tamamı İngilizcedir. Türkçe IPC/CPC açıklaması kullanılmaz.
- `Kapsam` hücresi tam olarak `Global (İlan edilmiş olan patent başvuruları)` olarak kalır; araştırma kesim tarihi bu sabit alana eklenmez.
- D1/D2 `Özet` alanları patentin doğrulanmış özgün İngilizce Abstract metnidir. Model özeti, Türkçe çeviri veya yeniden yazılmış abstract kullanılmaz. Kullanıcının orijinal patent dosyasındaki İngilizce Abstract önceliklidir; bulunamazsa resmi patent kaynağı/Google Patents kullanılır. Özgün abstract doğrulanamıyorsa rapor tamamlanmış sayılmaz.
- Şablon sadakati yalnız genel görünüm değildir: sabit hücre metinleri, paragraf sayıları, kalın/normal run ayrımı, anahtar kelime tablosu geometrisi ve dört ayrı uyarı paragrafı da korunur. Word çıktıdan önce otomatik şablon-biçim doğrulaması yapılır.

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
- Anahtar kelimeler ve IPC/CPC açıklamaları İngilizce olmak zorundadır. Anahtar kelime alanı 5x2 şablon tablosunda kalır; IPC/CPC satırında kod kalın, İngilizce açıklama normal yazıdır.
- D1/D2 `Özet` alanına ilgili patentin özgün İngilizce Abstract metni doğrudan aktarılır; Türkçe özet/çeviri veya model tarafından yeniden yazılmış metin kullanılmaz.
- `Kapsam` sabit şablon metni değiştirilmez ve uyarı bölümü dört ayrı paragraf olarak korunur. Section, marj, tablo sayısı, sabit alan, IPC biçimi ve paragraf geometrisi çıktıdan önce otomatik denetlenir.

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
- Eksik referans işareti eklenmeden veya mevcut ok düzeltilmeden önce **referans işareti → unsur adı → detaylı açıklamadaki teknik tanım/işlev → şekil üzerindeki fiziksel karşılık** eşleştirmesi yapılır. Mevcut numara/ok tek başına doğru kabul edilmez.
- Kılavuz çizgisi/ok ucu doğrudan ilgili fiziksel unsurda sonlanır; boş alanı, komşu parçayı veya genel tertibatı gösteremez. Alt parçaya ait referans tüm tertibatı gösteremez. Örneğin `9 = Travers` ise ok traversin kendisine, `1 = Topuz` ise topuzun kendisine yönelmelidir.
- Her görünür parça zorla numaralandırılmaz. Yalnız tarifnamede gerçek referans işaretiyle tanımlı ve ilgili şekilde fiziksel karşılığı güvenilir biçimde görülebilen unsur işaretlenir. Görünür ve referanslı bir unsur numarasız bırakılmışsa güvenilir konum tespitinde eksik numara/ok eklenir; konum belirsizse uydurma işaretleme yapılmaz.
- Mevcut referans numarası doğru fakat oku yanlış parçaya gidiyorsa şekil doğru kabul edilmez; güvenilir eşleştirmede yalnız referans numarası/kılavuz çizgisi düzeltilir. Müşteri geometrisi, kesit taraması, perspektif, parça biçimi ve teknik kurgu değiştirilemez.
- Otomatik düzeltmeden sonra özgün ve revize görsel ikinci kez karşılaştırılır. Teknik geometri korunmamışsa, yanlış/eksik/fazla referans kalmışsa veya doğrulama güvenilir değilse revize görsel kullanılmaz ve şekiller Word çıktısı oluşturulmaz.


## 2026-08-13.v2 — Tek Tuş Tarifname Kalite Kapısı (Bağlayıcı)

Bu sürümde tarifname üretimi için aşağıdaki kurallar yalnız prompt tavsiyesi değil, üretim öncesi kalite kapısının parçasıdır:

1. Teknik gerçeklik yalnız BBF ve açık teknik ek kaynaklardan alınır; kullanıcıya görünen metinde BBF/müşteri çizimi/ek belge atfı yapılmaz.
2. BBF/BOM referansları kaynak envanteri olarak çıkarılır; “Diğer parçalar/Diğer elemanlar” gibi belirsiz üst başlıklar nihai patent unsuru olmaz.
3. Referans yalnız gerçek ve açık teknik unsura verilir; yapıştırıcı, malzeme, kaplama gibi özellikler gerektiğinde numarasız kullanılabilir.
4. Unsur adı gereksiz özel uygulamaya kilitlenmez; kaynak destekliyorsa O-ring örneği “Sızdırmazlık elemanı” gibi genel teknik unsur altında açıklanır.
5. Ana istemde ilk/ana unsur henüz tanımlanmamış sonraki referanslar kullanılmadan tanımlanır; yeni unsurlar teknik sırayla daha önce tanımlanan unsurlara bağlanır.
6. Kural olarak her ana istem bullet'ı tek yeni referanslı unsur tanımlar.
7. Sistem/cihaz/ürün/tertibat/yapılanma aynı ürün istem dil ailesidir. Yöntem dışındaki istemlerde işlem isimleştirmesi kullanılmaz.
8. “somun flanşının gövdeye bağlanması” değil “gövdeye bağlanan somun flanşı” gibi unsur merkezli dil kullanılır.
9. Yöntem dışındaki bağımlı istemler “olmasıdır.” veya “içermesidir.” ile biter.
10. Ana istemde uzman “nasıl?” sorusuna cevap bulmalıdır: zorunlu unsurun temel ilişkisi ve gerekli işlevi kaynakta dayanaklı biçimde görünür olmalıdır.
11. Unsur yalnız konumuyla bırakılmaz; gerekli ve kaynakta destekli teknik işlevi de açıklanır.
12. Aynı olmayan fiziksel unsurlar “ve/veya” ile tek unsur gibi bulanıklaştırılmaz; gerçek alternatifler açık kurulur.
13. Ana istem buluşun farklılaştırıcı zorunlu çekirdeğini taşımalıdır.
14. “vidalanan/kaynaklanan/yapıştırılan” gibi daraltıcı mekanizmalar ancak zorunlu teknik çekirdek veya farklılaştırıcı mekanizmaysa ana istemde tutulur.
15. Ana istemde aynı özellik farklı bullet'larda tekrar edilmez.
16. Bağımlı istem semantik tekrar yapamaz; her alt istem gerçek yeni sınırlama/geri çekilme pozisyonu sağlar.
17. İstem silme/birleştirme sonrası bağımlılık numaraları yeniden doğrulanır.
18. Örnek ölçü, çap, diş standardı ve ebatlar zorunlu değilse istemlere taşınmaz; detaylı açıklamada örnek olarak korunur ve kaynak destekliyorsa farklı ölçülere uygulanabilirlik açıklanır.
19. Noktalı virgül yalnız standart “olup, özelliği;” kalıbında kullanılır.
20. Özet tek paragraf ve tek cümledir; özet içindeki buluş adı kalın ve ortalıdır.
21. `Tarifname_181176_template.docx` paragraf/boşluk/numaralandırma/page-break bakımından bağlayıcıdır.
22. Şekil açıklamaları başlıktan sonraki şablon boşluğu korunarak aralarında boş paragraf olmadan sıralanır; son şekil açıklamasından sonra şablon boşluğu ve “Çizimlerin...” paragrafı gelir.
23. İstem numaraları gerçek Word otomatik numaralandırmasıyla oluşturulur; istemler arası boşluk şablondan korunur.
24. Üretim akışı taslak → AI kalite turu → yerel kalite kapısı şeklindedir. Yerel kalite kapısı hata verirse aynı kullanıcı tıklaması içinde hata modele geri beslenir ve en fazla iki ek otomatik düzeltme turu yapılır.
25. Doğrulama geçmeden Word oluşturulmaz. Word üretildikten sonra başlık, şekil açıklaması boşlukları, referans geçişi, istem otomatik numaralandırması, ÖZET page-break'i ve kalın özet başlığı programatik olarak doğrulanır. LibreOffice ile PDF render smoke-test geçmeden dosya sunulmaz.
26. Şekil referans/ok kuralları v5.4.11'deki haliyle aynen korunur; bu sürüm şekil kuralını değiştirmez.


## 2026-08-13.v4 — BBF Atomik Tamlık Kapısı ve Yazılım Taşıyıcı Doğrulaması (Bağlayıcı)

- Tarifname oluşturmanın birinci ve en üst kalite kuralı BBF ve ek teknik kaynaklardaki **tüm teknik bilgilerin uygun yerde eksiksiz korunmasıdır**. Bu kural istem kapsamından bağımsızdır; isteme taşınmayan teknik bilgi teknik alan, önceki teknik, kısa açıklama, detaylı açıklama, çalışma prensibi, alternatif yapılanma veya özet içinde korunmalıdır.
- Kaynaklar atomik `technical_facts` maddelerine ayrılır. Teknik avantaj, teknik etki, kullanım koşulu, ayırt edici yön, bağımsızlık sonucu, performans sonucu ve görsel/akış bilgisi ayrı fact olarak tutulur.
- Kişi adı, sicil, ödül payı, imza, form talimatı, boş idari alan, proje/idari alan ve yalnız patent araştırması anahtar kelimeleri teknik fact değildir; ancak bu alan içinde gerçek teknik açıklama bulunursa o açıklama teknik fact olarak alınır.
- Nihai taslakta her mandatory technical fact için `source_coverage_map` kaydı zorunludur: fact_id + covered=true + en az bir bölüm + tarifnameden gerçek kanıt metni. Eksik tek fact varsa Word oluşturulmaz ve otomatik düzeltme turu çalışır.
- Yazılım/modül ağırlıklı istemlerde yalnız donanım/işlemci kelimesi yeterli değildir. Yazılım/modülün elektronik cihaz, işlem birimi veya kaynakta verilen özel donanım üzerinde çalıştığı/koşturulduğu açık teknik ilişkiyle yazılmalıdır. Kaynak özel taşıyıcı veriyorsa (örn. SIM/eSIM üzerindeki güvenli işlemci, bellek ve izole çalışma ortamı) bu taşıyıcı korunur.
- BULUŞUN DETAYLI AÇIKLAMASI giriş cümlesinde buluş başlığı cümle içinde normal küçük harf düzeninde yazılır; başlık biçimi cümle içine kopyalanmaz. SIM, eSIM, API, NFC gibi teknik kısaltmalar korunur.
- Arayüzde Word üretimi öncesi BBF teknik bilgi kapsam paneli gösterilir. Panel yalnız bilgi amaçlı değil, validator tarafından geçirilmiş coverage map'i gösterir.


## 2026-08-13.v5 — Referans listesi ve alt istem semantik kalite kapısı
- `İSTEMLER` ve `ÖZET` DOCX yapısında `page_break_before=True` ile ayrı sayfadan başlatılır ve doğrulayıcı her iki başlık için bunu zorunlu olarak kontrol eder.
- `REFERANS NUMARALARI` altındaki yöntem işlem adımları `1001. ...`, `1002. ...` biçiminde önde yöntem numarasıyla yazılır. Bu satırlarda sistem/cihaz unsur işaretleri `(1)`, `(2)` vb. bulunmaz; parantezli unsur referansları `BULUŞUN DETAYLI AÇIKLAMASI` bölümünden itibaren kullanılır.
- Referans listesi ↔ detaylı açıklama ↔ yöntem istemi senkronizasyonunda referans listesinden bilinçli olarak çıkarılan sistem/cihaz parantez işaretleri karşılaştırma dışı, teknik kelimeler ve işlem ilişkileri karşılaştırma dahilidir.
- Sistem ve yöntem bağımlı istemlerinin tümü semantik tekrar kontrolüne tabidir; her alt istem gerçek ek teknik sınırlama veya stratejik geri çekilme konumu getirmelidir.
- Sistem şekillerinde görünüşte temsil edilen zorunlu ana taşıyıcı referans atlanmaz. Ok uçları küçük ve tutarlı tutulur. Yöntem akışına kaynak açıkça işlem-adımı döngüsü vermedikçe son adımdan önceki adıma geri dönüş oku eklenmez.


## 2026-08-13.v6 — BBF Özgün Teknik Şekli Zorunlu Kullanım Kuralı (Bağlayıcı)

- BBF veya açık teknik müşteri kaynağında kullanılabilir özgün teknik şekil/şema/akış diyagramı varsa nihai Şekiller Word çıktısında bu kaynak görsel **zorunlu olarak** kullanılır. “Öncelik ver” ifadesi artık yeterli değildir; özgün görsel model üretimi veya yeniden çizilmiş şema ile ikame edilemez.
- Kaynakta sistem/yapılanma şekli ile yöntem/algoritma akış diyagramı birlikte varsa ikisi de kullanılır. Biri diğerinin yerine geçmez.
- Yardımcı/ek şekil ancak kaynakta açık dayanağı bulunan teknik ilişkiyi açıklamak için eklenebilir; kaynak şekli kaldırmaz.
- Kullanıcı açıkça kaynak şeklin kullanılmamasını istemedikçe veya daha sonra verilen düzeltilmiş müşteri şeklinin aynı şeklin yerine geçtiği açık olmadıkça kaynak şekil atlanamaz.
- Şekillerden önce `source_figure_inventory` oluşturulur; kullanılabilir her kaynak teknik görsel nihai seçime alınmış olmalıdır. Eksik kaynak şekil varsa Şekiller Word çıktısı kalite kapısından geçmez.


## 2026-08-13.v7 — Referans Unsur Kimliği, İlk Tanım Sırası ve Alt İstem Ortam Kuralı (Bağlayıcı)

- Parantezli unsur numarası her kullanımda REFERANS NUMARALARI listesindeki aynı unsur adı veya yalnız dilbilgisel çekimiyle birlikte kullanılır. `1 = İnsansız hava aracı` ise `İHA (1)` ve `İHA’dan (1)` yasaktır; `insansız hava aracı (1)` ve `insansız hava aracından (1)` kullanılır. Kısaltma numarasız kullanılabilir.
- Ana istemde bir madde henüz tanımlanmamış birden fazla referanslı unsuru aynı anda kullanamaz. Kural olarak her madde tek yeni referanslı unsur tanımlar ve sonraki unsurlar yalnız daha önce tanımlananlara bağlanır.
- Kaynakta referanssız olan elektronik işlem birimi gibi yazılım taşıyıcısı, 2/3/4/5 gibi henüz tanımlanmamış modülleri topluca sayan ayrı unsur maddesi yapılmaz. Taşıyıcı ilişkisi ilgili modül ilk tanımlanırken aynı madde içinde kurulur.
- Bağımlı sistem istemi yalnız `sistemin ... ortamında çalışmaya uygun sistem olmasıdır` biçiminde kurulamaz. Çalışma ortamı ilgili baz istasyonu/arayüz/iletişim birimi vb. somut teknik unsurun niteliği veya bağlantısına dönüştürülür. Bağımlı yöntem isteminde de yalnız gerçekleştirme ortamı yazılmaz; gerçek bir işlem adımı, girdi veya teknik taşıyıcı ile ilişkilendirilir.
- Alt istemlerin her biri ana/üst isteme karşı semantik tekrar kontrolünden geçmeye devam eder; yalnız farklı kelime kullanılması yeni sınırlama sayılmaz.

## 2026-08-13.v8 — Ortak taşıyıcı istem grubu ve şekil tamlık kapısı

- Aynı elektronik işlem birimi/elektronik cihaz üzerinde aynı şekilde koşturulan birden fazla ardışık yazılım modülü varsa, teknik taşıyıcıyı her modülde tekrar etmek yerine ana istemde numarasız bir üst bullet `bir elektronik işlem birimi üzerinde koşturulan yazılım vasıtasıyla çalışan ve;` biçiminde kurulabilir; altında her referanslı modül ayrı gerçek Word alt bullet olarak, ilk-tanım sırasıyla yazılır. Bu `ve;` kullanımı yalnız bu hiyerarşik grup için noktalı virgül istisnasıdır.
- Ortak taşıyıcı üst bullet hiçbir referans numarası taşımaz. Her alt bullet kural olarak tek yeni referanslı unsur tanımlar; önce tanımlanmamış sonraki referansı kullanamaz. Taşıyıcı kaynakta ayrı referanslı unsur ise bu gruplama kullanılmaz.
- BBF/müşteri şekli temel şekildir; ancak referans listesinde ayrı olan iki unsur kaynak şekil üzerinde `2-3` gibi tek bir hedefte birleştirilmişse nihai patent şekli bunu ayrı kutucuk/çağrı/oklarla ayırır. Ortak taşıyıcı korunabilir, fakat ayrı referanslı unsurlar tek ayırt edilemeyen kutuda gösterilemez.
- Şekiller Word dosyası istenmişse REFERANS NUMARALARI bölümündeki tüm gerçek sistem/cihaz/ürün/yapılanma unsurları nihai şekil setinde en az bir kez gösterilmelidir. Yöntem adımları referans listesinde yer alıyor ve sistem+yöntem şekilleri hazırlanıyorsa tüm yöntem referansları da akış/yöntem şekillerinde en az bir kez bulunmalıdır. Eksik referans varsa veya güvenilir konum belirlenemiyorsa Şekiller çıktısı oluşturulmaz.
- Şekil denetimi yalnız tek tek şekil bazında değil, son aşamada set bazında `beklenen referanslar ↔ nihai şekillerde görülen referanslar` karşılaştırmasıyla yapılır.


## 2026-08-14.v9 — Üçlü Son Kalite Kapısı ve Yürütülebilir Taşıyıcı Ayrımı

- Word üretiminden sonra, indirme verilmeden önce **3 zorunlu kontrol** yeniden yapılır: (1) BBF/ek teknik kaynaklardaki tüm mandatory teknik bilgilerin nihai Word'de kullanılması, (2) ana istem ve bütün alt istemlerin biçim + teknik kapsam + semantik tekrar + gerçek daraltma + gereklilik bakımından kontrolü, (3) BULUŞUN DETAYLI AÇIKLAMASI ve İSTEMLER içinde referans-listesi unsurlarının her kullanımında doğru parantezli unsur numarası bulunması.
- Ortak `elektronik işlem birimi üzerinde koşturulan yazılım vasıtasıyla çalışan ve;` üst maddesi yalnız yürütülebilir yazılım/modül/kontrolör/arayüz/yığın gibi teknik birimleri gruplayabilir. Veritabanı, bellek, veri deposu veya salt veri yapısı kaynakta açıkça yürütülebilir yazılım/modül olarak tanımlanmıyorsa bu grubun altına alınmaz.
- Sistem/cihaz bağımlı istemlerinde `bulunmasıdır` artık açıkça yasaktır. Profil gibi veri içeriği “veritabanında profil bulunmasıdır” diye değil, ilgili teknik unsurun `... profilini barındıran bir veritabanı olmasıdır` veya uygun biçimde `içermesidir` diliyle yazılır.
- Yöntem işlem adımlarının referans listesi görünümü sistem `(N)` işaretlerini içermez; aynı adım detaylı açıklama ve bağımsız yöntem isteminde referanslı unsur adlarıyla birlikte yazılır.


## 2026-08-14.v10 — Şablon ve İşlem-Adımı Dil Kapısı

- Tarifname sonrası kontrol artık 5 kapıdır: BBF tamlığı; ana/alt istem; referans; bağlayıcı şablon; unsur ve yöntem-adımı dili.
- `İSTEMLER` ve `ÖZET` yeni sayfa ve ortalı başlık, şablondaki boş paragraf ritmi ve bağımsız istem kapanışlarının şablondaki girinti/hiza özellikleri zorunludur.
- İstemlerde teknik eleman türü yerine `unsur` placeholder'ı kullanılmaz; `eleman` veya daha spesifik anten/modül/birim/sunucu/veritabanı adı kullanılır.
- Yöntem adımları salt `takibi`, `orkestrasyonu`, `kontrolü` gibi isimlerle bitemez; `takibinin yapılması`, `orkestrasyonunun gerçekleştirilmesi` gibi gerçek işlem sonu gerekir ve 3 görünümde birebir senkron tutulur.


## 2026-08-14.v11 — Tam Word Şablon Sadakati (Bağlayıcı)

- Tarifname sonrası 4. kalite kapısı artık yalnız İSTEMLER/ÖZET hizası ve birkaç boşluğu kontrol etmez; `Tarifname_181176_template.docx` ile tam yapısal karşılaştırma yapar.
- Section sayısı ve geometrisi, marjlar, sayfa boyutu/yönü, header/footer mesafeleri, header/footer içerikleri ve PAGE alanları şablonla aynı kalmalıdır. Sayfa numarası yalnız şablondaki üst/header konumunda bulunur; footer'a PAGE eklenemez.
- `BULUŞUN KISA AÇIKLAMASI` öncesindeki görsel boşluk, şablondaki sonuç paragrafının `space-after` değeriyle; `ŞEKİLLERİN KISA AÇIKLAMASI` öncesindeki boş paragraf; son Şekil açıklamasından sonra `Çizimlerin...` paragrafı öncesindeki boşluk; `BULUŞUN DETAYLI AÇIKLAMASI` öncesindeki boşluk; `İSTEMLER` öncesindeki iki boşluk ve `ÖZET` ritmi deterministik olarak doğrulanır.
- Sistem referansları ve yöntem referansları birlikteyse REFERANS NUMARALARI içinde aralarında tek boş paragraf bulunur; son yöntem adımından sonra detaylı açıklama başlığına geçmeden tek boş paragraf bulunur.
- Sabit/dinamik Word paragrafları yaklaşık biçimlendirmeyle sıfırdan kurulmak yerine mümkün olduğunca bağlayıcı şablondaki karşılık paragraf arketipinden kopyalanır.
- Tam şablon kapısı başarısızsa render alınmış olsa bile dosya kullanıcıya sunulmaz.


## 2026-08-14.v12 — Ham Kaynak, SVG ve İstem Türü Ek Sert Kapıları

- Otomatik istem türü seçiminde açık sistem/modül dayanağı ile açık yöntem adımları birlikte varsa `Sistem ve yöntem` zorunludur; modelin `yalnız yöntem`/`yalnız sistem` önerisi kaynak dayanağını düşüremez.
- Bütün Türkçe bağımlı yöntem istemleri gerçek ek işlem adımı diliyle kurulur ve son cümle tek işlem için `işlem adımını içermesidir.`, çoklu işlem için `işlem adımlarını içermesidir.` şeklinde biter.
- BBF ve açık ek teknik belgeler deterministik ham-pasaj kayıtlarına ayrılır. Her kayıt exactly-once mantığıyla bir veya daha fazla teknik fact'e bağlanır ya da yalnız açık idari/form niteliğinde gerekçeli teknik-dışı sınıflandırılır. Böylece modelin eksik `technical_facts` listesi üretip kendi eksik listesini 100% karşılaması artık yeterli değildir.
- `.svg` dosyaları ZIP, ek teknik belge ve şekil yükleme akışlarında birinci sınıf kaynak görselidir. Raster dönüşüm yalnız görüntüleme/Word yerleşimi içindir; özgün müşteri şekli başka model çizimiyle ikame edilemez. Kaynak şekil envanterindeki her kullanılabilir müşteri şekli nihai şekil setinde yer almalıdır.
