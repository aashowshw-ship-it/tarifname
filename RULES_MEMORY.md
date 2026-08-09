# Patent Atölyesi – Kayıtlı İş Kuralları

Kural sürümü: **2026-08-09.v1**

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

Bağımlı istemler ana istemi tekrar etmez. Yalnızca BBF'de dayanağı bulunan ve kapsamı gerçek anlamda daraltan teknik ayrıntıları ekler.

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
- Tarifname ve bölüm başlıkları kalındır. `İSTEMLER` yeni bir sayfadan, `ÖZET` ayrıca yeni bir sayfadan başlar.
- **“Buluşun bir gerçekleştirilmesinde” kullanılmaz; “Buluşun bir yapılanmasında” kullanılır.**
- Önceki teknik bölümünde müşterinin kaynakta verdiği teknik arka plan, eksiklikler, problem anlatımı ve karşılaştırmalar eksiksiz korunur. Seçilen patent literatürü bunların yerine geçmez, yalnızca ayrı patent paragrafları olarak eklenir.
- Tarifname oluşturma arayüzünde `Mevcut/revize tarifname` alanı bulunmaz. Mevcut tarifnameyi değiştirme işi ayrı `Tarifname düzenleme` akışıdır ve şimdilik bu pakete eklenmemiştir.
- Şekiller seçimi tarifname akışında literatür araştırmasından önce gösterilir.
- `TEKNİK ALAN` ilk cümlesi **“Buluş, ... ile ilgilidir.”** biçiminde yazılır. Teknik alan ya tek bir bütün paragraf halinde tamamlanır ve bu paragraf içinde ikinci kez **“Buluş, özellikle ...”** diye başlanmaz; ya da ek açıklama **“Buluş, özellikle ...”** ile verilecekse bu ifade mutlaka ayrı bir paragrafın başlangıcında yer alır.
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

