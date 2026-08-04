# Patent Atölyesi – Kayıtlı İş Kuralları

Kural sürümü: **2026-08-04.v2**

Bu dosya arayüzde kullanılan kuralların okunabilir özetidir. Uygulamanın çalıştırdığı tam metin `rules.py` içindedir.

## 1. Tarifname için değişmez temel

BBF'de bulunan bütün teknik bilgiler kullanılmalıdır. Önceki teknik açıklamaları, teknik problem, çözüm, unsurlar, işlevler, yöntem akışı, formüller, matematiksel ilişkiler, deneysel sonuçlar, tablolar, alternatif gerçekleştirmeler, kullanım senaryoları, şekil açıklamaları, referans tablosu ve teknik etkiler atlanamaz.

Kullanıcının mevcut tarifnameye aktardığı veya düzelttiği teknik metin korunur. Metin sırf kısaltmak amacıyla daraltılmaz. Önceden hazırlanmış benzer tarifnameler yalnızca unsur ve istem kurgusunu görmek için kullanılır; bunların teknik içeriği yeni buluşa taşınmaz.

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
