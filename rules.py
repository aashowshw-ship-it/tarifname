from __future__ import annotations

APP_VERSION = "v5.1"
RULESET_VERSION = "2026-08-07.v3"

TARIFNAME_RULES = r"""
TÜRK PATENT TARİFNAME OLUŞTURMA VE REVİZYON KURALLARI

A. KAYNAK SADAKATİ VE İÇERİK TAMLIĞI
1. Teknik içerik bakımından yalnızca ilgili BBF, kullanıcının revize edilmesini istediği mevcut tarifname ve ayrıca teknik kaynak olarak açıkça yüklenen müşteri belgeleri kullanılabilir. Önceden hazırlanmış benzer tarifnameler yalnızca unsur, yöntem adımı, istem ve biçim kurgusunun nasıl kurulacağını görmek için kullanılabilir; bunların teknik içeriği yeni dosyaya taşınamaz.
2. BBF'deki bütün teknik bilgiler eksiksiz kullanılmalıdır. Özellikle önceki teknik açıklamaları, teknik problem, çözüm, unsurlar, işlevler, yöntem akışı, formüller, matematiksel ilişkiler, deneysel sonuçlar, tablolar, alternatif gerçekleştirmeler, kullanım senaryoları, şekil açıklamaları, referans tablosu ve teknik etkiler atlanamaz veya kısa bir özetle ikame edilemez.
3. Kullanıcının mevcut tarifnameye aktardığı veya düzelttiği teknik bilgiler korunmalıdır. Kullanıcının metnini gereksiz yere kısaltma, teknik içeriği silme veya başka bir ifadeyle anlam daraltma yapma.
4. Kaynakta bulunmayan unsur, bağlantı, değer, algoritma, formül, teknik etki, kullanım biçimi veya avantaj ekleme. Bir husus açık değilse belirsizlik olarak belirt; uydurma yapma.
5. Önceki teknik bölümü BBF'de uzun ve ayrıntılı verilmişse aynı kapsam korunmalıdır. Patent literatürü paragrafları BBF'deki önceki teknik anlatımının yerine geçmez; yalnızca seçilmiş patent dokümanları için ayrı paragraflar olarak eklenir.
6. Formüller ve tablolar kaynakta bulunduğu biçim, değişken anlamları ve bağlamlarıyla korunmalıdır. Formülün zorunlu çekirdek mi yoksa tercihli gerçekleştirme mi olduğu ayrıca değerlendirilmelidir.

B. BULUŞUN YAPISINI BELİRLEME
7. Her buluş aynı istem mimarisiyle ele alınamaz. Önce şu ayrım yapılmalıdır: zorunlu teknik çekirdek, zorunlu işlem sırası, paralel/tekrarlanan işlem kolları, yalnızca belirli gerçekleştirmelere ait ayrıntılar, alternatifler ve sonuç/çıktılar.
8. Buluş yalnızca yöntem olarak daha doğru korunuyorsa sistem istemi oluşturma. Buluş yalnızca sistem olarak daha doğru korunuyorsa yöntem istemi oluşturma. Her iki yapının da açık teknik dayanağı varsa sistem ve yöntem istemleri birlikte hazırlanabilir.
9. Başlık, teknik alan, kısa açıklama, referans numaraları, detaylı açıklama, istemler ve özet seçilen istem yapısıyla tutarlı olmalıdır.
10. Ana istem, buluşun vazgeçilmez teknik çekirdeğini açık, sıralı, teknik taşıyıcılara bağlı ve gereksiz tekrarsız biçimde kapsamalıdır. Kısa kaldığında teknik ilişki ve işlem sırası açıklığa kavuşturularak genişletilmelidir; ancak kaynakta olmayan detay eklenmemelidir.
11. Aynı teknik işlemin birinci, ikinci ve k'ıncı metrik/kanal/modül için tekrarlanması gibi paralel adımlar ana istemde kapsayıcı biçimde yazılabilir. Örneğin "birbirinden farklı k adet görüntü kalite metriğine göre k adet ara gerçekçilik skorunun üretilmesi" şeklindeki zorunlu çoklu yapı ana istemde korunabilir.
12. Paralel analizlerin ayrı ayrı gerçekleştirilmesi ile bunlara karşılık gelen ayrı çıktıların elde edilmesi aynı alt teknik akışa aitse, bu ayrıntılar tek bir bağımlı istemde birlikte verilebilir. Gereksiz yere her bir paralel kol için ayrı bağımlı istem oluşturma.
13. Bir analiz işlemi ile bu işlemin çıktısı aynı şey değildir. Örneğin analiz adımları ve bunların skor çıktıları ayrı teknik kavramlardır; ancak aynı alt akışa aitlerse aynı bağımlı istem içinde beraber sınırlandırılabilir.
14. Eğitim/genel vektör oluşturma aşamasındaki paralel akış ile test aşamasındaki paralel akış aynı istem mantığıyla ele alınmalı, fakat aynı unsur sayılmamalıdır. Aynı tür hesaplama farklı veri ve farklı aşama üzerinde yürütülüyorsa numaraları ve işlevleri ayrı kalmalıdır.
15. BBF'de aynı girdinin farklı aşamalarda yeniden gösterildiği 1001 ve 1006 gibi unsurlar aynı metinle yazılmamalıdır. Metin, birinin tekil ara analiz girişi, diğerinin aynı verinin çoklu analiz aşamasına aktarılması gibi aşama farkını açıkça yansıtmalıdır.

C. REFERANS TABLOSU, YÖNTEM ADIMLARI VE TUTARLILIK
16. BBF'deki referans tablosu eksiksiz çıkarılmalı ve her referans unsurunun detaylı açıklamada teknik karşılığı bulunmalıdır.
17. BBF'deki unsur numaralandırmasını aynen koru. 1,2,3 ise aynı; 10,20,30 ise aynı. Kaynakta olmayan sisteme, yönteme veya genel kavrama yeni referans numarası verme.
18. REFERANS NUMARALARI bölümünden önce unsur veya işlem adımı referansı kullanma.
19. Referans listesinde önce sistem/cihaz unsurları, bir boş paragraf sonra yöntem işlem adımları yer alır.
20. Yöntem işlem adımlarının ilk kelimesi büyük harfle başlar. Metinler işlem isimleri şeklinde yazılır: "toplanması", "hesaplanması", "oluşturulması" gibi.
21. "Yöntemin gerçekleştirdiği işlem adımları aşağıdaki gibidir:" bölümündeki numaralı adımlar, REFERANS NUMARALARI bölümü ve yöntem istemlerinde aynı numara için kullanılan işlem metinleri anlam ve terminoloji bakımından uyumlu olmalıdır. Kullanıcı birebir eşleşme istiyorsa metinler birebir aynı tutulmalıdır.
22. Ana istemde kapsayıcı bir adım numarasız verilip ayrıntılı numaralı paralel adımlar bağımlı isteme taşınabilir. Bu durumda referans tablosu ve detaylı açıklamadaki numaralar değiştirilmez.
23. Referans numaraları değiştirilmeden önce tüm şekiller, detaylı açıklama ve istem bağlantıları kontrol edilmelidir. Gereksiz numara değişikliği yapılmamalıdır.
24. Detaylı açıklamada her yöntem adımı yalnızca kendi işleviyle ve doğru aşama içinde açıklanmalıdır. Aynı numara farklı işlev için kullanılamaz.

D. İSTEM KURGUSU
25. Ana sistem isteminde unsurlar BBF sırasıyla tanımlanmalı; bir unsur yalnızca daha önce tanımlanmış unsurlarla teknik ilişki kurmalıdır. Henüz tanımlanmamış sonraki unsur önceki unsurun içinde kullanılamaz.
26. Ana istemde unsurlar bağımsız bir liste olarak kalmamalı; aralarında veri, sinyal, kontrol, işlem veya fiziksel bağlantı ilişkisi kurulmalıdır.
27. Ana yöntem istemi zorunlu işlem sırasını ve teknik taşıyıcıyı açıkça göstermelidir. İnsan eylemleri yerine elektronik işlem birimi, cihaz, sunucu, bulut veya teknik modül gibi taşıyıcılar kullanılmalıdır.
28. Bağımlı istemler ana istemi tekrar etmemeli; yalnızca kaynakta dayanağı bulunan ve kapsamı gerçek anlamda daraltan teknik ayrıntıları eklemelidir.
29. Aynı alt teknik akışın analiz ve çıktı adımları, teknik bütünlük bozulmayacaksa tek bağımlı istemde toplanabilir.
30. Formüller kaynakta dayanaklıysa detaylı açıklamada korunmalı; zorunlu çekirdek değilse ana istemi gereksiz daraltmamak için uygun bağımlı istemlerde kullanılmalıdır.
31. Sistem alt istemlerini "bir modül olmasıdır" veya "içermesidir" biçiminde bitir. "yapmasıdır/etmesidir/belirlemesidir" kullanma.
32. Yöntem istemini "işlem adımlarını içermesidir" biçiminde bitir.
33. İstem numaraları kalın yazılmalıdır.
34. İstemler oluşturulduktan sonra ikinci bir istem kalite kontrolü yapılmalıdır: zorunlu çekirdek, kapsam, teknik taşıyıcı, unsur sırası, paralel adımların gruplanması, tekrar, dayanak, formül kullanımı ve dil.

E. DİL VE BİÇİM
35. İngilizce teknik terimi ilk geçtiği yerde Türkçesini önce, İngilizcesini parantez içinde ver. Sonraki kullanımlarda yalnızca Türkçe karşılığını kullan. AI yerine yapay zekâ yaz.
36. Unsur adlarını normal cümle düzeninde yaz; her kelimeyi başlık biçiminde büyük harfle başlatma.
37. Şablondaki kırmızı/mavi açıklama metinlerini ve biçimlerini koru.
38. Detaylı açıklamadaki modül ve çalışma prensibi anlatımını gereksiz yere çok sayıda küçük paragrafa bölme; teknik akış elverdiği ölçüde bağlantılı ve sürekli anlat.
39. İSTEMLER ve ÖZET başlıklarını ortala. Özet tek paragraf ve tercihen tek cümle olacak biçimde kısa tutulmalıdır.
40. Patent literatürü paragrafında yayın numarası, İngilizce başlık ve Türkçe karşılık birlikte verilmeli; dokümanın teknik konusu ve buluşta bulunmayan temel fark açıkça belirtilmelidir.

F. SON KALİTE KONTROLÜ
41. Çıktıdan önce şu kontroller birlikte yapılmalıdır: BBF'deki tüm bilgilerin aktarımı, mevcut revizyonların korunması, referans tablosu tamlığı, detaylı açıklama–referans–istem uyumu, yöntem adımı sırası, ana istemin buluşu gerçekten yansıtması, bağımlı istemlerin gerçek daraltma sağlaması, formüller, tablolar ve deneysel sonuçların korunması.
42. Aynı metinle yazılmış farklı yöntem adımı numaraları özellikle kontrol edilmelidir. Aynı veri farklı aşamalarda kullanılıyorsa aşama farkı metne yansıtılmalıdır.
43. Kaynaktaki önemli bir bölümün yalnızca özetlenip ayrıntılarının kaybolduğu tespit edilirse taslak tamamlanmış sayılmaz; eksik bilgiler yeniden eklenmelidir.
"""

GORUS_RULES = r"""
TÜRK PATENT GÖRÜŞ ÇALIŞMASI KURALLARI
1. Yalnızca raporda X veya Y olarak gösterilen dokümanlara karşı savunma yap. A kategorisi veya itiraz dayanağı yapılmayan dokümanlara karşı görüş yazma.
2. Araştırma raporuna karşı görüşte rapor, tarifname, D1/D2 ve varsa müşteri bilgilerini birlikte analiz et.
3. İnceleme raporuna karşı görüşte bunlara ek olarak önceki görüşü analiz et; uzmanın ikna olmadığı savunmaları aynen tekrarlamak yerine farklı teknik ayrım ve dayanaklar geliştir.
4. Müşteri bilgisini yalnızca tarifname/istemlerde açık dayanağı varsa doğrudan kullan. Dayanağı yoksa teknik gerçek gibi yazma; uygunsa çıkarım olduğunu belirterek yumuşat veya kullanma.
5. Teknik farklara, teknik etkiye ve unsurlar arasındaki işlevsel ilişkiye odaklan.
6. Tarifname dayanağı verilecek yerde şu kalıbı kullan: 'Tarifnamede bu durum şu şekilde belirtilmektedir:' Ardından tarifnamedeki ilgili cümle/pasajı tırnak içinde ve kalın ver.
7. Tarifname alıntısını kesme, değiştirme, sadeleştirme veya kelime ekleyip çıkarma. Alıntı tarifname metninde birebir bulunmalıdır.
8. Yenilik itirazında ilgili istemin tüm özelliklerinin tek dokümanda doğrudan ve açık biçimde açıklanmadığını göster.
9. Buluş basamağı itirazında D1 ve D2'yi tek başına ve birlikte değerlendir; teknik fark, teknik etki, objektif teknik problem, birleştirme motivasyonu ve geriye dönük değerlendirme riskini ele al.
10. Başvuru numarası ve başvuru sahibi rapordan çekilsin. Referans kullanıcıdan alınsın.
11. Görüş formatı bağlayıcı `Gorus_metni_696809_template.docx` şablonuna birebir sadık kalsın.
12. Çıktı oluşturulduktan sonra ikinci bir kalite kontrolü yap: yanlış doküman, dayanağı olmayan müşteri bilgisi, eksik alıntı, tekrar eden savunma ve sonuç tutarlılığı bakımından düzelt.
13. Görüş için bağlayıcı ve tek Word şablonu `Gorus_metni_696809_template.docx` dosyasıdır. Başka eski görüş dosyalarını şablon yerine kullanma. Şablonun logo, header/footer, marj, sayfa geometrisi, başlık konumu, font, punto ve paragraf düzenini değiştirme.
14. Başvuru bilgi alanlarında `Başvuru No`, `Başvuru Sahibi` ve `Referans` etiketleri kalın; bunların karşısındaki değerler normal yazı olmalıdır. Başvuru numarası veya diğer değerleri kalınlaştırma. Görüş gövdesindeki normal paragraflar iki yana yaslı olmalıdır.
15. İnceleme/araştırma raporunda istem değişikliği zorunlu görülmüyorsa istemleri sırf iyileştirmek amacıyla değiştirme. Bu durumda mevcut tarifname ve istemleri esas alarak görüş oluştur.
16. İstem revizyonu gerçekten gerekiyorsa önce yalnızca önerilen istem değişikliklerini kullanıcıyla netleştir. Kullanıcı açıkça onaylamadan görüş metni oluşturma ve onaylanmamış bir istem kurgusunu görüşe taşıma.
17. İstem revizyonlarında en az değişiklik ilkesi uygulanır. Yalnızca uzmanın itiraz ettiği veya düzeltme istediği noktaya ve bu noktayı gidermek için zorunlu olan ifadeye müdahale edilir; genel redaksiyon veya kapsamı gereksiz değiştiren yeniden yazım yapılmaz.
18. Revizyonlarda kapsam aşımı/yeni konu yaratma. Eklenen her teknik ifade mevcut tarifname veya istemlerde açık ve doğrudan dayanak bulmalıdır. Ürün/sistem istemlerinde yöntem dili yerine yapısal unsur dili kullan; ürün ve yöntem istemlerini birbirinden ayır.
19. İstem revizyonu onaylandıktan sonra görüş yalnızca kullanıcının onayladığı nihai istem seti üzerinden hazırlanır. Görüş oluşturma aşamasında istemlere kendiliğinden yeni değişiklik ekleme.
20. İndirilen çıktı dosyasının adı URL-kodlu görünmemelidir. `%20`, `%C3` gibi kodlanmış parçalar dosya adına taşınmamalı; Türkçe karakterler ve normal boşluklar korunmalıdır.
"""

ARASTIRMA_RULES = r"""
TİP 3 ÖN ARAŞTIRMA RAPORU KURALLARI
1. Amaç yalnızca yenilik ve buluş basamağı ön değerlendirmesidir.
2. İlk kaynak BBF'dir. Araştırmadan önce teknik unsurları, işlevleri, yöntem adımlarını, teknik problemi ve teknik etkiyi eksiksiz çıkar.
3. Global araştırma TR, EP, US, CN, KR, JP, GB, DE ve ilgili diğer patent veri tabanlarını kapsamalıdır.
4. En benzer tam 10 dokümanı doğrulanmış yayın/başvuru numarası, başlık, tarih, ülke/otorite ve kaynak bağlantısıyla belirle. Doküman uydurma.
5. İlk sonuçta kullanıcıya ayrıca tek satır halinde şu biçimi üret: 'TotalPatent arama sorgusu: CN... or US... or ...'. X/Y/A etiketi ekleme.
6. Kullanıcının yüklediği benzer dokümanları da incele; ilk seçilen D1/D2'nin yerini alabilecek daha yakın veya daha güçlü doküman varsa nihai seçimi değiştir.
7. Tek bir doküman araştırma konusu buluşun bütün esas teknik özelliklerini ve aralarındaki ilişkiyi doğrudan ve açık biçimde açıklıyorsa bu dokümanı D1 seç ve yenilik kriterinin sağlanmadığı sonucuna göre rapor hazırla. Bu durumda D2 zorunlu değildir.
8. Yeniliği tek başına bozan doküman yoksa en yakın D1 ve tamamlayıcı D2'yi seç; yenilik değerlendirmesini ayrı ayrı, buluş basamağını D1 ve D2 birlikte düşünülerek yap.
9. Yardımcı dokümanlar buluş basamağı değerlendirmesinde yalnızca destekleyici olabilir; nihai D1/D2 açıkça belirtilmelidir.
10. Kullanıcının seçtiği sonuç modu 'Buluş basamağı var', 'Buluş basamağı yok' veya 'Otomatik belirle' olabilir. Otomatik mod dışında teknik değerlendirme seçilen sonuca uygun biçimde yapılandırılmalı, ancak kaynaklarla açıkça çelişen iddia uydurulmamalıdır.
11. Rapor, Ön Araştırma Raporu_181612 şablonuna sadık kalmalıdır. Arial yazı tipi, boyutlar, boşluklar, logo, başlıklar ve tablo düzeni değiştirilmemelidir.
12. '2. DEĞERLENDİRME' bölümünde D1 ve D2 doküman numaraları kalın yazılmalıdır.
13. D1 ve D2 karşılaştırma tablolarında sol özellik sütunu aynı teknik özellikleri içermeli; yalnızca + veya - kullanılmalı, ± kullanılmamalıdır.
14. D1/D2 tanıtımı 2–3 cümle olmalı; tablo sonrası yenilik değerlendirmesi yaklaşık 5–10 satır olmalıdır.
15. 'İncelenen diğer yakın dokümanlar...' gibi şablonda bulunmayan ek cümleler eklenmemelidir.
16. Sistem unsurları ve yöntem adımları BBF ile birebir uyumlu olmalı; yöntem adımı numaraları ve metinleri rapor boyunca senkron tutulmalıdır.
17. Patent şekilleri mümkün olduğunca yüksek çözünürlükte kullanılmalıdır.
18. Araştırma kesim tarihi, DP referans numarası ve çıktı dosya adı arayüzde tutulmalıdır.
19. Sonuç açık olmalıdır: yenilik sağlanır/sağlanmaz; buluş basamağı sağlanır/sağlanmaz.
20. Teknik değerlendirmeyi BBF'ye sadık yap; tanıtım veya salt iş kuralı niteliğindeki yönleri teknik katkı gibi abartma.
21. Rapor metnini oluşturduktan sonra ikinci kalite kontrolü yap: D1/D2 seçimi, özellik eşleştirmesi, yenilik mantığı, birleştirme motivasyonu, kullanıcı sonuç modu ve sonuç tutarlılığı bakımından düzelt.
"""
