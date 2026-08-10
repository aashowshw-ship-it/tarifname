from __future__ import annotations

APP_VERSION = "v5.4.3"
RULESET_VERSION = "2026-08-10.v3"

TARIFNAME_RULES = r"""
TÜRK PATENT TARİFNAME OLUŞTURMA KURALLARI

A. KAYNAK SADAKATİ VE İÇERİK TAMLIĞI
1. Yeni tarifname oluşturma modülünde teknik içerik bakımından yalnızca ilgili BBF ve ayrıca teknik kaynak olarak açıkça yüklenen müşteri belgeleri kullanılabilir. Önceden hazırlanmış benzer tarifnameler yalnızca unsur, yöntem adımı, istem ve biçim kurgusunun nasıl kurulacağını görmek için kullanılabilir; bunların teknik içeriği yeni dosyaya taşınamaz.
2. BBF'deki bütün teknik bilgiler eksiksiz kullanılmalıdır. Özellikle önceki teknik açıklamaları, teknik problem, çözüm, unsurlar, işlevler, yöntem akışı, formüller, matematiksel ilişkiler, deneysel sonuçlar, tablolar, alternatif gerçekleştirmeler, kullanım senaryoları, şekil açıklamaları, referans tablosu ve teknik etkiler atlanamaz veya kısa bir özetle ikame edilemez. Tamlık kontrolü yalnız çıkarılmış metne göre yapılmaz; BBF ve ek teknik belgelerdeki gömülü şekil, grafik, diyagram, ısı haritası, eksen/etiket ve görsel olarak sunulan teknik sonuçlar da incelenir ve teknik anlatım açısından gerekli bilgiler tarifnameye aktarılır.
3. Yeni tarifname oluşturma ekranında “Mevcut/revize tarifname” kaynağı kullanılmaz. Mevcut bir tarifnamenin değiştirilmesi ayrı bir “Tarifname düzenleme” iş akışıdır; bu işlev yeni tarifname oluşturma akışına karıştırılmaz.
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
17. BBF'de sistem/cihaz unsurlarına açık referans numarası verilmişse bu numaralandırmayı aynen koru. Sistem/cihaz unsurları açıkça adlandırılmış fakat bunlara ayrı referans numarası verilmemişse Tarifname Oluşturma standardı olarak kaynak sırasıyla 1, 2, 3... referansları atanır. Yöntem işlem adımları sistem unsur referanslarından ayrı bir referans ailesidir ve kullanıcı aksini açıkça istemedikçe 1001, 1002, 1003... biçiminde sıralanır. Kaynakta yöntem satırları 1, 2, 3... veya S101, S102... gibi gösterilmiş olsa dahi yeni tarifname çıktısında sistem unsurlarıyla karışmaması için yöntem adımları 1001'den başlatılır. Açıkça tanımlanmamış genel kavramlara uydurma referans numarası verilmez.
18. REFERANS NUMARALARI bölümünden önce unsur veya işlem adımı referansı kullanma. Özellikle BULUŞUN KISA AÇIKLAMASI bölümünde ana istem yapısını özetlerken (1), (2), (3) gibi referans işaretlerini kopyalama; unsur adlarını numarasız yaz.
19. Referans listesinde önce sistem/cihaz unsurları yer alır. Kaynakta UW, UW_F, UW_PL, UW_R ve UW_M gibi sembolik referans işaretleri gerçekten teknik gösterim olarak kullanılıyorsa, sayısal unsur listesinden sonra bir boş paragraf bırakılarak "UW. Kullanılabilir ağırlık", "UW_F. İlave yakıt tahsisi" mantığında yazılır. Kaynakta gerçek unsur referans numarası olmayan 21-37 gibi şekil içi/geçici numaralar yeni referans numarası olarak uydurulmaz. Sembolik referanslardan sonra bir boş paragraf bırakılarak yöntem işlem adımları yazılır.
20. Yöntem işlem adımlarının ilk kelimesi büyük harfle başlar. Metinler işlem isimleri şeklinde yazılır: "toplanması", "hesaplanması", "oluşturulması" gibi.
21. "Yöntemin gerçekleştirdiği işlem adımları aşağıdaki gibidir:" bölümündeki numaralı adımlar, REFERANS NUMARALARI bölümü ve yöntem istemlerinde aynı numara için kullanılan işlem metinleri anlam ve terminoloji bakımından uyumlu olmalıdır. Kullanıcı birebir eşleşme istiyorsa metinler birebir aynı tutulmalıdır.
22. Ana istemde kapsayıcı bir adım numarasız verilip ayrıntılı numaralı paralel adımlar bağımlı isteme taşınabilir. Bu durumda referans tablosu ve detaylı açıklamadaki numaralar değiştirilmez.
23. Referans numaraları değiştirilmeden önce tüm şekiller, detaylı açıklama ve istem bağlantıları kontrol edilmelidir. Gereksiz numara değişikliği yapılmamalıdır.
24. Detaylı açıklamada her yöntem adımı yalnızca kendi işleviyle ve doğru aşama içinde açıklanmalıdır. Aynı numara farklı işlev için kullanılamaz.

D. İSTEM KURGUSU
25. Ana sistem isteminde unsurlar BBF sırasıyla tanımlanmalı; bir unsur yalnızca daha önce tanımlanmış unsurlarla teknik ilişki kurmalıdır. Henüz tanımlanmamış sonraki unsur önceki unsurun içinde kullanılamaz.
26. Ana istemde unsurlar bağımsız bir liste olarak kalmamalı; aralarında veri, sinyal, kontrol, işlem veya fiziksel bağlantı ilişkisi kurulmalıdır. Teknik olarak ilişkili sonraki unsur, mümkün olduğunda kendisinden önce tanımlanan unsurun çıktısını/girdisini veya o unsurla bağlantısını açıkça belirtmelidir. Ancak kaynakta ilişki bulunmayan unsurlar sırf biçim için yapay biçimde birbirine bağlanmamalıdır.
27. Ana yöntem istemi zorunlu işlem sırasını ve teknik taşıyıcıyı açıkça göstermelidir. İnsan eylemleri yerine elektronik işlem birimi, cihaz, sunucu, bulut veya teknik modül gibi taşıyıcılar kullanılmalıdır.
28. Bağımlı istemler ana istemi tekrar etmemeli; yalnızca kaynakta dayanağı bulunan ve kapsamı gerçek anlamda daraltan teknik ayrıntıları eklemelidir.
29. Aynı alt teknik akışın analiz ve çıktı adımları, teknik bütünlük bozulmayacaksa tek bağımlı istemde toplanabilir.
30. Formüller kaynakta dayanaklıysa detaylı açıklamada korunmalı; zorunlu çekirdek değilse ana istemi gereksiz daraltmamak için uygun bağımlı istemlerde kullanılmalıdır.
31. Sistem alt istemlerini "bir modül olmasıdır" veya "içermesidir" biçiminde bitir. "yapmasıdır/etmesidir/belirlemesidir" kullanma.
32. Yöntem istemini "işlem adımlarını içermesidir" biçiminde bitir.
33. İstem numaraları kalın yazılmalıdır.
34. İstemler oluşturulduktan sonra ikinci bir istem kalite kontrolü yapılmalıdır: zorunlu çekirdek, kapsam, teknik taşıyıcı, unsur sırası, paralel adımların gruplanması, tekrar, dayanak, formül kullanımı ve dil.

E. DİL, PARAGRAF VE BİÇİM
35. İngilizce teknik terimi ilk geçtiği yerde Türkçesini önce, İngilizcesini parantez içinde ver. Sonraki kullanımlarda yalnızca Türkçe karşılığını kullan. AI yerine yapay zekâ yaz.
36. Unsur adlarını normal cümle düzeninde yaz; her kelimeyi başlık biçiminde büyük harfle başlatma.
37. Şablondaki kırmızı/mavi açıklama metinlerini ve biçimlerini koru.
38. Detaylı açıklamadaki modül/unsur ve çalışma prensibi anlatımını gereksiz yere çok sayıda küçük paragrafa bölme; teknik akış elverdiği ölçüde bağlantılı ve sürekli anlat.
39. İSTEMLER ve ÖZET başlıklarını ortala. Özet tek paragraf ve tercihen tek cümle olacak biçimde kısa tutulmalıdır.
40. Patent literatürü paragrafında yayın/başvuru numarası ile doğrulanmış İngilizce başlık ve bunun Türkçe karşılığı birlikte verilmeli; dokümanın teknik konusu ve buluşta bulunmayan temel fark açıkça belirtilmelidir.
41. TEKNİK ALAN bölümü iki kademeli yazılır. İlk paragraf yalnız giriş cümlesidir ve "Buluş, ... ile ilgilidir." yapısında tamamlanarak biter. Daha ayrıntılı teknik alan açıklaması aynı paragrafta devam ettirilmez. İkinci paragraf mutlaka "Buluş, özellikle ..." ifadesiyle başlar ve buluşun daha özel teknik kapsamını açıklar. "Sistem ve yöntem..." gibi çıplak bir ifadeyle ikinci paragrafa başlanmaz ve "Buluş özellikle" yerine daima "Buluş, özellikle" kullanılır.
42. ÖNCEKİ TEKNİK bölümünde aynı teknik anlatımın devamı olan "Özellikle...", "Bununla birlikte...", "Bu nedenle...", "Ayrıca...", "Böylece..." gibi devam cümleleri sırf yeni cümle başladığı için ayrı paragraf yapılmaz; önceki paragrafın devamı olarak birleştirilir. Ancak patent literatüründeki her ayrı doküman kendi paragrafında kalır.
43. BULUŞUN DETAYLI AÇIKLAMASI bölümünde referanslı sistem/cihaz unsurları tek tek ayrı paragraf yapılmaz. Unsurlar BBF sırası ve teknik bağlantıları korunarak tek ve sürekli bir unsur-açıklama paragrafında anlatılır. Sistem unsurları ile yöntem işlem adımları arasındaki ilişki de "İşlem Adımı / Gerçekleştiren Unsur / Açıklama" türü açıklama tablosuna dönüştürülmez; modülün hangi önceki modülden veri aldığı, hangi işlem adımını gerçekleştirdiği ve hangi çıktıyı sonraki unsura aktardığı doğal teknik paragraf içinde açıklanır. Yalnız kaynakta gerçekten teknik veri tablosu niteliğinde olan sayısal/deneysel tablolar tablo olarak korunabilir. "Buluşun bir yapılanmasında...", farklı alternatif/gerçekleştirme, yöntem adımları veya çalışma prensibi gibi gerçekten farklı anlatımlar ayrı paragraf olabilir.
44. "Yöntemin gerçekleştirdiği işlem adımları aşağıdaki gibidir:" ifadesinden sonraki madde işaretli işlem adımlarında ara adımlar virgül ile, son adım nokta ile bitirilir. Yöntem bağımsız istemindeki madde işaretli işlem adımlarının her biri virgül ile biter ve listenin ardından ayrı satırda "işlem adımlarını içermesidir." yazılır. İşlem adımı satırları noktalamasız bırakılmaz.
45. ŞEKİLLERİN KISA AÇIKLAMASI bölümündeki açıklamalar kısa ve işlevsel tutulur. "Şekil 1, ...", "Şekil 2, ...", "Şekil 3, ..." açıklamaları aralarında boş paragraf olmadan alt alta sıralanır. Akış diyagramı açıklamasında gerekli değilse "1001-1004 numaralı" gibi adım numarası aralıkları tekrarlanmaz; "işlem adımlarını gösteren temsili akış diyagramıdır" türü ifade yeterlidir.
46. `Tarifname_181176_template.docx` yalnız font ve başlık açısından değil; boş paragraf düzeni, paragraf aralıkları, 1,5 satır aralığı, madde işareti/otomatik numaralandırma yapısı, istemler arası boşluklar, sayfa geçişleri ve hizalamalar bakımından da bağlayıcıdır. Çıktı bu şablonun görsel ritmini birebir takip etmelidir.

F. SON KALİTE KONTROLÜ
47. Çıktıdan önce şu kontroller birlikte yapılmalıdır: BBF'deki tüm bilgilerin aktarımı, referans tablosu tamlığı, detaylı açıklama–referans–istem uyumu, yöntem adımı sırası, ana istemin buluşu gerçekten yansıtması, bağımlı istemlerin gerçek daraltma sağlaması, formüller, tablolar ve deneysel sonuçların korunması.
48. Aynı metinle yazılmış farklı yöntem adımı numaraları özellikle kontrol edilmelidir. Aynı veri farklı aşamalarda kullanılıyorsa aşama farkı metne yansıtılmalıdır.
49. Kaynaktaki önemli bir bölümün yalnızca özetlenip ayrıntılarının kaybolduğu tespit edilirse taslak tamamlanmış sayılmaz; eksik bilgiler yeniden eklenmelidir.
50. Sistem ve yöntem istemlerinin birlikte oluşturulduğu durumda buluş başlığı da bu yapıyla uyumlu olmalı ve uygun ise “... Sistemi ve Yöntemi” biçimini taşımalıdır; yalnızca sistem başlığı bırakılmamalıdır.
51. Patent tarifnamesinin kullanıcıya sunulan metninde “BBF”, “buluş bildirim formu”, “kaynak formda açıklandığı üzere” veya benzeri kaynak-doküman atıfları kullanılmaz. Teknik bilgi doğrudan buluşun açıklaması olarak yazılır.
52. REFERANS NUMARALARI bölümünde unsur adları başlık biçiminde yazılmaz. Yalnızca ilk kelimenin ilk harfi büyük olur; standart teknik kısaltmalar (SIM, IMEI, API vb.) kendi yazımıyla korunabilir. Aynı unsur adı cümle içinde geçtiğinde cümle gereği küçük harfle başlatılır.
53. Detaylı açıklamada “Yöntemin gerçekleştirdiği işlem adımları aşağıdaki gibidir:” ifadesinden sonra her işlem madde işaretli yazılır ve numara metnin başında “1001.” biçiminde değil, işlem metninin sonunda “(1001)” biçiminde gösterilir.
54. Sistem ve yöntem bağımsız istemlerinde ayrı sistem unsurları ve işlem adımları düz cümle halinde arka arkaya verilmez; her unsur/adım ayrı gerçek Word madde işaretiyle gösterilir.
55. TARİFNAME, bölüm başlıkları, buluş başlığı, İSTEMLER ve ÖZET başlıkları kalın yazılır. İSTEMLER yeni bir sayfadan, ÖZET ayrıca yeni bir sayfadan başlatılır.
56. “Buluşun bir gerçekleştirilmesinde” kalıbı kullanılmaz. Bu anlatım gereken yerde “Buluşun bir yapılanmasında” yazılır.
57. Önceki teknik bölümünde kaynakta müşterinin verdiği teknik arka plan, eksiklik, problem ve karşılaştırma bilgileri eksiksiz aktarılır; seçilen patent literatürü bu bilgilerin yerine geçmez ve yalnızca bunlara eklenir.
58. Kısa açıklamadaki şekillere geçiş cümlesi “Mevcut buluş...” ile başlamaz; “Buluşun yapılanması...” yapısında yazılır. Çizim açıklama kapanışında da “mevcut buluş” kullanılmaz; “buluş” kullanılır.
59. Otomatik kalite kontrolünde TEKNİK ALAN giriş kalıbı, önceki teknik devam paragrafları, literatürde İngilizce+Türkçe başlık, detaylı açıklamadaki unsur paragraf bütünlüğü, yöntem adımlarının noktalaması ve şablon boşluk yapısı ayrıca kontrol edilir.
60. BULUŞUN KISA AÇIKLAMASI bölümündeki amaç cümleleri çıplak mastarla bitmez. "Buluşun ana amacı, ... karşılaştırmaktır.", "Buluşun diğer bir amacı, ... sağlamaktır." gibi tam yüklemli ve dilbilgisel olarak tamamlanmış cümleler kullanılır.
61. Tarifnamenin açıklama bölümlerinde noktalı virgül gereksiz yere kullanılmaz. Normal teknik anlatımda virgül veya nokta tercih edilir. İstemlerin standart "olup, özelliği;" kalıbındaki noktalı virgül korunabilir.
62. İstemlerde HPU_W, FW_min, UW, UW_F, FE, TE gibi semboller tek başına teknik unsur adı yerine kullanılmaz. Anlamın gerekli olduğu yerde önce teknik açılım yazılır ve sembol parantez içinde verilir; örneğin "kuru hibrit güç ünitesi ağırlığı (HPU_W)", "asgari görev yakıtı ağırlığı (FW_min)", "ilave yakıt tahsisi (UW_F)". Matematiksel ilişkilerde semboller aynen korunur.
63. Yöntem işlem adımlarının teknik metni REFERANS NUMARALARI, BULUŞUN DETAYLI AÇIKLAMASI ve yöntem istemi arasında terminoloji bakımından aynı tutulur. Varsayılan yöntem referansları 1001, 1002... biçimindedir. Sadece bulunduğu bölüme göre numara konumu ve son noktalama değişebilir: referans listesinde "1001. ...", detaylı açıklamada ara adım için "... (1001)," ve son adım için "... (100N).", bağımsız yöntem isteminde "... (1001),".
64. Kaynakta bir sembolün açılımı veriliyorsa istemlerde ve kritik teknik açıklamalarda çıplak sembol kullanarak anlamı belirsiz bırakma; kullanıcı sonradan "bu değişken neydi" demeyecek şekilde açılımı metne taşı.
G. TARİFNAME OLUŞTURMA - ŞEKİLLERİN OLUŞTURULMASI
65. Şekil üretiminde ilk tercih müşterinin BBF, teknik ek veya ayrıca yüklediği şekil dosyasında verdiği özgün teknik şekildir. Müşterinin teknik kurgusu, kutu-ok ilişkileri, geometri ve teknik anlamı sırf daha estetik görünmesi için yeniden tasarlanmaz.
66. Şekiller ayrı Word dosyasında sırasıyla “ŞEKİL 1”, “ŞEKİL 2”, “ŞEKİL 3” şeklinde devam eder. Şekil numarası başlığı ilgili görselin altında, ortalı ve kalın yer alır.
67. Şekiller Word dosyasının her sayfasının üst kısmında “mevcut sayfa / toplam sayfa” biçiminde sayfa göstergesi bulunur; örneğin “1 / 3”. Toplam sayfa sayısı sabit değildir. Şekil sayısı, şekillerin boyutu ve okunabilirliği dikkate alınarak dinamik belirlenir. Bir sayfada bir veya birden fazla şekil bulunabilir.
68. Şekiller dosyasında gereksiz açıklama paragrafı kullanılmaz. Temel çıktı, müşteri şekli ile altında “ŞEKİL N” başlığından oluşur. Görsel okunabilirliği korunmalı, görsel gereksiz sıkıştırılmamalı veya düşük çözünürlüğe dönüştürülmemelidir.
69. Nihai şekiller hazırlanırken önce müşterinin şekilleri alınır, ardından şekillerde görünen teknik referans işaretleri REFERANS NUMARALARI bölümü ile karşılaştırılır. Nihai şekil içinde kullanılan hiçbir gerçek referans işareti karşılıksız bırakılamaz.
70. Şekilde “1”, “2”, “3” gibi unsur referansları veya “S101”, “1001” gibi yöntem adımı referansları kullanılıyorsa bunların tamamı tarifnamedeki REFERANS NUMARALARI bölümünde tanımlı olmalı ve detaylı açıklama/istemlerle aynı teknik karşılığı taşımalıdır.
71. Kaynakta UW, UW_F, UW_PL, UW_R, UW_M gibi semboller şekil üzerinde unsur/karar değişkeni referansı olarak kullanılıyorsa bunlar sayısal numaraya dönüştürülmez. REFERANS NUMARALARI bölümünde “UW. Kullanılabilir ağırlık”, “UW_F. İlave yakıt tahsisi” biçiminde tanımlanır. Detaylı açıklama ve istemlerde ilk uygun kullanım “ilave yakıt tahsisi (UW_F)” biçiminde teknik ad + parantez içinde sembol olarak yazılır.
72. Müşteri şekillerindeki 21-37 gibi yalnız şeklin hazırlanması sırasında verilmiş, tarifnamenin gerçek unsur/referans sistemiyle uyumlu olmayan geçici sayılar yeni referans numarası olarak tarifnameye taşınmaz. Teknik anlam kaybı oluşturmadan güvenle kaldırılabiliyorsa şekilden kaldırılır; kaldırılamıyorsa nihai çıktıdan önce kullanıcıya uyumsuzluk bildirilir.
73. Bir teknik unsur REFERANS NUMARALARI bölümünde şekille ilişkili olarak tanımlanmış fakat ilgili müşteri şekli üzerinde referans işareti eksik bırakılmışsa şekil nihai kabul edilmez. Teknik konumu kaynaktan açıkça belirlenebiliyorsa yalnız eksik referans işareti eklenir; konum belirsizse unsurun yeri uydurulmaz ve kullanıcıdan düzeltilmiş şekil istenir.
74. Patent şekillerinde açıklama yazısı mümkün olduğunca azaltılır. Bununla birlikte müşteri tarafından sağlanan özgün teknik şeklin içindeki yazı, formül veya ilişki kaldırıldığında teknik anlam, hesaplama mantığı veya müşteri açıklamasının kapsamı kaybolacaksa şekil müşterinin sunduğu biçime yakın korunabilir. Bu, teknik içeriği korumak amacıyla bilinçli olarak alınan bir şekil-formalite riskidir; teknik içerik yalnız formalite amacıyla değiştirilmez.
75. Şekil kalite kontrolü, nihai tarifname ile birlikte şu eşleşmeleri zorunlu olarak denetler: şekil ↔ REFERANS NUMARALARI, şekil ↔ detaylı açıklama, şekil ↔ istemler ve yöntem şekilleri için şekil ↔ yöntem adımları. Yeni tarifnamede yöntem adımları 1001... standardına dönüştürülmüşse, müşteri akış diyagramındaki yöntem adımı işaretleri de teknik anlam ve yerleşim güvenle korunabiliyorsa aynı 1001... referanslarıyla senkronize edilir; güvenli biçimde değiştirilemiyorsa kullanıcıya uyumsuzluk bildirilir. Eksik, fazla, çelişkili veya farklı anlamda kullanılan referans varsa şekiller tamamlanmış sayılmaz.
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
21. Görüş akışı tek adımda doğrudan Word üretmez. İlk düğme yalnızca raporu, önceki görüşü, tarifnameyi, X/Y dokümanlarını ve varsa müşteri bilgisini analiz eder; analiz sonucu ekranda gösterildikten sonra görüş oluşturma aşamasına geçilir.
22. İlk analiz istem revizyonu gerekip gerekmediğini açıkça belirler. Revizyon gerekmiyorsa kullanıcıya bu durum gösterilir ve mevcut istemlerle görüş oluşturma düğmesi açılır.
23. Revizyon gerekiyorsa önce önerilen değişiklikler istem bazında gösterilir. Her değişiklik için istem numarası, değişiklik gerekçesi, tarifname dayanağı, eski ifade ve önerilen yeni ifade ayrı ayrı verilmelidir.
24. Revizyon önerileri kullanıcı onayı olmadan uygulanmaz. Kullanıcı isterse önerileri ek talimatla yeniden analiz ettirebilir. Kullanıcı önerilen revizyonları açıkça onayladıktan sonra revize istem çıktıları hazırlanır.
25. Onaylı istem revizyonu için iki Word çıktısı oluşturulur: gerçek OOXML Track Changes işaretlerini içeren MARKUP sürümü ve aynı değişikliklerin kabul edilmiş halini içeren TEMİZ sürüm. Track Changes değişiklikleri mümkün olan en küçük ifade/kelime düzeyinde uygulanır; bütün istem paragrafı sırf kolaylık için silinip yeniden eklenmez.
26. Markup üretimi için kaynak tarifname DOCX olmalıdır. PDF veya eski DOC kaynakta istem revizyonu gerekiyorsa kullanıcıdan DOCX tarifname istenir; dosya dönüştürülmüş gibi varsayım yapılmaz.
27. Revize istemler kullanıcı tarafından son kez onaylanmadan görüş Word dosyası oluşturulmaz. Görüş metni yalnızca kullanıcının onayladığı mevcut veya revize nihai istem seti üzerinden hazırlanır.
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
11. Rapor için tek bağlayıcı biçim kaynağı `On_Arastirma_Raporu_181612_template.docx` dosyasıdır. Şablonun gövdesi silinip yeniden kurulmaz. Logo, header/footer, marj, sayfa geometrisi, paragraf boşlukları, font, punto, başlık konumları, tablo ölçüleri, değerlendirme bölümünün sırası, Ekler ve Önemli Not alanı şablondan korunur.
12. `2. DEĞERLENDİRME`, `2.1. Yenilik Değerlendirmesi`, D1 bölümü, D2 bölümü, `2.2. Buluş Basamağı Değerlendirmesi` ve `3. SONUÇ` sırası ve mantığı şablondaki gibi kalır; şablonda olmayan yeni alt başlıklar eklenmez.
13. D1 ve D2 karşılaştırma tablolarında sol sütundaki teknik özellik listesi birebir aynı ve aynı sırada olmalıdır. Sağ hücre yalnız çıplak `+` veya `-` değildir: hücre `+` veya `-` ile başlar ve mümkün olan her durumda özelliğin dokümanda nerede bulunduğunu `Özet`, `İstem X`, `Şekil X`, paragraf/sütun/sayfa veya ilgili tarifname bölümü gibi somut dayanakla belirtir. `±` kullanılmaz.
14. D1/D2 tanıtımı 2–3 cümle olmalı; tablo sonrası yenilik değerlendirmesi yaklaşık 5–10 satır olmalıdır.
15. 'İncelenen diğer yakın dokümanlar...' gibi şablonda bulunmayan ek cümleler veya bölüm başlıkları eklenmemelidir. Yardımcı/yeni bulunan doküman gerekiyorsa buluş basamağı değerlendirmesinin doğal paragraf akışı içinde açıklanır veya D1/D2'den daha güçlü ise D1/D2 olarak seçilir.
16. Sistem unsurları ve yöntem adımları kaynak teknik bilgilerle birebir uyumlu olmalı; yöntem adımı numaraları ve metinleri rapor boyunca senkron tutulmalıdır.
17. Patent şekilleri model tarafından çizilmez, yeniden üretilmez veya temsili olarak oluşturulmaz. D1/D2 bölümünde yalnızca ilgili patent dokümanının özgün/orijinal şekli kullanılır. Şekil mümkün olan en yüksek çözünürlükte resmi patent kaynağından, Google Patents patentimages kaynağından veya kullanıcı tarafından yüklenen orijinal patent PDF/DOCX dosyasından alınır. Orijinal şekil temin edilemiyorsa yapay şekil oluşturmak yerine kullanıcıya uyarı verilir.
18. Araştırma kesim tarihi, DP referans numarası ve çıktı dosya adı arayüzde tutulmalıdır.
19. Sonuç açık olmalıdır: yenilik sağlanır/sağlanmaz; buluş basamağı sağlanır/sağlanmaz.
20. Rapor metninde 'BBF', 'buluş bildirim formu', 'ilk BBF', 'ikinci BBF' gibi kaynak-form ifadeleri kullanılmaz. Normal Tip 3 raporunda 'araştırma konusu'; güncelleme raporunda gerekli olduğunda 'revize araştırma konusu' veya 'ilk araştırma konusu' ifadeleri kullanılır.
21. Rapor gövdesinde yapay zekâ çıktısı hissi veren `→`, `=>`, ok zincirleri, denklem gibi kurulmuş `özellik + özellik + özellik` kısa gösterimleri veya benzeri sembolik özetler kullanılmaz. Teknik ilişkiler tam ve doğal Türkçe cümlelerle açıklanır. `+` ve `-` yalnız karşılaştırma tablosunun sağ hücrelerinde şablon mantığında kullanılabilir.
22. Teknik değerlendirmeyi kaynak teknik bilgilere sadık yap; tanıtım veya salt iş kuralı niteliğindeki yönleri teknik katkı gibi abartma.
23. Rapor metnini oluşturduktan sonra ikinci kalite kontrolü yap: D1/D2 seçimi, özellik eşleştirmesi, yenilik mantığı, birleştirme motivasyonu, kullanıcı sonuç modu, özgün patent şekilleri ve sonuç tutarlılığı bakımından düzelt.
"""

ARASTIRMA_GUNCELLEME_RULES = ARASTIRMA_RULES + r"""

ARAŞTIRMA GÜNCELLEME – TİP 3 EK KURALLARI
1. Arayüzde birbirinden açıkça ayrılmış üç kaynak alanı bulunur: `İlk BBF`, `Revize BBF` ve `İlk Ön Araştırma Raporu`. Kullanıcının dosyaları karıştırmaması için etiketler aynen açık yazılır.
2. İlk aşamada yeni rapor üretilmez. Önce ilk araştırma konusu ile revize araştırma konusu teknik olarak karşılaştırılır. Yalnız metin farkı değil, yeni teknik sınırlama, yeni veri işleme ilişkisi, yeni unsur/işlev, yeni teknik etki ve teknik problemin değişip değişmediği belirlenir.
3. İlk ön araştırma raporundaki D1/D2, yenilik ve buluş basamağı gerekçeleri çıkarılır. Revize edilen özelliklerin ilk D1/D2 karşısında gerçekten yeni bir teknik katkı sağlayıp sağlamadığı ayrı ayrı değerlendirilir.
4. Kullanıcıya analiz ekranında en az şu bilgiler gösterilir: esas teknik farklar, her farkın teknik katkı oluşturup oluşturmadığı, ilk D1/D2 karşısındaki ön etkisi ve sistemin açık teknik kanaati.
5. Teknik açıdan anlamlı fark varsa veya kullanıcı yeni araştırma yapılmasını isterse, araştırma yalnızca eski sorgunun tekrarı olarak yapılmaz; revize edilen ayırt edici teknik özellikler ile ilk rapordaki D1/D2 başlangıç noktası alınarak global patent araştırması yapılır.
6. Yeni araştırmada bulunan daha güçlü dokümanlar açıkça gösterilir. Yeni doküman D1/D2'den daha yakınsa nihai D1/D2 değiştirilebilir; değilse yardımcı doküman olarak buluş basamağı değerlendirmesinde kullanılır.
7. Yeni araştırma tamamlandığında Word raporu otomatik oluşturulmaz. Arayüz önce sistem kanaatini gösterir; örneğin `Kanaatim: yenilik sağlanıyor ancak buluş basamağı halen sağlanmıyor.` Ardından kullanıcıdan rapor sonucunun `Buluş basamağı sağlanıyor` veya `Buluş basamağı sağlanmıyor` yönünde hazırlanması istenir. Kullanıcı seçimi yapılmadan Word üretimi başlamaz.
8. Yenilik sonucu kaynaklara göre otomatik ve dürüst belirlenir. Tek doküman revize araştırma konusunun bütün esas teknik özelliklerini doğrudan ve açık biçimde açıklıyorsa kullanıcı seçimi buluş basamağı yönünde olsa dahi yenilik varmış gibi yazılmaz.
9. Güncelleme raporu da yeni bir rapor formatı değildir. Çıktı, bağlayıcı `On_Arastirma_Raporu_181612_template.docx` dosyasının tam Tip 3 Ön Araştırma Raporu formatında hazırlanır. `Revizyon karşılaştırması`, `BBF farkları`, `Güncelleme özeti` gibi şablonda bulunmayan bölümler Word raporuna eklenmez; bu bilgiler yalnız arayüz analizinde gösterilir.
10. Word raporunda ilk rapora atıf gerektiğinde `ilk ön araştırma raporu`, teknik konuya atıf gerektiğinde `ilk araştırma konusu` ve `revize araştırma konusu` denebilir; `BBF` ifadesi yazılmaz.
11. D1/D2'nin özgün patent şekilleri rapora doğrudan patent kaynağından eklenir. İlk rapordaki D1/D2 korunuyorsa onların orijinal patent şekilleri yeniden kullanılır; yeni bir doküman D1/D2 olarak seçilirse onun özgün şekli alınır. Yardımcı dokümanın ayrı şekil başlığı şablonda yoksa Word'e yeni şekil bölümü açılmaz.
12. DP referans numarası, araştırma kesim tarihi ve çıktı dosyasının adı arayüzde kullanıcı tarafından belirlenir. Varsayılan çıktı adı `Ön Araştırma Raporu_<DP REF>_rev.docx` olabilir.
"""
