

> **UYARI / GEÇERSİZ ÜSTÜNLÜK (18.11.2026):** Bu dosya geçmiş **700801** biçim örneğini belgeler. Kullanıcının en son yüklediği **700851** Word'ü artık bağlayıcıdır. Bu dosyadaki boş paragraf veya alt çizgi dağılımının 700851 ile çelişen hükümleri uygulanmaz. Güncel tek biçim otoritesi `700851_SOURCE_OF_TRUTH.md` ve Kitaplıktaki orijinal Word dosyasıdır. Yeni UI `INPUT_SCHEMA.json`; yeni analiz `WORKFLOW.md`.

# BAĞLAYICI YİDD / YİDK RED KARARINA İTİRAZ WORD ŞABLONU — 700801
Sürüm: 2026-11-18 / kaynak: kullanıcı yüklemesi `Kurum kararına itiraz_700801.docx`.
Kaynak SHA-256: `41ca699cf55e1d79bbeaad81d52f5d9dc29fda14f9d259467711ab37dfad5198`.
Asıl Word dosyası ChatGPT özel dosya Kitaplığı: `/Patent Atolyesi/Sablonlar/Red Kararina Itiraz/Kurum kararına itiraz_700801_BAGLAYICI.docx`.
Bu kural dosyası gerçek Word'ün yerine GEÇMEZ: biçim üretimi doğrudan Word örneğini klonlayıp OOXML paragraf özelliklerini devralır.

## 1. Sabit antet ve genel stil (tek tek zorunlu)
01. Orijinal Destek Patent anteti, üst logo, alt çizgi/logolar, adres ve telefon bilgilerinin görsel varlığı ve konumlandırması korunur. Kaynaksız yeniden çizim veya yeniden tasarım FAIL.
02. Kaynak dosyada ana gövdeden önce **ardışık iki gerçek boş paragraf** vardır. Bu iki `w:p` silinmez veya tek boşluk karakterine çevrilmez. Her birinde 1,5 satır aralığı ve paragraf sonrası 7 punto bulunur.
03. `TÜRK PATENT VE MARKA KURUMU`: tek paragraf, ortalı, Arial 11, kalın, 1,5 satır aralığı, paragraf sonrası **8 pt**.
04. `YENİDEN İNCELEME VE DEĞERLENDİRME DAİRESİ BAŞKANLIĞI’NA`: tek paragraf, ortalı, Arial 11, kalın, 1,5 satır aralığı, paragraf sonrası **22 pt**.
05. `Başvuru Nu.`: kaynaktaki **iki gerçek tab** korunur; `Başvuru Nu.\t\t: 2025/...`. Etiket kalın, başvuru değeri normal, Arial 11. **2,5 satır aralığı** ve paragraf sonrası **2 pt**.
06. `Başvuru Sahibi`: kaynaktaki **bir gerçek tab** korunur; `Başvuru Sahibi\t: ...`. Etiket kalın, şirket adı normal, Arial 11. **2,5 satır aralığı** ve paragraf sonrası **2 pt**.
07. `Referansımız`: kaynaktaki **iki gerçek tab** korunur; `Referansımız\t\t: ...`. Etiket kalın, ana dosya referansı normal, Arial 11. **2,5 satır aralığı** ve paragraf sonrası **22 pt**. Bu aralık iptal edilemez ve sonraki paragrafla birleştirilemez.
08. Ana dosya referansı Word içinde kullanılır. İşlem/görüş referansı yalnız dosya adı ve iş kimliğindedir; `700851` ana dosya referansı yerine yazılamaz.
09. Giriş paragrafı: iki-yan hizalı, Arial 11, normal, 1,5 satır, paragraf sonrası **9 pt**. Giriş doğrudan ret kararına itiraz gerekçesine geçer; `Başvuru istemlerinde herhangi bir değişiklik yapılmamaktadır.` veya buna benzer iç süreç bildirimi **KESİNLİKLE YASAK**.
10. `D1: ...`, `D2: ...` bibliyografik satırları ayrı tutulur, kısaltma ve belge kimliği kalın. Tek D satırlı kaynakta belgeden sonraki boşluk **13 pt**. Birden fazla D satırında ara satırlar 2–3 pt, son D satırından sonra 13 pt korunur. D belgesi adı doğrulanır.
11. Buluşun konu özeti normal paragraf, iki-yan hizalı, Arial 11, 1,5 satır; kaynak örnekte paragraf sonrası **11 pt**.
12. Genel savunma gövdesi normal Arial 11, iki-yan hizalı, 1,5 satır aralığı, paragraf sonrası **7 pt**. Paragraflar tek metin hâlinde birleştirilmez. Metne gereksiz boş satır karakteri veya `\n\n` eklenerek görünüm taklit edilemez.
13. Bölüm/alt bölüm başlıkları Arial 11 **kalın**, iki-yan hizalı, 1,5 satır aralığı; paragraf öncesi **14 pt**, sonrası **7 pt**. Savunma metninin tamamını kalın yapmak FAIL; başlıkları normalleştirmek FAIL.
14. Kaynakta başlık öncesi ayrı `w:p` bulunan geçişler gerçek **boş paragraf** olarak korunur. Örnek: özgün resim alt yazısından sonra D1 başlığı; D1 resminden sonra istem matrisi; yeni ana teknik konuya geçişler; sonuç başlığından önceki bölüm geçişi. İki başlığın metnini bitiştirmek FAIL.
15. İstem unsurlarının sıralı karşılaştırmasından önce, kaynak örneğin 20 numaralı paragrafındaki gibi **ayrı bir boş paragraf** bulunur. Her `1. ... — VAR/YOK.` kaleminin teknik unsur / VAR–YOK işareti **kalın**, delil açıklaması normal; her kalem paragraf sonrası **7 pt**. Kaynakta olmayan VAR/YOK ifadeleri delilsiz uydurulmaz.
16. Şekil görüntüsü kaynak PDF/DOCX'ten alınır, orijinal teknik içerik korunur. Şekil paragrafı ortalıdır; altındaki özgün kaynak adı ve `Şekil ...` açıklaması ayrı, Arial 11 kalın, ortalı, paragraf sonrası **12 pt**. Şekil–başlık arasındaki fiziki boşluk atlanmaz.
17. Şekillerde sonradan eklenmiş kırmızı işaretleme **yalnız hangi parçayı gösterdiği açık ve ayrı kaynakta doğrulanmışsa** yapılır. Belirsiz/yanlış daire, yeniden çizim, kaynağın başka unsurunu gösterecek işaretleme FAIL. Gerekirse hiçbir işaretleme kullanılmaz.
18. Tarifnameden **birebir dayanak alıntısı kalın**, önceki açıklama ve alıntı sonrasındaki teknik savunma normal. Gerçek basılı sayfa/satır yalnız kaynak Word'den çıkarılmış doğrulanmış PDF ile verilir. Tahminî sayfa/satır FAIL.
19. İtiraz metninin tamamında ve tüm Word stillerinde belge düzeltme/dil denetimi **Türkçe (Türkiye), `tr-TR`**. `w:lang`, temel stil `w:rPrDefault` ve DOCX çekirdek `dc:language` kontrol edilir. Arayüzde İngilizce görünmesi FAIL.
20. `Sonuç ve Talep` kalın başlığı öncesi 14 pt, sonrası 7 pt. Sonuç uzun itiraz tekrarlarından oluşmaz: **en çok iki kısa paragraf**. İkinci paragrafın **son cümlesi** aşağıdaki değişkenli kapanışla tam eşleşir.
21. İmza `Saygılarımızla,` ve `DESTEK PATENT A.Ş.` olarak korunur: her ikisi Arial 11 kalın, 1,5 satır; sırasıyla paragraf sonrası **5 pt** ve **6 pt**. Kurumsal imza metni keyfî değiştirilmez.
22. Taslakta bulunmayan `E-39616753-110-...` uzun evrak kimlik satırı, gereksiz usul numarası veya `Başvuru istemlerinde herhangi bir değişiklik...` açıklaması konmaz. İlgili kararın **tarihi** ve başvuru numarası yeterli giriş bağlamını oluşturur.
23. Kategoriye göre koşullu bölüm üretimi: **FM için yalnız yenilik ve teknik katkı**; **patent için yenilik ve buluş basamağı** (raporda bu itirazlar varsa). FM metninde `buluş basamağı` ibaresinin hiçbir yerde geçmemesi denetlenir. Sanayiye uygulanabilirlik olumsuz değilse ek savunma uydurulmaz.
24. Nihai görsel kontrol Word sayfalarının **tamamında** yapılır: logo/metin veya altbilgi çakışması, kopmuş alt yazı, yalnız başlık kalmış sayfa, gereksiz büyük boşluk, aralık ve tablo/şekil kayması varsa FAIL; Word örneğinin fiziksel boşluk ve satır yapısı doğrulanmadan teslim edilmez.

## 2. Zorunlu kısa kapanış (parametrik)
**Faydalı model:** “Tüm bu açıklamalarımız uyarınca `{BASVURU_NO}` numaralı başvurunun yenilik kriteri taşıdığı ortaya konmuş olup, ret kararının Kurulunuzca yeniden değerlendirilmesini arz ederiz.”
**Patent (yenilik + buluş basamağı itirazı varsa):** “Tüm bu açıklamalarımız uyarınca `{BASVURU_NO}` numaralı başvurunun yenilik ve buluş basamağı kriteri taşıdığı ortaya konmuş olup, ret kararının Kurulunuzca yeniden değerlendirilmesini arz ederiz.”
Patentte yalnız başka kriterler ret konusu ise talep, gerçek ret gerekçelerine uyarlanır; savunulmamış bir kriterin sağlandığına dair kaynaksız sonuca gidilmez. Son cümleden sonra şablondaki kurumsal imza gelir.

## 3. Şablon seçimi
- Otomatik kural bağlayıcı belge: kullanıcı tarafından **en son onaylanan** `Kurum kararına itiraz_700801.docx`. Eski `Görüş Metni_696523_YIDK_REV2.docx` bu yeni sürüme üst değildir.
- Bu dosya **FM ret kararı** örneği olduğundan biçim ve sabit resmî hitap patent için de kullanılabilir; fakat patentte ayrı yenilik/buluş basamağı savunma blokları oluşturulur.
- Şablondaki **2023/007417**, **ÖMER HALAÇ**, **696523**, **D1 US8860362B2** ve bunlara ilişkin teknik metin yeni dosyaya kaynak olarak taşınmaz.
- Kurallar repo içinde kayıtlı olsa da gerçek DOCX şablonunun makine tarafından otomatik yüklenmesi ve bu iş türünün uygulama menüsünden başlatılması için **ayrı kod entegrasyonu** gerekir. Kaydetmek çalışır entegrasyon kanıtı değildir.
