# Red Kararına İtiraz — 700851 bağlayıcı örnek ve kaynak hiyerarşisi

**Onay tarihi:** 18.11.2026. **Kaynak:** Kullanıcının kendi düzeltmeleriyle yüklediği `Red Kararına İtiraz_700851.docx`.
**Gerçek Word SHA-256:** `a90aef7237862e2de03c2585d3a1d687b81b1ce2611c9b2fb26484fb1b6df7bb`.
**Özel Kitaplık konumu:** `/Patent Atolyesi/Sablonlar/Red Kararina Itiraz/Red Kararına İtiraz_700851_BAGLAYICI_ORNEK.docx`.
**Örnek türü:** Faydalı model başvurusunun reddine karşı YİDD/YİDK itirazı. Biçim, düzen ve üslup için bağlayıcıdır. İçindeki CONRAD, numara, teknik unsurlar ve D1/D2 bilgileri yeni müşteri dosyasına aktarılamaz.

## 0. Üstünlük
1. Patent Atölyesi v5.4.93 bağlayıcı ana kuralları ve yasal gereklilikler üstündür.
2. **Bu son kullanıcı dosyası**, önceki `Kurum kararına itiraz_700801.docx` ve `Görüş Metni_696523_YIDK_REV2.docx` örneklerine göre **biçimsel üstünlüğe** sahiptir.
3. Eski `700801_EXACT_STYLE.md` içindeki bu örnekle çatışan başlık/alt çizgi/boşluk özellikleri **geçersizdir**. Yeni üretim doğrudan 700851 Word'ü okuyarak yapılır. Biçim salt Markdown'dan yeniden tasarlanmaz.
4. Bu sayfa, özel Kitaplıkta saklanan DOCX'in yerini tutmaz. Kamuya açık GitHub'a müşteri dosyasının ham Word'ü yüklenmez.

## 1. Gerçek kaynaktan ölçülmüş biçim koşulları
- Word: 11 sayfa, 92 ana gövde paragrafı, 1 section. İki başlangıç paragrafı **fiziksel olarak boştur**. Logo, üstbilgi ve altbilgi korunur. Dosya `SHA-256` eşleşmesiyle doğrulanır.
- Kurum başlığı iki satır ve **ortalanmış kalın**. İlk başlığın ardından 8 pt, ikinci başlığın ardından 22 pt paragraf boşluğu bulunur.
- `Başvuru Nu.`, `Başvuru Sahibi`, `Referansımız` bilgi satırları: etiketler bold, değerler normal, gerçek tab karakterleri korunur. 2,5 satır aralığı. İlk iki satır sonrası 2 pt, referans satırının ardından **22 pt**. Başvuru numarası, başvuru sahibi ve ana dosya referansı resmi kaynaklardan doldurulur.
- Açılış paragrafı iki yana yaslı, Arial 11 normal, 1,5 satır ve **9 pt sonrası**. Ret kararının tarihini ve esas itirazı belirtir. `Başvuru istemlerinde herhangi bir değişiklik yapılmamaktadır.`, örnek iç süreç cümlesi ve uzun `E-... ` karar numarası girişte bulunamaz.
- Her `D1: ...`, `D2: ...` satırı ayrı paragrafta. D etiketi kalın, belge adı normal. D2'den sonra örnek par. **10**, gerçek boş paragraftır.
- Buluş özeti normal metin. Şekiller gerçek kaynak çizimleri, ortalı, bağımsız kalın şekil alt yazılarıyla sunulur. Hatalı kırmızı halka çizilmez. Yalnız kesin olarak doğrulanmış işaretleme yapılır.
- **Ana bölüm başlıkları** Arial 11 kalın, iki yana yaslı, paragraf **öncesi 14 pt**, **sonrası 7 pt**, 1,5 satır. Başlıkların alt çizgisi **kaynak DOCX'teki her bir başlığın kendi durumuna göre** uygulanır. Tüm başlıklara topluca alt çizgi eklenmez veya topluca kaldırılmaz.
- **Özgün örnekte ALT ÇİZİLİ başlıklar**: `D1 ... Yenilik Değerlendirmesi` (par.15), `Mevcut 1 Numaralı İstemin ... D1 ile Karşılaştırılması` (par.21), `Ekran Ünitesi ... İşlevsel Ayrım` (par.34), `İki Ayrı Sök-Tak ...` (par.39), `Ret Kararında Önceki İtiraza Verilen Cevabın Değerlendirilmesi` (par.64), `D1 ve D2 ... Birlikte ... Hatalı Olması` (par.70), `2 ve 3 Numaralı Bağımlı İstemlerin ...` (par.75), `Faydalı Modelin ... Son Değerlendirme` (par.81).
- **Özgün örnekte altı ÇİZİLMEYEN ama KALIN başlıklar**: `D2 ... Yenilik Değerlendirmesi` (par.47), `Mevcut 1 Numaralı İstemin ... D2 ile Karşılaştırılması` (par.53), `Sonuç ve Talep` (par.86). Bu gerçek farklılık aynen korunur. Yeni belgede farklı D-belgesi sayısı varsa başlık tipine göre stil örneğinden kopya alınır, bütün doküman kalın/altçizili yapılmaz.
- Kaynaktaki gerçek boş paragraf indeksleri: **0, 1, 10, 12, 14, 18, 20, 23, 31, 33, 38, 44, 46, 50, 52, 55, 63, 69, 74, 80, 85, 89**. Her biri kendi konumu ve paragraf boşluğu/öncesi-sonrası aralığı korunarak yeniden üretilir. Yeni bölüm eklenirse eşdeğer **bölüm geçişi** yapısı kullanılır. Salt `\n\n` karakterleriyle boş paragraf taklidi YASAK.
- Ana savunma paragrafları Arial 11 normal, 1,5 satır, iki yana yaslı, genel olarak **7 pt sonrası**. Görsel altyazıları ortalı, Arial 11 kalın, genellikle 12 pt sonrası. Alt yazı ayrı Word paragrafıdır.
- İstem unsurlarının `1. ... — VAR/YOK.` başındaki teknik unsur ve etiket kalın, devamındaki kaynak incelemesi normal. Her D1/D2 için ayrı ve tam istem unsur matrisi gerekir.
- Özgün tarifnameden alınan **birebir** dayanaklar kalın, dayanak öncesi/sonrası normal. Başlık/alıntı dışı paragraf topluca kalınlaştırılamaz.
- `Sonuç ve Talep` en çok iki kısa paragraf ve ardından başvuru türüne uygun **tek kesin kapanış cümlesi**. Sonrasında şablonun `Saygılarımızla,` ve `DESTEK PATENT A.Ş.` (kalın) imzaları. Belgenin dil kontrolü, Word core properties ve ilgili stiller `tr-TR`.
- **Noktalı virgül ilkesi:** Yeni YİDK metinleri noktalı virgülsüz oluşturulur. Noktalı virgül ancak özgün kaynaktan birebir alıntıda zorunluysa korunur. Normal savunma metninde cümleyi uygun yerde noktayla veya virgülle ayır. Denetim `; ` / `;` bulgularını listeler ve teknik anlam korunarak giderilene kadar metin dili kontrolünden geçmez. **Onaylanan kaynak Word'ün içeriği değiştirilmeyecek**, yeni metinlere bu kural uygulanacaktır.
- Kesin ve iddialı ancak yalnız belgelere dayalı hukuk dili. `kısmen`, `olabilir`, `yapılabilir`, `düşünülmektedir`, `muhtemelen`, `...olmalıdır` gibi tereddüt ve prosedür tavsiyesi ifadeleri yeni savunmaya girmez.

## 2. Hukuki sonucu ürün türüne göre ayır
- **Faydalı model:** “Tüm bu açıklamalarımız uyarınca {BAŞVURU_NO} numaralı başvurunun yenilik kriteri taşıdığı ortaya konmuş olup, ret kararının Kurulunuzca yeniden değerlendirilmesini arz ederiz.” Faydalı modelde `buluş basamağı` bu kapanışta veya olumlu kriter savunmasında YOK.
- **Patent:** Yenilik ve buluş basamağı ret gerekçeleriyse “Tüm bu açıklamalarımız uyarınca {BAŞVURU_NO} numaralı başvurunun yenilik ve buluş basamağı kriteri taşıdığı ortaya konmuş olup, ret kararının Kurulunuzca yeniden değerlendirilmesini arz ederiz.” Gerçek ret başka gerekçedense savunulan gerekçeye uygun son cümle kurulur.
- **Patent/Faydalı model ikisinde de** sonuç kapsamına girmeyen, olumlu rapor kriterlerini tartışmaya gereksiz yer ayırma.

## 3. YİDD savunması içerik zorunlulukları
1. Ret kararındaki her maddi gerekçeye ayrı, gerekçe içeriğiyle doğrudan ilişkili cevap.
2. Son **resmî istem** lafzı ve gerçek teknik unsur bağlantıları. Tarifnamede açıklanmış ama istemde olmayan özellikleri istemin yenilik farkı olarak ileri sürme. İddia edilen eşdeğerliği işlev, konum ve fiziksel bağlantıyla çürüt.
3. D1, D2 ve her kaynak için **tek belge tek yenilik incelemesi**. Unsur sıralı VAR/YOK matrisi kaynakla doğrulanır. D1+D2 birleştirme sadece patentte **buluş basamağında** tartışılır. FM'de birleştirme yapılıp yenilik yokluğu kurulmuşsa prosedürel hata somutlaştırılır.
4. Önceki görüşteki hangi iddiaya Kurumun ne cevap verdiği ayrı analiz edilir. Aynı savunma aynen kopyalanmaz. Müşteri açıklaması yoksa araştırma dosyasıyla çalışılır.
5. Bağımlı istemler ayrı ele alınır. FM'de SMK 142/2 teknik katkı itirazı fiziksel işlev ve delille karşılanır.
6. İstemleri rutin olarak değiştirme, Markup/Clean hazırlama, değişiklik yapılmadığını gereksiz cümleyle ilan etme.
7. Son çift kontrol: kaynak alıntısı+sayfa/satır, resmi istem, D şekli, başlık/boş paragraflar, çakışma/taşma ve `tr-TR`/noktalı virgül.

## 4. Mühendislik ve teslim
Yalnızca `Görüş Metni_696523` veya `Kurum kararına itiraz_700801` üzerinden genel paragraf boşluğu tahmini yapılmaz. Yeni kaynak Word'ün OOXML'i gerçek biçim otoritesidir. İstem D karşılaştırma matrisi ve ayrı patent/FM promptları `WORKFLOW.md` içindedir. Arayüz tek ekran sözleşmesi `INPUT_SCHEMA.json` içindedir. Kaynak belgesi veya doğrulanmış görsel yoksa hüküm veya alıntı uydurma. Eksik kaynak ve NOT_RUN = FAIL.
