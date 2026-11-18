# Patent Atölyesi — Red Kararına İtiraz / YİDD-YİDK
**Kayıt:** 18.11.2026. **Kaynak:** Kullanıcının Word üzerinde son onaylayarak düzenlediği `Red Kararına İtiraz_700851.docx`.

## Tek geçerli biçim otoritesi
- **700851 DOCX** (özel Kitaplık): `/Patent Atolyesi/Sablonlar/Red Kararina Itiraz/Red Kararına İtiraz_700851_BAGLAYICI_ORNEK.docx`.
- Orijinal SHA-256: `a90aef7237862e2de03c2585d3a1d687b81b1ce2611c9b2fb26484fb1b6df7bb`.
- DOCX'in 92 paragraf ve 22 fiziksel boş paragraf düzeni, gerçek başlık alt çizgileri ve Word dili kuralları: **[700851_SOURCE_OF_TRUTH.md](700851_SOURCE_OF_TRUTH.md)**.
- Otomatik karşılaştırmaya esas paragraf bazlı biçim kaydı özel Kitaplıkta `700851_STYLE_MANIFEST.json`.
- 696523 ve 700801 Word'leri yalnız tarihsel örnek olup **güncel biçime üstün değildir**. Kamuya açık GitHub'a müşteri Word'ü kopyalanmaz.

## Birlikte çalışan dört kural
1. **[RULES_FM.md](RULES_FM.md)**: Faydalı model. Tek belge-bazlı yenilik, SMK 142 teknik katkı, buluş basamağı yok.
2. **[RULES_PATENT.md](RULES_PATENT.md)**: Patent. Yenilik tek belge, buluş basamağı ayrı problem/çözüm ve belgelerin birlikte değerlendirilmesi.
3. **[INPUT_SCHEMA.json](INPUT_SCHEMA.json)**: Arayüzde tüm evrakların tek seferde yüklenmesi, koşullu patent/FM belge alanları, otomatik veri çıkarımı, referanslar ve tebliğ günü.
4. **[WORKFLOW.md](WORKFLOW.md)**: Ret kararı → son istem/şekiller → rapor geçmişi → her OBJ → ayrı D matrisi → alt istem → özgün görsel → kararın kaldırılması → Word QA.

## Yazım kuralları
- Yeni itirazda **noktalı virgül yazma** ilkesi vardır. Yalnız orijinal birebir alıntıda korunması veya teknik anlam açısından zorunlu durumlarda bağımsız kontrolle istisna yapılır. Onaylı 700851 örneğine sonradan düzenleme yapılmaz.
- Şablondaki belirli başlıklar altı çizilidir, diğerleri değildir. Kullanıcının yaptığı bu seçim, bölüm/satır boşlukları ve diğer Word biçimleri **tek tek** korunur.
- Girişte gereksiz ‘Başvuru istemlerinde değişiklik yapılmamaktadır’ veya uzun `E-396...` numarası gibi usul metası yer almaz.
- Müşteriden görüş bilgisi yoksa bağımsız araştırma ve savunma devam eder. Müşteri görüşü veya teknik bilgi uydurulmaz.
- Mevcut istemde olmayan özelliği varmış gibi savunma, gerçek kaynakta bulunmayan teknik fark veya muğlak savunma dili yasaktır.
- Sonuç ve Talep en fazla iki kısa paragraf, ardından FM veya patent tipine uygun kesin kapanış ve Destek Patent imzası.

## Uygulama durumu
Bu dizindeki dosyalar **kaydedilmiş işlem kuralları ve arayüz veri sözleşmesidir**. İstenen tek ekran UI ve analiz sırası ayrıntılı olarak hazırlanmıştır. **Mevcut `app.py` menüsüne ve otomatik Word üretim motoruna bu turda bağlanmamıştır.** Çalışır ürün olduğu iddia edilmez. Entegrasyon aşamasında giriş ekranı `INPUT_SCHEMA.json` baz alınmalı, kaynak Word kullanımı özel depolama kimlik bilgileriyle yetkilendirilmeli ve `WORKFLOW.md` içindeki PASS/FAIL kapıları programatik çalıştırılmalıdır.
