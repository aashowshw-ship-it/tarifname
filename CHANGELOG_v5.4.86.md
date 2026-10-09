# Patent Atölyesi v5.4.86 — 2026-10-09.v78

## Görüş: Microsoft Word satır kayması / yanlış sayfa-satır dayanağı düzeltmesi

- **Kök neden:** v5.4.85, DOC/DOCX'ten LibreOffice aracılığıyla elde ettiği PDF'yi fiziksel kaynak kabul ediyor, anotasyonu ve son doğrulamayı aynı PDF'deki aynı geometrik hesaptan yapıyordu. Microsoft Word ile LibreOffice satır düzeninin farklı olması durumunda her iki otomatik kontrol aynı yanlış dayanağa PASS verebiliyordu.
- **Yeni zorunlu `word_origin_pdf` teslim kapısı:** DOC/DOCX tarifname varsa kullanıcının o **aynı dosyadan Microsoft Word üzerinden PDF olarak dışa aktardığı** metin katmanlı PDF gereklidir. Sırf dahili LibreOffice export'u yetkili sayfa/satır kaynağı sayılamaz. Kaynak dosya / yetkili PDF metinlerinin bağımsız olarak çözümlenen içerikleri karşılaştırılır; uyumsuzsa indirme durur.
- **İstem revizyonu sonrası:** kullanıcıca onaylanmış SON Markup DOCX için ayrıca SON Markup'ın Word'den dışa aktarılmış PDF'si gerekir; orijinal PDF veya temiz sürüm kullanılamaz. Görüşün her yeniden üretim, revizyon ve indirme kontrolünde bu otorite yeniden denetlenir.
- **Öncelik:** doğrudan PDF yüklenmiş tarifnamede yalnız onun basılı fiziksel numaraları kullanılır. TXT için mevcut fiziksel PDF hazırlama süreci korunur. Word'de kayma varsa otomatik +/-1/+/-2 düzeltme yapılmaz; gerçek otorite PDF gerekli olup eksikliğinde FAIL-CLOSED uygulanır.
- **Regresyon:** 699151'deki gerçek alıntı ifadesi kullanılarak 5–6 ve 7–8 düzenlerinin farklı PDF'lerde çıkabileceğini gösteren testler eklendi; Word kökeni doğrulanmayan LibreOffice PDF'si ile sahte PASS engellendi.
- **Şablon/diğer iş kuralları:** görüs açılışı, kapanışı, X/Y değerlendirme akışı, istem revizyonu karar kapıları, Markup/Temiz kontrolü, şablonlar ve Türkçe dosya adları değiştirilmedi.

**Önemli sınır:** Bir dosyanın gerçekten Microsoft Word'den dışa aktarılmış olduğunu yalnız PDF içeriğinden kriptografik olarak ispatlamak mümkün değildir. Arayüz kullanıcıdan bu belgeyi açıkça ister; belge eşleşmesi ve fiziksel anchor hesabı kodla ayrıca doğrulanır. Microsoft Word dışa aktarımı yoksa doğru kabul edilmez.
