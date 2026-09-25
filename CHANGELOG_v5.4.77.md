# v5.4.77 — Atomik İstem Revizyonu / Insertion-First

Bu sürüm, görüş çalışmasındaki istem revizyonunu serbest metin yeniden-yazımı olmaktan çıkarıp ayrı ve fail-closed bir kalite alt sistemine dönüştürür.

## Uygulanan davranış

1. **Uzman itiraz envanteri:** İlk analiz her maddi itirazı `OBJ-1`, `OBJ-2`... kimliğiyle çıkarır. Teknik revizyon, en az bir gerçek itiraz kimliğine bağlanmadan uygulanamaz.
2. **Insertion-first:** Teknik revizyonda varsayılan işlem `INSERT`tir. Mevcut istem kelimeleri korunabiliyorsa `DELETE`/`REPLACE` yasaktır. Üslup, akıcılık veya yeniden formülasyon tek başına silme gerekçesi değildir.
3. **Kelime/ibare düzeyi anti-rewrite:** Zorunlu silme/değiştirme yalnız yerel kelime veya kısa ibare düzeyinde kalabilir. Aynı mevcut kelimeleri silip başka sırada yeniden ekleyen gizli cümle-rewrite deterministik olarak engellenir.
4. **Doğrudan as-filed dayanak:** Her ek teknik özellik ayrı `direct_support` kaydıyla as-filed tarifname/istemden birebir dayanak alır. Müşteri görüş formu ve D dokümanları istem değişikliğinin as-filed dayanağı değildir.
5. **Etki kartı:** Her öneri açıklık etkisi, kapsam etkisi, önceki teknik/D-doküman etkisi ve kalan risk ile kullanıcıya sunulur.
6. **Aday revizyon karşılaştırması:** Birden fazla güvenli aday varsa ilk analiz aday üretir. Bağımsız auditor bütün adayları `itirazı tam karşılama → doğrudan dayanak → minimum müdahale → gereksiz daraltmadan kaçınma → açıklık` sırasıyla karşılaştırır ve seçilmiş adayı ayrıca doğrular.
7. **Bağımsız Amendment Auditor:** Auditor doğrudan dayanağı, uzmanın itirazını gerçekten karşılayıp karşılamadığını, insertion-first/minimum edit durumunu, yeni belirsizliği, kapsam etkisini, D-doküman karşısındaki teknik değeri ve kalan riski kontrol eder. `partial/no` sonuç Markup üretimini durdurur.
8. **Deterministik Track Changes:** Yapay zekâ `w:ins/w:del` üretmez. Onaylı lokal `old_text/new_text` hedefi kod tarafından atomik farklara çevrilir. Değişmeyen Word metni redline içine alınmaz.
9. **Üçlü bütünlük:** `Markup reddedilmiş görünüm = kaynak`, `Markup kabul edilmiş görünüm = Temiz`, `Markup redline = onaylı atomik plan` eşitlikleri zorunludur. Claim amendment indirmeleri ayrıca merkezi final compliance gate'e tabidir.
10. **Revizyon karar kapısı:** Revizyon gerekmiyorsa görüş otomatik ilerler. Revizyon gerekiyorsa kullanıcı `Önerilen istem revizyonunu uygula` veya `Mevcut istemlerle devam et` kararını vermeden görüş üretilemez.
11. **Revizyon sonrası kaynak yenileme:** Revizyon uygulanırsa önceki görüş/cache ve orijinal fiziksel sayfa/satır indeksi geçersizleşir. Görüş ve teslim anındaki sayfa/satır doğrulaması son Markup'tan arka planda yeniden üretilen PDF üzerinden çalışır.
12. **Görüş görünür sırası:** `giriş → D1/D2/... bibliyografik satırlar → İstemlerde Yapılan Değişiklikler ve Dayanakları → esas D savunmaları → sonuç`. Bu sıra final Word üzerinde deterministik olarak doğrulanır.
13. **Resmi dil kapısı:** `minimum ölçüde değiştirilmiştir`, `sistemimizce uygun görülmüştür` gibi vekil-içi süreç/meta ifadeleri resmi görüşte reddedilir.
14. **Müşteri görüş formu:** Kaynakla ve tarifnameyle doğrulanan müşteri savunma noktaları görüşte kullanılmaya devam eder; ancak istem revizyonu dayanağına dönüştürülemez.

## Yeni / güncellenen teknik bileşenler

- `claim_amendment.py`: atomik fark, insertion-first, candidate ve auditor receipt kapıları.
- `gorus_audit.py`: Markup/Temiz/kaynak üçlü bütünlük, görünür görüş sıra ve meta-dil kapıları.
- `app.py` / `app_core.py`: itiraz envanteri, aday revizyon, bağımsız auditor, kullanıcı karar ekranı, revizyon sonrası cache ve fiziksel satır kaynağı yenilemesi.
- `rules.py`: v5.4.77 / ruleset `2026-09-25.v69` bağlayıcı kuralları ve `claim_amendment` final compliance türü.
- `README.md` / `RULES_MEMORY.md`: yeni davranışın okunabilir kayıtları; v5.4.71 eski otomatik-akış davranışı tarihsel ve v5.4.77 ile geçersiz olarak işaretlendi.
- `test_v5477_atomic_claim_amendment.py`: insertion-only, gereksiz rewrite, geniş phrase rewrite, gerçek OOXML minimum redline, candidate auditor, Word sıra/meta dil, final compliance ve son-Markup sayfa/satır kaynağı regresyonları.

## Doğrulama

- Python syntax compile: PASS
- Tam otomatik test paketi: **373 passed**
- Test çıktısındaki 13 uyarı, Pillow `Image.getdata` deprecation uyarılarıdır; test başarısızlığı değildir.
