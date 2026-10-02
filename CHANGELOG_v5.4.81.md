# v5.4.78 — Bağımlı İstem ve Önceki Teknik Hardening

Bu sürüm yeni tarifname oluşturma akışında gerçek dosyada yakalanan dört sınıf hatayı fail-closed kapatır: bağımlı istem tekrarları, Türkçe istem yönelme ekleri, çok-unsurlu alt istemlerin okunabilirliği ve ÖNCEKİ TEKNİK giriş sırası.

## Değişiklikler

1. `%92` kelime örtüşme eşiği kaldırıldı; sistem ve yöntem bağımlı istemleri tam ata zincirine karşı teknik özellik imzalarıyla denetleniyor.
2. `1004/1005` gibi ana yöntem adımlarının bağımlı istem olarak tekrar edilmesi artık kesin FAIL.
3. `İstem 9’e` / `İstem 10’e` gibi yanlış Türkçe yönelme ekleri sayı okunuşuna göre otomatik normalize edilip ayrıca deterministik doğrulanıyor.
4. İki veya daha fazla yeni referanslı unsur getiren bağımlı istem gerçek bullet yapısı olmadan geçemiyor; Word renderer gerçek madde işareti üretiyor.
5. `bir fonksiyon olmasıdır` soyut unsur kapanışı reddediliyor; standart olmayan yazılımsal fonksiyon bileşenleri için birim/modül adlandırması yönlendiriliyor.
6. `ajanlar-arası` house-style ile `ajanlar arası` olarak normalize ediliyor.
7. ÖNCEKİ TEKNİK ilk genel paragrafı `Günümüzde ...` ile mevcut tekniği genel olarak anlatmadan özel müşteri örneklerine geçemiyor.
8. Kalite promptları bu kontrolleri yüzde-benzerliksiz, feature-by-feature mantıkla yeniden denetliyor.

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

---

# v5.4.79 — Unsur Türü Semantiği ve İstem Kapanış Yerleşimi

- Uygulama sürümü `v5.4.79`, ruleset `2026-09-25.v71`.
- Ham BBF'de genel `... fonksiyonu` olarak adlandırılmış referans unsur, açıklamada gerçekte bir birim/modül/sunucu/cihaz olarak tarif ediliyorsa referans numarası korunarak bu teknik taşıyıcı türüyle kanonikleştirilir.
- Standart özel ağ fonksiyonu adları (ör. açıkça SMF/UPF vb. standardize isimler) keyfi değiştirilmez.
- `Ajan keşif fonksiyonu`, sıra sıfatlı genel `çekirdek şebeke fonksiyonu` gibi standart özel ad olmayan genel işlev adları referans unsuru olarak fail-closed reddedilir.
- `fonksiyon üzerinde konumlandırılan/bulunan/çalışan ajan` biçimindeki soyut barındırma ilişkisi reddedilir; kaynak destekliyorsa birim/modül bünyesinde ilişki kurulması gerekir.
- Çok maddeli bağımlı istemde son madde `keep_with_next` ile kapanış paragrafına bağlanır; kapanış şablon girintisini korur ve tek başına kopamaz.
- v5.4.79 regresyonları mevcut test dosyasına eklenmiştir; dosya sayısı artırılmamıştır.
- Tam otomatik test paketi: **381 passed**, başarısız test yok. 13 uyarı yalnız Pillow `Image.getdata` deprecation uyarısıdır.


# v5.4.80 — Tip 3 Multi-Pass Prior-Art Search
- Added four-pass hidden prior-art discovery before freezing the visible top 10: GLOBAL_RECALL, GAP_SEARCH, FAMILY_NEIGHBOUR, D1_CHALLENGE.
- Preserved the existing visible UI sequence and 10+ user-document behavior.
- Added deterministic final-top10 integrity checks and D1/D2 membership/query-order validation.
- Kept research global; no jurisdiction-specific privileged pass.

# v5.4.81 — Tarifname Düzenleme Structure Hardening

- Application version: `v5.4.81`; ruleset: `2026-10-02.v73`.
- Added fail-closed single-real-paragraph rule for every `insert_paragraph_after/before` operation. One operation may not hide multiple paragraphs behind manual line breaks.
- Extended the OMML equation rule to the Tarifname Düzenleme workflow. `[[EQ: ...]]` / `[[FORMULA: ...]]` markers are converted to real Word math objects and verified in both Markup and Clean outputs.
- Added deterministic claim-family block validation. An independent claim and all of its dependent claims must remain contiguous; a later family may not be followed by a dependent claim that returns to an earlier family.
- Added system/product vs method category block validation and sequential claim-number/dependency validation.
- Added regression coverage in `test_v5432_tarifname_update.py` for paragraph structure, OMML conversion, invalid system→method→system ordering, and valid in-family insertion with downstream renumbering.
- Full regression suite: **387 passed**, 13 Pillow deprecation warnings, no failures.
