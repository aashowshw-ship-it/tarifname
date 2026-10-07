# Patent Atölyesi v5.4.84 — 2026-10-07.v76

## Şekil-türü ayrımı ve fonksiyonel diyagram orta-yol kuralı

1. Şekil audit'i fiziksel/yapısal çizim, fonksiyonel mimari/veri akışı, karar/işlem akışı ve kanonik yöntem akışını birbirinden ayırır.
2. Fiziksel/yapısal tek unsur kutusu ve küçük çağrı etiketi için `tek unsur = tek referans` zorunluluğu aynen sürer.
3. Fonksiyonel diyagramlarda teknik anlaşılabilirliği sağlayan Türkçe metin kutuları korunabilir; süreç/karar/çıktı kutularına sırf biçim için uydurma referans eklenmez.
4. Kaynak veya kullanıcı onaylı üst-seviye `group_container` aynı mantıksal taşıyıcıdaki birden fazla alt birimi açıkça gruplayabiliyorsa çoklu referans aynı grup kutusunda korunabilir. Fiziksel tek unsur kutularında bu istisna uygulanmaz.
5. Boş kutu nihai şekil olarak kabul edilmez.
6. Yöntem akış şekli değişmez: her 1001+ adım kendi ayrı kutusunda, yalnız kendi referansı ile, tam bir kez ve kanonik sırada gösterilir.
7. Referans düzeltme ve ikinci görsel doğrulama promptları yeni şekil türü ayrımına göre güncellendi; fiziksel şekiller gereksiz gevşetilmedi.

## Doğrulama

- Yeni regresyon testi: `test_v5484_functional_figure_exception.py`.
- Tam test paketi çalıştırılmalıdır.
