# Patent Atölyesi v5.4.82 — 2026-10-06

Ruleset: `2026-10-06.v74`

1. Sistem ve yöntem şekilleri kesin ayrıldı; sistem şekline 1001+ yöntem referansı bindirilemez.
2. Şekillerde `tek kutu = tek ana referans` kuralı zorunlu hale getirildi. Aynı taşıyıcıdaki ek unsurlar ayrı tek-ref callout/kılavuz ile gösterilir.
3. Şekil denetim JSON'u `one_reference_per_box` ve `reference_box_occupancy` alanlarıyla sertleştirildi; doğrulama bu PASS olmadan otomatik düzeltmeyi kabul etmez.
4. Sistem şekline kaynakta yanlışlıkla yöntem ref'i konmuşsa audit otomatik olarak `needs_edit` üretir ve temizlenmesini ister.
5. Yöntem akış şekli kutularında yalnız 1001+ yöntem ref'i bulunur; sistem unsur ref'leri yöntem şekline yazılmaz.
6. Teknik olarak eşdeğer yalnız yazı/ref katmanı farklı kaynak şekil varyantlarının tekilleştirilmesi ve son şekil açıklamalarının 1..N yeniden senkronizasyonu kuralı eklendi.
7. Çoklu-şekil sayfa yerleşiminde tek şeklin sayfa sınırında bölünmesi/kırpılması yasaklandı.
8. REFERANS NUMARALARI yöntem satırlarında sistem/cihaz parantez referanslarının görünmemesi mevcut kuralı teyit edilip sürüm notuna taşındı; bu referanslar DETAYLI AÇIKLAMA'dan itibaren gösterilir.
9. ÖNCEKİ TEKNİK problem sentezinde mekanik `... gerektirir` tekrarlarını reddeden doğal akış kuralı eklendi.
