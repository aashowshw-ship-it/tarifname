# RED KARARINA İTİRAZ — ŞABLON, ARAYÜZ VE QA
Sürüm: 2026-11-18 / ilk taslak, henüz mevcut Streamlit menüsüne bağlanmadı.

## Bağlayıcı Word taslağı
- **Örnek tür:** YİDD/YİDK'ye ret kararına itiraz, **belge kararına itiraz değildir**.
- Kullanıcının geçmiş çalışma örneği: `Görüş Metni_696523_YIDK_REV2.docx` (kullanıcının ChatGPT özel dosya Kitaplığında). Bu dosyadan **yalnız biçim, antet, bölüm hiyerarşisi, hukuki hitap, görünür paragraf/başlık stili, istem matrisi ve özgün görsel yerleşimi** alınır. Başka dosyaya özgü teknik açıklama, kişi, adres, istem, tarih, D1 alıntısı taşınmaz.
- Örnek kamuya açık GitHub deposuna kişisel/müşteri belgesi olarak aktarılmaz. Uygulamanın otomatik kullanımı için dosya yetkili depolama alanında sabit/okunabilir Word şablon konumuna konulmalı ve yükleme/sürüm kontrolü eklenmelidir.
- **Faydalı model:** aynı resmî YİDK biçimindeki örneğin FM yenilik-only bölüm yapısı. **Patent:** ayrı şablon varyantında yenilik ve buluş basamağı / varsa diğer ret gerekçeleri için ayrı başlıklar. FM metnine buluş basamağı paragrafı sızması FAIL.
- Antet/telif/görsel kaynakları birebir korunur; resmî makam, karar, başvuru sahibi, ana dosya referansı, görüş/işlem referansı dosyaya göre doldurulur. Ana dosya referansı Word metadata içindir; **görüş referansı çıktı adında** kullanılır: `Red Kararına İtiraz_<görüş_ref>.docx`.
- Başlıklar **bold**, doğrudan tarifname dayanak alıntıları **bold**, normal savunma/atıf metinleri normal. Bağımsız özgün örnekten stil/hiza/boşluk kaybı FAIL. Tüm metni bold yapmak veya tüm başlıkları normal yapmak FAIL.
- D1/D2 özgün resim sayfalarına yapılacak işaretleme yalnız teknik açıklayıcı kırmızı daire/ok gibi görsel üstkatman; özgün resim kaynağı değişmez, orijinal sayfa ve işaretlenen unsur kayda alınır.

## Tek ekrandaki dosya alanları
1. Ret kararı PDF — zorunlu.
2. Redde esas son tarifname/istem DOC/DOCX/PDF — zorunlu.
3. Şekiller PDF/DOCX — varsa.
4. Son rapor (FM araştırma, patent inceleme) — zorunlu.
5. Önceki görüş ve varsa istem değişiklikleri — varsa.
6. Karardaki önceki teknik kaynakları D1/D2/... — zorunlu kaynak doğrulaması.
7. Üçüncü kişi görüşü — varsa.
8. Müşteri görüş bilgisi — isteğe bağlı, yoksa otomatik bağımsız savunma.
9. Ana dosya ref / görüş ref — zorunlu; tebliğ günü — bilinmiyorsa süre sonu hesaplanmaz.
- Otomatik: ret kararından başvuru numarası, başvuru sahibi, tür, tarih, D belgeleri, OBJ konuları çıkar; kullanıcıdan tekrar sorma.
- Uygulama iş aşamaları: **dosya envanteri → başvuru türü ayrımı → son istem/karar envanteri → D1/D2 ayrı teknik inceleme → gerekçe bazlı savunma → kaynak/şekil kontrolü → Word QA**. Bunlar kullanıcının tek seferlik formundan sonra iç adımlardır, ayrı anket değildir.

## Sunum öncesi kapılar
- `source_complete`, `protect_type`, `legal_route`, `notification_deadline`, `claims_verbatim`, `novelty_single_source`, `inventive_step_if_patent`, `obj_response_coverage`, `dependent_claims`, `direct_source_quotes`, `images_original`, `template_fixed_text`, `font_spacing_bold`, `word_render_qa`, `independent_second_review`.
- Eksik asli delil / resmi belge: son belgeyi uydurma. Yalnız gerekli belgeyi iste veya taslak çalışmayı **açıkça eksik dayanaklı** işaretle.
- Tek belge kaynaklı yenilik analizi ile buluş basamağı kombinasyonunu karıştırma. SMK 142/2 teknik katkı açıklamasını yenilik tespitinin yerine geçirme.
- V5.4.93 bağlayıcı mevcut README/RULES_MEMORY/rules.py üstündür. Bu dosyalar bu yeni iş türü için **taslak kural tanımıdır**, kaynak kodunda iş türü henüz kayıtlı değilse otomatik yürütüldüğünü iddia etme.
