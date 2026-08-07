# Patent Atölyesi v5.2

Bu paket, mevcut Render/GitHub tabanlı **Patent Atölyesi** uygulamasının 07.08.2026 tarihli güncel tam sürümüdür.

## GitHub'a yükleme

ZIP'i açın ve içindeki dosyaların tamamını mevcut GitHub deposunun **ana dizinindeki** dosyalarla değiştirin. `tarifname-main-v5.2-final` klasörünü ikinci bir alt klasör olarak yüklemeyin.

Depo kökünde en az şu dosyalar doğrudan görünmelidir:

- `app.py`
- `rules.py`
- `RULES_MEMORY.md`
- `Dockerfile`
- `render.yaml`
- `requirements.txt`
- `packages.txt`
- `Tarifname_181176_template.docx`
- `Gorus_metni_696809_template.docx`
- `On_Arastirma_Raporu_181612_template.docx`

GitHub web arayüzünde **Add file → Upload files** ile bütün dosyaları yükleyin, aynı isimli dosyaların değiştirilmesine izin verin ve commit edin. Render servisi `autoDeploy: true` ise yeni commit otomatik yayınlanır.

## Render ortam değişkenleri

- `OPENAI_API_KEY`
- `OPENAI_MODEL`

`render.yaml` varsayılan model değerini `gpt-5.6` olarak taşır. Hesabınızda farklı model adı kullanılıyorsa Render Environment ekranından değiştirin.

## v5.2'de görüş akışı

Görüş bölümü artık doğrudan Word üretmez.

1. Dosyalar yüklenir.
2. **1. Raporu analiz et** çalıştırılır.
3. Sistem rapordaki itirazları, önceki görüşü, tarifnameyi, X/Y dokümanlarını ve varsa müşteri bilgisini birlikte analiz eder.
4. İstem revizyonu gerekmiyorsa mevcut istemlerle **Görüş metnini oluştur** düğmesi açılır.
5. Revizyon gerekiyorsa önce istem bazında öneriler gösterilir: istem no, gerekçe, tarifname dayanağı, mevcut ifade, önerilen ifade.
6. Kullanıcı isterse revizyon önerilerini ek talimatla yeniden analiz ettirir.
7. Kullanıcı revizyonu onaylarsa `.docx` kaynak tarifname üzerinde iki çıktı üretilir:
   - gerçek OOXML Track Changes içeren `Düzenlenen_tarifname_track_changes_<referans>.docx`
   - kabul edilmiş değişiklikleri içeren `Düzenlenen_tarifname_temiz_<referans>.docx`
8. Kullanıcı revize istemleri son kez onayladıktan sonra görüş Word dosyası oluşturulur.
9. Kullanıcı revizyon önerisini gördükten sonra açıkça revizyonsuz devam etmeyi de seçebilir.

Track Changes, bütün istem paragrafını silip yeniden eklemek yerine mümkün olan en küçük ifade/kelime düzeyinde uygulanır.

## v5.2'de tarifname kuralları

- `Tarifname_181176_template.docx` bağlayıcı şablondur.
- Kaynaktaki bütün teknik bilgi, özellikle önceki teknik, teknik problem/çözüm, unsurlar, yöntem adımları, formüller, tablolar, deneysel sonuçlar, alternatifler, kullanım senaryoları ve teknik etkiler eksiksiz aktarılır.
- Kullanıcıya sunulan tarifname metninde `BBF` veya `buluş bildirim formu` gibi kaynak-form atıfları kullanılmaz.
- Sistem ve yöntem istemleri birlikteyse başlık da uygun biçimde `... Sistemi ve Yöntemi` olur.
- `REFERANS NUMARALARI` unsur adlarında yalnızca ilk kelimenin ilk harfi büyük yazılır; teknik kısaltmalar korunabilir.
- Detaylı açıklamadaki yöntem adımları madde işaretli yazılır ve numara adımın sonunda `(1001)` biçiminde yer alır.
- Sistem/yöntem bağımsız istemlerindeki ayrı unsur ve adımlar madde işaretlidir.
- `İSTEMLER` ve `ÖZET` ayrı yeni sayfalardan başlar; başlıklar kalındır.
- `Buluşun bir gerçekleştirilmesinde` kullanılmaz; `Buluşun bir yapılanmasında` kullanılır.
- Önceki teknik bölümündeki müşteri anlatımı eksiksiz korunur; patent literatürü bunun yerine geçmez.
- Tarifname arayüzünde ayrıca `Özel talimat/not` alanı yoktur. Mevcut/revize tarifname varsa doğrudan yüklenir; ayrıca Var/Yok sorusu sorulmaz.
- Şekiller seçimi literatür araştırmasından önce gösterilir.

## Tip 3 ön araştırma

- Araştırma kesim tarihi kullanılır.
- Tam 10 doküman belirlenir.
- Tek satır `TotalPatent arama sorgusu: ... or ...` üretilir.
- Kullanıcının bulduğu benzer dokümanlar nihai D1/D2 analizine eklenir.
- `Buluş basamağı var / Buluş basamağı yok / Otomatik belirle` seçimi korunur.
- D1/D2 özellik tabloları aynı özellik listesini kullanır ve yalnızca `+` / `-` işaretleri içerir.
- `On_Arastirma_Raporu_181612_template.docx` bağlayıcı şablondur.

## Bağlayıcı şablonlar

- Tarifname: `Tarifname_181176_template.docx`
- Görüş: `Gorus_metni_696809_template.docx`
- Tip 3: `On_Arastirma_Raporu_181612_template.docx`

## Kural sürümü

Uygulamadaki kuralların tek yürütme kaynağı `rules.py` dosyasıdır. İnsan tarafından okunabilir kayıt `RULES_MEMORY.md` içindedir.

Kural sürümü: `2026-08-07.v4`

## Yerel çalıştırma

```bash
pip install -r requirements.txt
export OPENAI_API_KEY="..."
streamlit run app.py
```

Windows PowerShell:

```powershell
pip install -r requirements.txt
$env:OPENAI_API_KEY="..."
streamlit run app.py
```
