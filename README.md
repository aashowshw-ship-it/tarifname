# Patent Atölyesi v5

Bu paket, Render üzerinde çalışan mevcut **Patent Atölyesi v4** projesinin güncellenmiş tam sürümüdür. Ayrı bir GitHub Pages projesi değildir.

## v4 deposuna nasıl yüklenir?

Bu ZIP dosyasını açın. ZIP içindeki dosyaların tamamını mevcut GitHub deposunun **ana dizinindeki** dosyalarla değiştirin.

GitHub ana dizininde doğrudan şu dosyalar görünmelidir:

- `app.py`
- `rules.py`
- `Dockerfile`
- `render.yaml`
- `requirements.txt`
- `packages.txt`
- `Tarifname_181176_template.docx`
- `Gorus_metni_696809_template.docx`
- `On_Arastirma_Raporu_181612_template.docx`

`patent-atolyesi-v5` klasörünü GitHub deposunun içine ikinci bir klasör olarak yüklemeyin. Dosyalar depo kökünde olmalıdır.

GitHub web arayüzüyle güncelleme:

1. ZIP'i bilgisayarda açın.
2. Mevcut deponun ana sayfasında **Add file → Upload files** seçin.
3. ZIP içindeki bütün dosyaları sürükleyin.
4. Aynı isimli dosyaların değiştirilmesine izin verin.
5. Commit mesajı olarak `Patent Atölyesi v5` yazın.
6. **Commit changes** düğmesine basın.

Render servisi mevcut GitHub deposuna bağlı ve `autoDeploy: true` olduğu için commit sonrasında otomatik olarak yeniden yayınlanır.

## Render ayarları

Render servisinde şu ortam değişkenleri bulunmalıdır:

- `OPENAI_API_KEY`
- `OPENAI_MODEL`

`render.yaml` varsayılan model değerini `gpt-5.6` olarak taşır. OpenAI hesabında farklı bir model adı kullanılıyorsa Render Environment ekranından değiştirilmelidir.

## v5'te eklenenler

### Tarifname

- BBF yanında mevcut/revize tarifname yükleme
- Ek teknik müşteri belgeleri yükleme
- Örnek tarifnameleri yalnızca kurgu amacıyla yükleme
- `Yalnızca yöntem` istem seçeneği
- Otomatik sistem/yöntem istem türü analizi
- BBF ham metninin doğrudan taslak oluşturma ve ikinci kalite kontrol aşamasında kullanılması
- Formül, tablo, deneysel sonuç, alternatif gerçekleştirme ve referans tablosu tamlık kontrolü
- Aynı metinle yazılmış farklı yöntem adımı numaralarının kontrolü
- Ana istemde paralel adımların kapsayıcı yazılması ve ayrıntıların tek bağımlı istemde toplanabilmesi
- Eğitim/genel aşama ile test aşamasının paralel fakat ayrı değerlendirilmesi
- İstem numaralarının kalın yazılması
- `Yöntemin gerçekleştirdiği işlem adımları` bölümünün referans tablosundan otomatik üretilmesi
- BBF veya ayrıca yüklenen görsellerden ayrı Şekiller Word dosyası oluşturma

### Tip 3 ön araştırma

- Araştırma kesim tarihi
- Tam 10 doküman zorunluluğu
- Tek satırlık `TotalPatent arama sorgusu`
- Kullanıcının benzer dokümanlarını nihai D1/D2 analizine ekleme
- `Buluş basamağı var / Buluş basamağı yok / Otomatik belirle` seçimi
- D1 ve D2 tablolarında aynı teknik özellik listesi zorunluluğu
- Yalnızca `+` ve `-` kullanımı

## Kurallar nerede?

Bütün kalıcı üretim kuralları tek kaynakta tutulur:

```text
rules.py
```

İnsan tarafından okunabilir özet ve tam kayıt:

```text
RULES_MEMORY.md
```

Arayüzde görünen kural sürümü:

```text
2026-08-04.v2
```

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

## Önemli not

DOCX ve PDF içindeki büyük görseller ayrı şekiller dosyasına alınabilir. PDF görsel çıkarımı için `PyMuPDF` bağımlılığı v5 ile eklenmiştir. Otomatik çıkarımda belge logosu veya teknik şekil olmayan büyük bir görsel bulunması ihtimaline karşı indirilen Şekiller dosyası son kez gözle kontrol edilmelidir.
