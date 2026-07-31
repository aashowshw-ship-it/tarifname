# Tarifname Atölyesi — Web MVP

BBF dosyasını analiz eden, unsur ve işlem adımlarını kullanıcıya kontrol ettiren, isteğe bağlı patent literatürü araştırması yapan ve `Tarifname_181176` düzeninde DOCX çıktısı oluşturan Streamlit web uygulaması.

## En hızlı yayınlama: Render

1. Bu klasörü özel bir GitHub deposuna yükleyin.
2. Render üzerinde **New > Blueprint** seçin ve depoyu bağlayın.
3. `OPENAI_API_KEY` değerini Render ortam değişkenlerinde tanımlayın.
4. Deploy tamamlandığında HTTPS web adresi oluşur.

`render.yaml` ve `Dockerfile` hazırdır.

## Streamlit Community Cloud

1. Klasörü GitHub'a yükleyin.
2. Streamlit Community Cloud'da `app.py` dosyasını seçin.
3. Secrets alanına şunu ekleyin:

```toml
OPENAI_API_KEY = "sk-..."
OPENAI_MODEL = "gpt-5.6"
```

Eski `.doc` dosyaları için `packages.txt` içindeki `antiword` ve `libreoffice-writer` paketlerinin kurulması gerekir. Platform bu paketleri kurmazsa BBF'leri `.docx` olarak yükleyin.

## Yerel test

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
export OPENAI_API_KEY="sk-..."  # Windows PowerShell: $env:OPENAI_API_KEY="sk-..."
streamlit run app.py
```

## Güvenlik

- API anahtarını arayüzden istemez; sunucu ortam değişkeninden okur.
- BBF metni uygulama belleğinde işlenir ve varsayılan olarak veri tabanına kaydedilmez.
- Üretim ortamında özel GitHub deposu, kullanıcı doğrulaması ve şirket içi erişim katmanı kullanılması önerilir.

## Uygulanan tarifname kuralları

- BBF unsur numaraları aynen korunur.
- Referans numaralarından önceki bölümlerde parantezli numara kullanılmaz.
- Yöntem adımı referansları işlem sonunda `(1001)` biçimindedir.
- Sistem istemindeki unsurlar sırayla ve önce tanımlanan unsurlarla ilişkili kurulur.
- Alt istemler ana istemi tekrarlamaz.
- İngilizce terimler ilk kullanımda Türkçesiyle açıklanır ve devamında Türkçesi kullanılır.
- Şablondaki kırmızı/mavi açıklama paragrafları korunur.
