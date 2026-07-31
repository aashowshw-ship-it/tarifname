# Patent Atölyesi v4

Tek arayüzde üç modül:

1. Tarifname oluşturma
2. Araştırma / inceleme raporuna karşı görüş hazırlama
3. Tip 3 ön araştırma raporu

## GitHub'a yükleme

ZIP'i açın. ZIP'in içindeki dosyaların tamamını GitHub deposunun ana dizinine yükleyin. ZIP dosyasını doğrudan yüklemeyin.

Ana dizinde şunlar görünmelidir:

- app.py
- Dockerfile
- render.yaml
- requirements.txt
- packages.txt
- Tarifname_181176_template.docx
- Gorus_metni_696809_template.docx
- On_Arastirma_Raporu_181612_template.docx

## Render

- Render > New > Blueprint
- GitHub deposunu seçin
- OPENAI_API_KEY değerini Environment bölümüne girin
- OPENAI_MODEL değerini hesabınızda erişiminiz bulunan model adıyla değiştirin
- Deploy Blueprint

## Araştırma akışı

- BBF yüklenir
- Sistem global araştırmayla en yakın 10 dokümanı bulur
- Önerilen D1/D2 ekranda gösterilir
- Kullanıcıya kendi benzer dokümanları olup olmadığı sorulur
- PDF/ZIP vb. yüklenirse yeniden analiz edilir
- Nihai D1/D2 belirlenir
- Yenilik ve buluş basamağı sonucu bağlanarak Word raporu oluşturulur
