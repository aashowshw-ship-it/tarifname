# Tarifname Atölyesi - doğrudan çıktı sürümü

## GitHub'a yükleme
ZIP'i açın ve içindeki altı dosyayı deponun ana dizinine yükleyin:
- app.py
- Tarifname_181176_template.docx
- requirements.txt
- Dockerfile
- render.yaml
- packages.txt

## Render ayarı
Servis > Environment bölümünde:
- Key: `OPENAI_API_KEY`
- Value: `sk-...`

Kaydederken **Save, rebuild and deploy** seçin.

## Kullanım
BBF dosyasını yükleyin, çıktı adını ve istem yapısını seçin. Literatür araştırması seçilirse belirtilen sayıda doğrulanmış patent dokümanı önceki teknik bölümüne otomatik eklenir. Ara ekran gösterilmez; işlem sonunda yalnızca Word indirme düğmesi çıkar.
