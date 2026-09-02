FROM python:3.12-slim
WORKDIR /app

# Patent DOCX'leri çoğunlukla Arial/Times New Roman ile hazırlanıyor.
# LibreOffice bu fontlar yoksa belgeyi sessizce başka fontla render eder ve
# satır/sayfa düzeni değişebilir. Debian contrib'deki Microsoft Core Fonts
# yükleyicisi ile dönüşüm ortamında gerçek Arial/Times New Roman sağlanır.
RUN if [ -f /etc/apt/sources.list.d/debian.sources ]; then \
      sed -i 's/^Components: main$/Components: main contrib non-free non-free-firmware/' /etc/apt/sources.list.d/debian.sources; \
    elif [ -f /etc/apt/sources.list ]; then \
      sed -i 's/ main$/ main contrib non-free non-free-firmware/' /etc/apt/sources.list; \
    fi \
 && apt-get update \
 && echo 'ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true' | debconf-set-selections \
 && apt-get install -y --no-install-recommends \
      antiword libreoffice libcairo2 fontconfig cabextract wget xfonts-utils ttf-mscorefonts-installer \
 && fc-cache -f \
 && fc-match Arial | grep -qi 'Arial' \
 && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV PORT=10000
CMD ["sh", "-c", "streamlit run app.py --server.port=${PORT} --server.address=0.0.0.0 --server.headless=true"]
