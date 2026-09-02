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
      antiword libreoffice libcairo2 fontconfig cabextract wget curl xfonts-utils ttf-mscorefonts-installer tesseract-ocr tesseract-ocr-tur tesseract-ocr-eng \
 && fc-cache -f \
 && fc-match Arial | grep -qi 'Arial' \
 && rm -rf /var/lib/apt/lists/*

# Süreçler modülünde kredi gerektirmeyen küçük yerel AI. Model yalnız başvuru
# bilgi çıkarımı/doğrulaması için kullanılır; OpenAI ile bağlantısı yoktur.
ENV PATH="/root/.local/bin:${PATH}" \
    LOCAL_AI_MODEL_PATH="/opt/models/Qwen2.5-0.5B-Instruct-IQ2_XS.gguf" \
    LOCAL_AI_CONTEXT="3072" \
    LOCAL_AI_MAX_TOKENS="1100"
RUN SKIP_CUDA=1 SKIP_ROCM=1 SKIP_VULKAN=1 sh -c 'curl -LsSf https://llama.app/install.sh | sh' \
 && mkdir -p /opt/models \
 && curl -L --fail --retry 3 --retry-delay 2 \
      -o /opt/models/Qwen2.5-0.5B-Instruct-IQ2_XS.gguf \
      'https://huggingface.co/ThomasBaruzier/Qwen2.5-0.5B-Instruct-GGUF/resolve/main/Qwen2.5-0.5B-Instruct-IQ2_XS.gguf?download=true' \
 && test -s /opt/models/Qwen2.5-0.5B-Instruct-IQ2_XS.gguf

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV PORT=10000
CMD ["sh", "-c", "streamlit run app.py --server.port=${PORT} --server.address=0.0.0.0 --server.headless=true"]
