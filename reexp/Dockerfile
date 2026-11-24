FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PORT=7860

# OCR para PDFs escaneados.
RUN apt-get update \
 && apt-get install -y --no-install-recommends tesseract-ocr \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py index.html ./

EXPOSE 7860
CMD ["python", "app.py"]
