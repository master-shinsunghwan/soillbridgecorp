FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV WORKHUB_HOST=0.0.0.0
ENV WORKHUB_PORT=8765
ENV WORKHUB_DATA_DIR=/data

WORKDIR /app

COPY requirements.txt /app/requirements.txt
RUN apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates rclone tesseract-ocr tesseract-ocr-eng tesseract-ocr-kor \
    && rm -rf /var/lib/apt/lists/*
RUN pip install --no-cache-dir -r /app/requirements.txt

COPY scripts /app/scripts
COPY templates /app/templates
COPY static /app/static

RUN mkdir -p /data

EXPOSE 8765

CMD ["python", "/app/scripts/workhub_delivery_app.py"]
