FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=7860

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 7860

CMD ["sh", "-c", "gunicorn wsgi:application --bind 0.0.0.0:${PORT} --workers 2 --timeout 300"]
