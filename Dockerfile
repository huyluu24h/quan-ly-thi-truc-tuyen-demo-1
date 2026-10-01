FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    APP_DATA_DIR=/app/storage

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
       libreoffice-core \
       libreoffice-writer \
       libreoffice-calc \
       libreoffice-impress \
       libreoffice-draw \
       texlive-xetex \
       texlive-latex-recommended \
       fonts-dejavu-core \
       fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY . ./
RUN mkdir -p /app/storage

EXPOSE 10000

CMD ["gunicorn", "-c", "gunicorn.conf.py", "wsgi:app"]

