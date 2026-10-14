FROM python:3.11-slim
RUN apt-get update && apt-get install -y --no-install-recommends \
      libreoffice-writer-nogui poppler-utils fonts-noto-cjk fonts-arphic-ukai fonts-arphic-uming \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
RUN cp fonts-tw.conf /etc/fonts/conf.d/60-tw-alias.conf && fc-cache -f
ENV PORT=8080
CMD exec gunicorn --bind :$PORT --workers 1 --threads 4 --timeout 180 main:app
