FROM python:3.12-slim
WORKDIR /app
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 EPGURU_DATA=/data
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY epg_processor.py fix_dates.py ./
COPY templates ./templates
COPY server ./server
COPY web ./web
RUN useradd -r -u 1000 epguru && mkdir -p /data && chown epguru /data
USER epguru
EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=5s CMD python -c "import urllib.request;urllib.request.urlopen('http://127.0.0.1:8000/healthz')"
CMD ["uvicorn", "server.main:app", "--host", "0.0.0.0", "--port", "8000", "--proxy-headers", "--forwarded-allow-ips", "*"]
