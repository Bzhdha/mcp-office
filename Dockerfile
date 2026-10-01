FROM python:3.12-slim
RUN useradd -r -u 10001 app && mkdir -p /out /data/catalog && chown app /out /data/catalog
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY server.py .
COPY catalog/ catalog/
USER app
ENV OUTPUT_DIR=/out CATALOG_DIR=/data/catalog PORT=8000 PYTHONUNBUFFERED=1 PYTHONDONTWRITEBYTECODE=1
EXPOSE 8000
HEALTHCHECK CMD python -c "import urllib.request;urllib.request.urlopen('http://127.0.0.1:8000/health')"
CMD ["python","server.py"]
