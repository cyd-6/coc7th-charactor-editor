FROM python:3.12-slim-trixie

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    LANG=C.UTF-8

WORKDIR /app

RUN apt-get update \
    && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends libreoffice-calc fontconfig tzdata \
    && rm -rf /var/lib/apt/lists/* \
    && useradd --create-home --uid 10001 coc7

COPY requirements.lock.txt ./
RUN pip install --no-cache-dir -r requirements.lock.txt

COPY app.py ./
COPY coc7_card ./coc7_card
COPY static ./static
COPY assets ./assets
COPY THIRD_PARTY_NOTICES.txt ./
RUN mkdir -p /usr/local/share/fonts/coc7 \
    && cp assets/fonts/NotoSansSC-Regular.ttf /usr/local/share/fonts/coc7/ \
    && fc-cache -f

USER coc7
EXPOSE 8765
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8765/api/health', timeout=3)"

CMD ["python", "-m", "uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8765", "--no-access-log"]
