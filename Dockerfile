# The web app as a container, for servers where Traefik (or another Docker proxy) already owns ports 80/443.
# scripts/deploy_vps.sh builds and runs this automatically when it finds Traefik.
FROM python:3.12-slim

WORKDIR /app
COPY pyproject.toml README.md ./
COPY src ./src
RUN pip install --no-cache-dir .
COPY .streamlit ./.streamlit
COPY app ./app

# The library and download cache live in /data: mount a host directory there.
ENV SUMMERS_DSY_DATA=/data HOME=/tmp PYTHONUNBUFFERED=1
EXPOSE 8501
HEALTHCHECK --interval=30s --timeout=5s --start-period=30s \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8501/_stcore/health')"
CMD ["streamlit", "run", "app/streamlit_app.py", "--server.address", "0.0.0.0", "--server.port", "8501", \
     "--server.headless", "true", "--browser.gatherUsageStats", "false"]
