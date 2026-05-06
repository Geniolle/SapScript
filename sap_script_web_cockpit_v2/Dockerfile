FROM python:3.12-slim

WORKDIR /srv/sap-script-web

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY web_api ./web_api

ENV DATA_DIR=/data
ENV WORKER_TOKEN=change-me

EXPOSE 8000

CMD ["uvicorn", "web_api.main:app", "--host", "0.0.0.0", "--port", "8000"]
