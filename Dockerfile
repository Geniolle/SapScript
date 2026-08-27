FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY main.py jira_download_anexos.py jira_sheet_daemon.py sap_session.py ./
COPY workflow ./workflow
COPY Processos ./Processos

CMD ["python", "jira_sheet_daemon.py"]
