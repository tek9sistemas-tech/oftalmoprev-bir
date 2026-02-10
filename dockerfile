FROM python:3.12-slim

WORKDIR /app

# instala deps do sistema (opcional, mas seguro)
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
 && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PORT=5000
EXPOSE 5000

CMD ["gunicorn", "BIR:app", "--bind", "0.0.0.0:5000", "--access-logfile", "-", "--error-logfile", "-"]
