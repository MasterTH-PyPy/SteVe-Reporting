FROM python:3.12-slim

WORKDIR /app

# System-Pakete (docker CLI für docker exec)
RUN apt-get update && apt-get install -y --no-install-recommends \
    docker.io \
 && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Datenbank initialisieren und App starten
CMD ["sh", "-c", "mkdir -p /app/data && python -c 'from app import init_db; init_db()' && gunicorn -w 2 -b 0.0.0.0:5000 app:app"]
