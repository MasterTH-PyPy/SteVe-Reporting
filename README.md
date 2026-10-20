# SteVe Ladehistorie – Web-Frontend

Flask-Webanwendung zur Anzeige der Ladehistorie aus einer SteVe-OCPP-Instanz.

## Berechtigungsstufen

| Rolle | Zugriff |
|---|---|
| **admin** | Benutzerverwaltung (SteVe-Mitarbeiter zuweisen) |
| **buchhaltung** | Alle Chips, alle Zeiträume, Excel-Export |
| **userplus** | Eigene Chips, Monats- + Tagesansicht |
| **user** | Eigene Chips, nur Monatsansicht |

## Start

```bash
# Im selben Verzeichnis wie das SteVe-Compose:
docker compose -f docker-compose.yml up -d --build
```

Anschließend: http://localhost:5000  
Standard-Login: `admin` / `admin` → **Passwort sofort ändern!**

## Konfiguration

Umgebungsvariablen in `docker-compose.yml` anpassen:

| Variable | Beschreibung |
|---|---|
| `DOCKER_CONTAINER` | Name des MariaDB-Containers |
| `DB_NAME` | Datenbankname (`steve`) |
| `DB_USER` | DB-Benutzer |
| `DB_PASSWORD` | DB-Passwort |
| `STEVE_VERSION` | `current` (3.7+) oder `legacy` (3.6) |
| `SECRET_KEY` | Zufälliger langer String für Flask-Sessions |

## Netzwerk

Der Reporting-Container benötigt:
1. Zugriff auf den **Docker-Socket** (`/var/run/docker.sock`) für `docker exec`
2. Dasselbe **Docker-Netzwerk** wie der SteVe-DB-Container

Standard-Netzwerkname: `steve_default` – prüfen mit `docker network ls`.
