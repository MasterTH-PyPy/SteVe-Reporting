# SteVe Ladehistorie – Web-Frontend

Flask-Webanwendung zur Anzeige der Ladehistorie aus einer SteVe-OCPP-Instanz.

## Berechtigungsstufen

| Rolle | Zugriff |
|---|---|
| **admin** | Benutzerverwaltung (SteVe-Mitarbeiter zuweisen) |
| **buchhaltung** | Alle Chips, alle Zeiträume, Excel-Export |
| **userplus** | Eigene Chips, Monats- + Tagesansicht |
| **user** | Eigene Chips, nur Monatsansicht |

## Installation

```bash
git clone https://github.com/MasterTH-PyPy/SteVe-Reporting.git
cd SteVe-Reporting

# 1. Konfiguration anlegen
cp .env.example .env
nano .env                      # Werte eintragen, siehe unten

# 2. Optional: lokale Anpassungen (Port, Netzwerk, …)
cp docker-compose.override.yml.example docker-compose.override.yml

# 3. Starten
docker compose up -d --build
docker compose logs -f
```

Anschließend: http://localhost:5000  
Standard-Login: `admin` / `admin` → **Passwort sofort ändern!**

> Das Session-Cookie ist als `Secure` markiert. Außer über `localhost` funktioniert
> der Login daher nur per **HTTPS** (z. B. hinter Nginx Proxy Manager).

## Konfiguration (`.env`)

Alle Zugangsdaten stehen in der Datei `.env` im Projektverzeichnis. Sie wird über
`env_file: .env` in den Container geladen und ist in `.gitignore` eingetragen –
sie landet also **nie im Git-Repo** und bleibt bei `git pull` unverändert.
Vorlage ist `.env.example`.

| Variable | Pflicht | Standard | Beschreibung |
|---|---|---|---|
| `DOCKER_CONTAINER` | ja | `steve-db-1` | Name des MariaDB-Containers von SteVe |
| `DB_NAME` | ja | `steve` | Datenbankname (= `db.schema` in SteVes `application-docker.properties`) |
| `DB_USER` | ja | `steve` | DB-Benutzer (= `db.user`) |
| `DB_PASSWORD` | ja | – | DB-Passwort (= `db.password`) |
| `STEVE_VERSION` | nein | `current` | `current` für SteVe 3.7+ (das EVSE-Schema ab 3.13 wird automatisch erkannt), `legacy` für 3.6 |
| `SECRET_KEY` | ja | – | Langer Zufallswert für die Flask-Sessions |

Beispiel:

```ini
DOCKER_CONTAINER=steve-db-1
DB_NAME=steve
DB_USER=steve
DB_PASSWORD=geheim
STEVE_VERSION=current
SECRET_KEY=3f9c…(64 Hex-Zeichen)
```

### Werte ermitteln

```bash
# Name des DB-Containers
docker ps --format '{{.Names}}' | grep db

# DB-Name, -User und -Passwort aus der SteVe-Installation
grep -E '^db\.' <steve-ordner>/src/main/resources/application-docker.properties

# SECRET_KEY erzeugen
python3 -c "import secrets; print(secrets.token_hex(32))"
```

**Hinweise**

- Werte **ohne Anführungszeichen** eintragen – Docker Compose übernimmt sie sonst wörtlich.
- Fehlt `SECRET_KEY` oder steht noch der Platzhalter drin, erzeugt die App bei jedem
  Start einen zufälligen Key (Warnung im Log). Dann sind nach jedem Neustart alle
  Benutzer abgemeldet.
- Nach Änderungen an der `.env` reicht `docker compose up -d` – der Container wird
  automatisch neu erstellt (ein reines `restart` liest die `.env` **nicht** neu ein).
- Ergebnis prüfen: `docker compose config` zeigt die zusammengeführte Konfiguration
  (Achtung: inklusive Passwort im Klartext).

### Lokale Anpassungen

Port, zusätzliche Netzwerke oder weitere Umgebungsvariablen gehören in die
`docker-compose.override.yml` (Vorlage: `docker-compose.override.yml.example`),
nicht in die `docker-compose.yml`. Auch diese Datei ist git-ignoriert.

## Zugriff auf die SteVe-Datenbank

Die App fragt die SteVe-DB per `docker exec` im MariaDB-Container ab. Dafür wird
der **Docker-Socket** (`/var/run/docker.sock`) in den Container gemountet.
Ein gemeinsames Docker-Netzwerk mit SteVe ist **nicht** nötig, ebenso keine
freigegebenen DB-Ports.

## Update

```bash
git pull
docker compose up -d --build
```

`.env`, `docker-compose.override.yml` und die interne Benutzerdatenbank
(Volume `reporting-data`) bleiben dabei erhalten.

### SteVe-Update

Wird SteVe neu installiert (z. B. neue Version in einem neuen Ordner), ändert sich
oft der Name des DB-Containers. Dann `DOCKER_CONTAINER` in der `.env` anpassen
und `docker compose up -d` ausführen.
