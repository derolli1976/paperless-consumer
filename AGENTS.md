# AGENTS.md – Paperless Mail Consumer

## Projektübersicht

Dieses Projekt ist ein Python-Dienst, der ein Microsoft 365 Outlook-Postfach per **Microsoft Graph API** überwacht und E-Mail-Anhänge automatisch an eine **Paperless-ngx** Instanz weiterleitet.

Der Dienst läuft als **Docker-Container** auf einer Ubuntu-VM, auf der Paperless-ngx ebenfalls in einem Docker-Container betrieben wird.

### Kernfunktionen

- Überwacht einen konfigurierten Outlook-Ordner (per `MAIL_FOLDER`) auf neue Nachrichten
- Lädt unterstützte Anhänge (PDF, JPEG, PNG, TIFF) per Paperless REST API hoch
- Konvertiert Word-Dokumente (.doc, .docx) automatisch zu PDF vor dem Upload (via LibreOffice)
- Wartet auf den asynchronen Task-Status der Paperless Consumption Pipeline
- Verschiebt erfolgreich verarbeitete Mails nach `<MAIL_FOLDER>/verarbeitet`
- Verschiebt fehlgeschlagene Mails (z.B. Duplikate) nach `<MAIL_FOLDER>/fehlerhaft`
- Markiert verarbeitete Mails als gelesen
- Legt die Unterordner `verarbeitet` und `fehlerhaft` automatisch an, falls nicht vorhanden
- Protokolliert alle Import-Ergebnisse in einer JSON-Lines Log-Datei (`import.log`)
- Versendet täglich zu einer konfigurierbaren Uhrzeit eine HTML-Zusammenfassung per E-Mail mit Übersicht über erfolgreiche und fehlgeschlagene Importe
- Alle Zeitangaben in deutscher Zeitzone (Europe/Berlin)

---

## Repository

Das Repository liegt in einer privaten **Git**-Instanz. Es ist nicht öffentlich erreichbar.

---

## Projektstruktur

```
/
├── consumer.py                      # Hauptskript
├── test_consumer.py                 # Tests (pytest)
├── Dockerfile                       # Container-Image Definition
├── docker-compose.yml               # Docker Compose Konfiguration
├── .dockerignore                    # Dateien die nicht ins Image gehören
├── requirements.txt                 # Python-Abhängigkeiten mit Versionen
├── data/                            # Volume-Mount für import.log (nicht im Repo)
├── .env                             # Konfiguration (nicht im Repository)
├── .env.example                     # Beispielkonfiguration (im Repository)
├── AGENTS.md                        # Diese Datei
└── README.md                        # Installationsanleitung
```

---

## Konfiguration

Alle Einstellungen werden über eine `.env` Datei gesetzt. Siehe `.env.example`:

```env
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_SECRET=<your-secret>
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
USER_EMAIL=user@example.com
PAPERLESS_URL=http://localhost:8000
PAPERLESS_TOKEN=<your-paperless-token>
MAIL_FOLDER=<your-mail-folder>
INBOX_TAG_ID=2
POLL_INTERVAL=300
TASK_TIMEOUT=120
TASK_INTERVAL=3
SUMMARY_HOUR=9
SUMMARY_RECIPIENT=
IMPORT_LOG_FILE=/app/data/import.log
```

| Variable | Beschreibung | Default |
|---|---|---|
| `SUMMARY_HOUR` | Uhrzeit (Stunde, 0–23, Europe/Berlin) für den täglichen E-Mail-Versand | `9` |
| `SUMMARY_RECIPIENT` | Empfänger der Zusammenfassung (leer = `USER_EMAIL`) | – |
| `IMPORT_LOG_FILE` | Pfad zur JSON-Lines Log-Datei für Import-Ergebnisse | `/app/data/import.log` |

**Wichtig:** Die `.env` Datei enthält Secrets und darf nie ins Repository eingecheckt werden.

---

## Abhängigkeiten

Siehe `requirements.txt` für die gepinnten Versionen:

```
msal
requests
python-dotenv
```

Installation:
```bash
pip3 install -r requirements.txt
```

---

## Deployment

Der Dienst läuft als Docker-Container auf der Ziel-VM.

### Docker Compose (empfohlen)

```bash
# Erststart
cp .env.example .env
nano .env
docker compose build
docker compose up -d

# Update nach git pull
docker compose up -d --build

# Logs anzeigen
docker compose logs -f

# Stoppen
docker compose down
```

### Remote-Deployment via SSH

```bash
ssh user@<VM-HOST> "cd /opt/paperless-consumer && git pull && docker compose up -d --build"
```

---

## Hinweise für den Copilot Agent

### Was du tun darfst

**Code:**
- `consumer.py` erweitern, refactoren und Fehler beheben
- Neue Funktionen hinzufügen (z.B. Unterstützung weiterer Dateitypen, Retry-Logik, Benachrichtigungen)

**Tests:**
- Tests in `test_consumer.py` schreiben (pytest)
- Externe API-Calls in Tests mit `unittest.mock` mocken
- Testabdeckung für Fehlerfälle (Duplikate, Timeouts, fehlende Anhänge) sicherstellen

**Dokumentation:**
- `.env.example` und `README.md` aktualisieren wenn sich die Konfiguration ändert
- `AGENTS.md` aktualisieren wenn sich Projektstruktur, APIs oder Deployment ändern
- Inline-Kommentare auf Deutsch pflegen

**Deployment:**
- `Dockerfile` und `docker-compose.yml` anpassen wenn sich Container-Konfiguration ändert
- Installationsschritte in `README.md` aktuell halten

### Was du nicht tun darfst

- Die `.env` Datei anlegen, verändern oder ins Repository einchecken
- Secrets, Tokens oder Passwörter in den Code schreiben
- Produktive Daten (Mails, Paperless-Dokumente) verändern oder löschen
- Abhängigkeiten hinzufügen ohne Kommentar im Code warum sie nötig sind

### Coding-Konventionen

- Sprache: **Python 3.12+**
- Logging: immer über das `logging` Modul, nie `print()`
- Konfiguration: ausschließlich über `.env` und `os.getenv()`, keine Hardcoded-Werte
- Fehlerbehandlung: Exceptions im Hauptloop abfangen und loggen, Dienst darf nicht abstürzen
- Funktionen klein halten, eine Funktion = eine Aufgabe
- Kommentare auf **Englisch**

### Externe APIs

**Microsoft Graph API**
- Basis-URL: `https://graph.microsoft.com/v1.0/`
- Auth: OAuth2 Client Credentials Flow via `msal.ConfidentialClientApplication`
- Scope: `https://graph.microsoft.com/.default`
- Relevante Endpoints:
  - `GET /users/{email}/mailFolders` – Ordnerliste
  - `GET /users/{email}/mailFolders/{id}/childFolders` – Unterordner
  - `GET /users/{email}/mailFolders/{id}/messages?$expand=attachments` – Nachrichten mit Anhängen
  - `PATCH /users/{email}/messages/{id}` – Mail als gelesen markieren
  - `POST /users/{email}/messages/{id}/move` – Mail verschieben
  - `POST /users/{email}/mailFolders/{id}/childFolders` – Unterordner anlegen
  - `POST /users/{email}/sendMail` – E-Mail senden (für tägliche Zusammenfassung)

**Paperless-ngx REST API**
- Basis-URL: Konfigurierbar via `PAPERLESS_URL`
- Auth: Token-Header `Authorization: Token {PAPERLESS_TOKEN}`
- Relevante Endpoints:
  - `POST /api/documents/post_document/` – Dokument hochladen, gibt Task-UUID zurück
  - `GET /api/tasks/?task_id={uuid}` – Task-Status abfragen
  - Task-Status-Werte: `PENDING`, `STARTED`, `SUCCESS`, `FAILURE`

### Bekannte Eigenheiten

- `post_document` gibt immer HTTP 200 zurück, auch wenn das Dokument später als Duplikat erkannt wird. Den eigentlichen Status erst über den Task-Endpoint abfragen.
- Duplikat-Fehler erscheinen als `FAILURE` im Task mit "duplicate" in der `result` Property.
- Der konfigurierte Outlook-Ordner kann ein Unterordner von `Posteingang` sein, nicht zwingend Top-Level. Die Funktion `get_folder_id()` durchsucht deshalb auch alle Unterordner.
- Anhänge sind in der Graph API base64-kodiert und müssen vor dem Upload dekodiert werden.
- Die tägliche Zusammenfassung wird einmal pro Tag versendet. Der Versand-Status wird als `summary_sent`-Sentinel in der Log-Datei persistiert und überlebt damit auch Neustarts des Dienstes.
- Das Import-Log (`import.log`) verwendet JSON-Lines-Format: pro Zeile ein JSON-Objekt mit `type`, `ts`, `file`, `subject`, `status` (`success`/`failed`), `error`.
- `_read_log_entries_since_last_summary()` liest alle Import-Einträge nach dem letzten `summary_sent`-Sentinel, unabhängig vom Datum. Damit werden auch Importe erfasst, die nach dem letzten Versand am Vortag noch eingegangen sind.
