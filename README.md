# paperless-consumer

Python-Dienst, der ein Microsoft 365 Postfach überwacht und Anhänge automatisch an Paperless-ngx weiterleitet.

## Funktionen

- Überwacht einen konfigurierbaren Outlook-Ordner auf neue Nachrichten
- Lädt unterstützte Anhänge (PDF, JPEG, PNG, TIFF) per REST API in Paperless-ngx hoch
- Konvertiert Word-Dokumente (.doc, .docx) automatisch zu PDF vor dem Upload (via LibreOffice)
- Wartet auf den Abschluss des Paperless Consumption Tasks
- Verschiebt verarbeitete Mails nach `<MAIL_FOLDER>/verarbeitet`, fehlerhafte nach `<MAIL_FOLDER>/fehlerhaft`
- Markiert verarbeitete Mails als gelesen
- Legt Unterordner automatisch an, falls nicht vorhanden
- Protokolliert alle Import-Ergebnisse in einer JSON-Lines Log-Datei (`import.log`)
- Versendet täglich eine HTML-Zusammenfassung per E-Mail (konfigurierbare Uhrzeit)
- Alle Zeitangaben in deutscher Zeitzone (Europe/Berlin)

## Quick Start (Docker)

### 1. Konfiguration erstellen

```bash
cp .env.example .env
nano .env
```

Trage deine Daten ein (siehe [Konfiguration](#konfiguration)).

### 2. Container starten

```bash
docker compose build
docker compose up -d
```

### 3. Logs prüfen

```bash
docker compose logs -f
```

## Deployment (laufender Betrieb)

Nach Änderungen im Repository:

```bash
git push
```

Dann auf der Ziel-VM deployen – via SSH:

```bash
ssh user@<VM-HOST> "cd /opt/paperless-consumer && git pull && docker compose up -d --build"
```

Oder direkt auf der VM:

```bash
cd /opt/paperless-consumer
git pull
docker compose up -d --build
```

## Konfiguration

Alle Einstellungen in `.env` – siehe `.env.example` für alle verfügbaren Variablen.

| Variable | Beschreibung | Default |
|---|---|---|
| `AZURE_CLIENT_ID` | Azure AD App Client ID | – |
| `AZURE_CLIENT_SECRET` | Azure AD App Client Secret | – |
| `AZURE_TENANT_ID` | Azure AD Tenant ID | – |
| `USER_EMAIL` | Überwachte Outlook-E-Mail-Adresse | – |
| `PAPERLESS_URL` | URL der Paperless-ngx Instanz | – |
| `PAPERLESS_TOKEN` | API-Token für Paperless-ngx | – |
| `MAIL_FOLDER` | Name des zu überwachenden Outlook-Ordners | – |
| `INBOX_TAG_ID` | Paperless Tag-ID für importierte Dokumente | `2` |
| `POLL_INTERVAL` | Abfrageintervall in Sekunden | `300` |
| `TASK_TIMEOUT` | Timeout für Paperless-Task in Sekunden | `120` |
| `TASK_INTERVAL` | Abfrageintervall für Task-Status in Sekunden | `3` |
| `SUMMARY_HOUR` | Uhrzeit (Stunde, 0–23, Europe/Berlin) für den täglichen E-Mail-Versand | `9` |
| `SUMMARY_RECIPIENT` | Empfänger der Zusammenfassung (leer = `USER_EMAIL`) | – |
| `IMPORT_LOG_FILE` | Pfad zur JSON-Lines Log-Datei | `/app/data/import.log` |

> **Hinweis:** `PAPERLESS_URL` muss aus dem Container erreichbar sein. Falls Paperless-ngx
> ebenfalls auf derselben VM als Docker-Container läuft, die IP des Docker-Hosts oder das
> Docker-Netzwerk nutzen (z.B. `http://host.docker.internal:8000` oder `http://<VM-IP>:8000`).

## Container-Befehle

```bash
# Status prüfen
docker compose ps

# Logs anzeigen
docker compose logs -f

# Neustarten
docker compose restart

# Stoppen
docker compose down

# Neu bauen und starten
docker compose up -d --build
```

## Ohne Docker ausführen

```bash
pip3 install -r requirements.txt
cp .env.example .env
# .env anpassen
python consumer.py
```



