# AGENTS.md – Paperless Mail Consumer

## Projektübersicht

Dieses Projekt ist ein Python-Dienst, der ein Microsoft 365 Outlook-Postfach per **Microsoft Graph API** überwacht und E-Mail-Anhänge automatisch an eine **Paperless-ngx** Instanz weiterleitet.

Der Dienst läuft als **systemd-Service** auf einer Ubuntu-VM, auf der Paperless-ngx in einem Docker-Container betrieben wird.

### Kernfunktionen

- Überwacht einen konfigurierten Outlook-Ordner (`_ecoDMS`) auf neue Nachrichten
- Lädt unterstützte Anhänge (PDF, JPEG, PNG, TIFF) per Paperless REST API hoch
- Wartet auf den asynchronen Task-Status der Paperless Consumption Pipeline
- Verschiebt erfolgreich verarbeitete Mails nach `_ecoDMS/verarbeitet`
- Verschiebt fehlgeschlagene Mails (z.B. Duplikate) nach `_ecoDMS/fehlerhaft`
- Markiert verarbeitete Mails als gelesen
- Legt die Unterordner `verarbeitet` und `fehlerhaft` automatisch an, falls nicht vorhanden

---

## Repository

Das Repository liegt in einer privaten **Forgejo**-Instanz im lokalen Netzwerk. Es ist nicht öffentlich erreichbar.

---

## Projektstruktur

```
/
├── consumer.py                      # Hauptskript
├── test_consumer.py                 # Tests (pytest)
├── paperless-consumer.service       # systemd Service-Datei
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
AZURE_CLIENT_SECRET=dein-secret
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
USER_EMAIL=deine@email.de
PAPERLESS_URL=http://localhost:8000
PAPERLESS_TOKEN=dein-paperless-token
MAIL_FOLDER=_ecoDMS
INBOX_TAG_ID=2
POLL_INTERVAL=300
TASK_TIMEOUT=120
TASK_INTERVAL=3
```

**Wichtig:** Die `.env` Datei enthält Secrets und darf nie ins Repository eingecheckt werden.

---

## Abhängigkeiten

```
msal
requests
python-dotenv
```

Installation:
```bash
python3 -m pip install msal requests python-dotenv --break-system-packages
```

---

## Deployment

Der Dienst läuft als systemd-Service unter dem User `oliver` auf der Ubuntu-VM `wd9-vm-docker-p01`.

Service-Datei: `/etc/systemd/system/paperless-consumer.service`

Relevante Befehle:
```bash
sudo systemctl status paperless-consumer
sudo systemctl restart paperless-consumer
journalctl -u paperless-consumer -f
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
- `paperless-consumer.service` (systemd) anpassen wenn sich Start-Parameter oder Abhängigkeiten ändern
- Installationsschritte in `README.md` aktuell halten

### Was du nicht tun darfst

- Die `.env` Datei anlegen, verändern oder ins Repository einchecken
- Secrets, Tokens oder Passwörter in den Code schreiben
- Den systemd-Service direkt manipulieren
- Produktive Daten (Mails, Paperless-Dokumente) verändern oder löschen
- Abhängigkeiten hinzufügen ohne Kommentar im Code warum sie nötig sind

### Coding-Konventionen

- Sprache: **Python 3.12+**
- Logging: immer über das `logging` Modul, nie `print()`
- Konfiguration: ausschließlich über `.env` und `os.getenv()`, keine Hardcoded-Werte
- Fehlerbehandlung: Exceptions im Hauptloop abfangen und loggen, Dienst darf nicht abstürzen
- Funktionen klein halten, eine Funktion = eine Aufgabe
- Kommentare auf **Deutsch**

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
- Der Outlook-Ordner `_ecoDMS` ist ein Unterordner von `Posteingang`, nicht Top-Level. Die Funktion `get_folder_id()` durchsucht deshalb auch alle Unterordner.
- Anhänge sind in der Graph API base64-kodiert und müssen vor dem Upload dekodiert werden.
