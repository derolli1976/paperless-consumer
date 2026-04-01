import os
import json
import time
import base64
import logging
import datetime
import requests
import msal
from dotenv import load_dotenv

load_dotenv("/opt/paperless-consumer/.env")

CLIENT_ID       = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET   = os.getenv("AZURE_CLIENT_SECRET")
TENANT_ID       = os.getenv("AZURE_TENANT_ID")
USER_EMAIL      = os.getenv("USER_EMAIL")
PAPERLESS_URL   = os.getenv("PAPERLESS_URL")
PAPERLESS_TOKEN = os.getenv("PAPERLESS_TOKEN")
MAIL_FOLDER     = os.getenv("MAIL_FOLDER")
INBOX_TAG_ID    = int(os.getenv("INBOX_TAG_ID", "2"))
POLL_INTERVAL      = int(os.getenv("POLL_INTERVAL", "300"))
TASK_TIMEOUT       = int(os.getenv("TASK_TIMEOUT", "120"))
TASK_INTERVAL      = int(os.getenv("TASK_INTERVAL", "3"))
SUMMARY_HOUR       = int(os.getenv("SUMMARY_HOUR", "9"))
SUMMARY_RECIPIENT  = os.getenv("SUMMARY_RECIPIENT")
IMPORT_LOG_FILE    = os.getenv("IMPORT_LOG_FILE", "/opt/paperless-consumer/import.log")

SUPPORTED_TYPES = [
    "application/pdf",
    "image/jpeg",
    "image/png",
    "image/tiff",
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)


def get_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise Exception(f"Token-Fehler: {result.get('error_description')}")
    return result["access_token"]


def graph_get(token, path):
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/{path}",
        headers={"Authorization": f"Bearer {token}"},
    )
    r.raise_for_status()
    return r.json()


def graph_patch(token, path, payload):
    r = requests.patch(
        f"https://graph.microsoft.com/v1.0/{path}",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=payload,
    )
    r.raise_for_status()


def graph_post(token, path, payload):
    r = requests.post(
        f"https://graph.microsoft.com/v1.0/{path}",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=payload,
    )
    r.raise_for_status()
    return r.json()


def get_folder_id(token, folder_name):
    # Erst Top-Level prüfen
    data = graph_get(token, f"users/{USER_EMAIL}/mailFolders?$top=100")
    for folder in data.get("value", []):
        if folder["displayName"] == folder_name:
            return folder["id"]

    # Dann Unterordner aller Top-Level Ordner durchsuchen
    for parent in data.get("value", []):
        children = graph_get(
            token,
            f"users/{USER_EMAIL}/mailFolders/{parent['id']}/childFolders?$top=100"
        )
        for folder in children.get("value", []):
            if folder["displayName"] == folder_name:
                return folder["id"]

    raise Exception(f"Ordner '{folder_name}' nicht gefunden")


def get_or_create_subfolder(token, parent_id, subfolder_name):
    data = graph_get(
        token,
        f"users/{USER_EMAIL}/mailFolders/{parent_id}/childFolders?$top=100"
    )
    for folder in data.get("value", []):
        if folder["displayName"] == subfolder_name:
            return folder["id"]
    log.info(f"Unterordner '{subfolder_name}' wird angelegt...")
    result = graph_post(
        token,
        f"users/{USER_EMAIL}/mailFolders/{parent_id}/childFolders",
        {"displayName": subfolder_name},
    )
    return result["id"]


def get_messages(token, folder_id):
    data = graph_get(
        token,
        f"users/{USER_EMAIL}/mailFolders/{folder_id}/messages"
        f"?$expand=attachments&$top=50",
    )
    return data.get("value", [])


def upload_to_paperless(filename, file_bytes, content_type):
    headers = {"Authorization": f"Token {PAPERLESS_TOKEN}"}
    files = {"document": (filename, file_bytes, content_type)}
    data = {"tags": [INBOX_TAG_ID]}
    r = requests.post(
        f"{PAPERLESS_URL}/api/documents/post_document/",
        headers=headers,
        files=files,
        data=data,
    )
    r.raise_for_status()
    # Paperless gibt die Task-UUID als String zurück (mit Anführungszeichen)
    task_id = r.text.strip().strip('"')
    return task_id


def wait_for_task(task_id):
    """
    Wartet auf den Abschluss eines Paperless Tasks.
    Gibt (True, None) bei Erfolg zurück,
    oder (False, "Fehlermeldung") bei Fehler.
    """
    headers = {"Authorization": f"Token {PAPERLESS_TOKEN}"}
    elapsed = 0
    while elapsed < TASK_TIMEOUT:
        r = requests.get(
            f"{PAPERLESS_URL}/api/tasks/?task_id={task_id}",
            headers=headers,
        )
        r.raise_for_status()
        tasks = r.json()

        if not tasks:
            time.sleep(TASK_INTERVAL)
            elapsed += TASK_INTERVAL
            continue

        task = tasks[0]
        status = task.get("status", "")

        if status == "SUCCESS":
            return True, None
        elif status == "FAILURE":
            result = task.get("result", "Unbekannter Fehler")
            return False, result
        else:
            log.info(f"  Task {task_id} Status: {status}, warte...")
            time.sleep(TASK_INTERVAL)
            elapsed += TASK_INTERVAL

    return False, f"Timeout nach {TASK_TIMEOUT}s"


def _log_import_result(eintrag):
    """Schreibt einen Import-Eintrag als JSON-Zeile in die Log-Datei (append)."""
    try:
        with open(IMPORT_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(json.dumps(eintrag, ensure_ascii=False) + "\n")
    except Exception as e:
        log.error(f"Fehler beim Schreiben in die Import-Log-Datei: {e}")


def _read_log_entries(datum_iso):
    """
    Liest alle Import-Einträge für ein Datum (YYYY-MM-DD) aus der Log-Datei.
    Gibt zwei Listen zurück: (erfolgreich, fehlerhaft).
    """
    erfolgreich = []
    fehlerhaft = []
    if not os.path.exists(IMPORT_LOG_FILE):
        return erfolgreich, fehlerhaft
    with open(IMPORT_LOG_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                entry = json.loads(line)
            except json.JSONDecodeError:
                continue
            if entry.get("typ") != "import":
                continue
            if not entry.get("ts", "").startswith(datum_iso):
                continue
            try:
                zeitpunkt = datetime.datetime.fromisoformat(entry["ts"]).strftime("%d.%m.%Y %H:%M")
            except (ValueError, KeyError):
                zeitpunkt = entry.get("ts", "")
            if entry.get("status") == "erfolgreich":
                erfolgreich.append({
                    "datei": entry.get("datei", ""),
                    "betreff": entry.get("betreff", ""),
                    "zeitpunkt": zeitpunkt,
                })
            elif entry.get("status") == "fehlerhaft":
                fehlerhaft.append({
                    "datei": entry.get("datei", ""),
                    "betreff": entry.get("betreff", ""),
                    "fehler": entry.get("fehler", ""),
                    "zeitpunkt": zeitpunkt,
                })
    return erfolgreich, fehlerhaft


def _get_last_summary_date():
    """
    Liest das Datum der zuletzt versendeten Zusammenfassung aus der Log-Datei.
    Gibt ein datetime.date-Objekt zurück, oder None wenn noch keine gesendet wurde.
    """
    if not os.path.exists(IMPORT_LOG_FILE):
        return None
    letztes = None
    with open(IMPORT_LOG_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                entry = json.loads(line)
            except json.JSONDecodeError:
                continue
            if entry.get("typ") == "summary_gesendet":
                try:
                    letztes = datetime.date.fromisoformat(entry["datum"])
                except (ValueError, KeyError):
                    pass
    return letztes


def graph_send_mail(token, to_address, subject, html_body):
    """Sendet eine E-Mail über die Graph API (sendMail). Gibt keinen Body zurück."""
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": to_address}}],
        },
        "saveToSentItems": False,
    }
    r = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/sendMail",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=payload,
    )
    r.raise_for_status()


def _build_summary_html(erfolgreich, fehlerhaft, datum):
    """Baut den HTML-Body der täglichen Zusammenfassung."""
    def zeilen(eintraege, felder):
        if not eintraege:
            return "<tr><td colspan='{}' style='color:#888'>keine</td></tr>".format(len(felder))
        rows = ""
        for e in eintraege:
            rows += "<tr>" + "".join(f"<td style='padding:4px 8px;border:1px solid #ddd'>{e.get(f,'')}</td>" for f in felder) + "</tr>"
        return rows

    html = f"""
    <html><body style='font-family:sans-serif;color:#333'>
    <h2>Paperless Import – Zusammenfassung {datum}</h2>
    <p>
        <b style='color:green'>Erfolgreich importiert:</b> {len(erfolgreich)} Datei(en)<br>
        <b style='color:red'>Fehlgeschlagen:</b> {len(fehlerhaft)} Datei(en)
    </p>

    <h3 style='color:green'>Erfolgreich importierte Dateien</h3>
    <table style='border-collapse:collapse;width:100%'>
        <thead><tr style='background:#e8f5e9'>
            <th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Datei</th>
            <th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Betreff der Mail</th>
            <th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Zeitpunkt</th>
        </tr></thead>
        <tbody>{zeilen(erfolgreich, ['datei','betreff','zeitpunkt'])}</tbody>
    </table>

    <h3 style='color:red;margin-top:24px'>Fehlgeschlagene Importe</h3>
    <table style='border-collapse:collapse;width:100%'>
        <thead><tr style='background:#ffebee'>
            <th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Datei</th>
            <th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Betreff der Mail</th>
            <th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Fehler</th>
            <th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Zeitpunkt</th>
        </tr></thead>
        <tbody>{zeilen(fehlerhaft, ['datei','betreff','fehler','zeitpunkt'])}</tbody>
    </table>
    </body></html>
    """
    return html


def send_daily_summary(token):
    """
    Liest die Tageseinträge aus der Log-Datei und versendet die Zusammenfassungs-Mail.
    Schreibt anschließend einen summary_gesendet-Sentinel ins Log.
    """
    empfaenger = SUMMARY_RECIPIENT or USER_EMAIL
    heute = datetime.date.today()
    datum_iso = heute.isoformat()
    datum_anzeige = heute.strftime("%d.%m.%Y")
    betreff = f"Paperless Import Zusammenfassung – {datum_anzeige}"
    erfolgreich, fehlerhaft = _read_log_entries(datum_iso)
    html = _build_summary_html(erfolgreich, fehlerhaft, datum_anzeige)
    graph_send_mail(token, empfaenger, betreff, html)
    # Sentinel-Eintrag damit Neustarts den heutigen Versand erkennen
    _log_import_result({"typ": "summary_gesendet", "datum": datum_iso})
    log.info(
        f"Tägliche Zusammenfassung gesendet an {empfaenger}: "
        f"{len(erfolgreich)} erfolgreich, {len(fehlerhaft)} fehlerhaft."
    )


def mark_as_read(token, message_id):
    graph_patch(token, f"users/{USER_EMAIL}/messages/{message_id}", {"isRead": True})


def move_message(token, message_id, destination_folder_id):
    graph_post(
        token,
        f"users/{USER_EMAIL}/messages/{message_id}/move",
        {"destinationId": destination_folder_id},
    )


def process_messages(token, source_folder_id, done_folder_id, error_folder_id):
    """Verarbeitet alle Nachrichten im Quell-Ordner und schreibt Ergebnisse ins Logfile."""
    messages = get_messages(token, source_folder_id)
    if not messages:
        log.info("Keine Nachrichten im Ordner gefunden.")
        return

    for msg in messages:
        subject = msg.get("subject", "(kein Betreff)")
        msg_id = msg["id"]
        attachments = msg.get("attachments", [])

        uploaded = 0
        failed = 0

        for att in attachments:
            content_type = att.get("contentType", "")
            if content_type not in SUPPORTED_TYPES:
                log.info(f"  Überspringe Anhang '{att.get('name')}' (Typ: {content_type})")
                continue

            filename = att.get("name", "dokument.pdf")
            file_bytes = base64.b64decode(att["contentBytes"])

            try:
                task_id = upload_to_paperless(filename, file_bytes, content_type)
                log.info(f"  Hochgeladen: '{filename}' -> Task {task_id}")

                success, error = wait_for_task(task_id)
                if success:
                    log.info(f"  Task erfolgreich abgeschlossen: '{filename}'")
                    uploaded += 1
                    _log_import_result({
                        "typ": "import",
                        "ts": datetime.datetime.now().isoformat(timespec="seconds"),
                        "datei": filename,
                        "betreff": subject,
                        "status": "erfolgreich",
                        "fehler": None,
                    })
                else:
                    log.warning(f"  Task fehlgeschlagen für '{filename}': {error}")
                    failed += 1
                    _log_import_result({
                        "typ": "import",
                        "ts": datetime.datetime.now().isoformat(timespec="seconds"),
                        "datei": filename,
                        "betreff": subject,
                        "status": "fehlerhaft",
                        "fehler": error or "Unbekannt",
                    })
            except Exception as e:
                log.error(f"  Fehler beim Upload von '{filename}': {e}")
                failed += 1
                _log_import_result({
                    "typ": "import",
                    "ts": datetime.datetime.now().isoformat(timespec="seconds"),
                    "datei": filename,
                    "betreff": subject,
                    "status": "fehlerhaft",
                    "fehler": str(e),
                })

        # Verschieben je nach Ergebnis
        if failed > 0:
            mark_as_read(token, msg_id)
            move_message(token, msg_id, error_folder_id)
            log.warning(f"Mail '{subject}' wegen Fehler nach 'fehlerhaft' verschoben.")
        elif uploaded > 0:
            mark_as_read(token, msg_id)
            move_message(token, msg_id, done_folder_id)
            log.info(f"Mail '{subject}' erfolgreich verarbeitet und verschoben.")
        else:
            log.info(f"Mail '{subject}' hatte keine unterstützten Anhänge, wird übersprungen.")


def run():
    log.info("Starte Paperless Mail Consumer...")

    # Letztes Summary-Datum aus Log-Datei lesen – überlebt Neustarts
    letztes_summary_datum = _get_last_summary_date()
    if letztes_summary_datum:
        log.info(f"Letzte Zusammenfassung wurde am {letztes_summary_datum} versendet.")

    while True:
        try:
            token = get_token()

            source_folder_id = get_folder_id(token, MAIL_FOLDER)

            done_folder_id = get_or_create_subfolder(
                token, source_folder_id, "verarbeitet"
            )
            error_folder_id = get_or_create_subfolder(
                token, source_folder_id, "fehlerhaft"
            )

            process_messages(token, source_folder_id, done_folder_id, error_folder_id)

            # Tägliche Zusammenfassung einmal um SUMMARY_HOUR Uhr versenden
            jetzt = datetime.datetime.now()
            if jetzt.hour >= SUMMARY_HOUR and jetzt.date() != letztes_summary_datum:
                try:
                    send_daily_summary(token)
                    letztes_summary_datum = jetzt.date()
                except Exception as e:
                    log.error(f"Fehler beim Versenden der Zusammenfassung: {e}")

        except Exception as e:
            log.error(f"Fehler im Hauptloop: {e}")

        log.info(f"Warte {POLL_INTERVAL} Sekunden...")
        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    run()
