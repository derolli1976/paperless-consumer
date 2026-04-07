import os
import json
import time
import base64
import logging
import datetime
import subprocess
import tempfile
from zoneinfo import ZoneInfo
import requests
import msal
from dotenv import load_dotenv

# Load .env file if present (no-op when running in Docker with env vars injected)
_env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
load_dotenv(_env_path)

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
IMPORT_LOG_FILE    = os.getenv("IMPORT_LOG_FILE", "/app/data/import.log")

SUPPORTED_TYPES = [
    "application/pdf",
    "image/jpeg",
    "image/png",
    "image/tiff",
]

# Word MIME types that will be converted to PDF before upload
WORD_TYPES = [
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/msword",
]

# Fallback mapping: file extension -> MIME type for attachments with generic contentType
EXTENSION_TYPE_MAP = {
    ".pdf": "application/pdf",
    ".jpeg": "image/jpeg",
    ".jpg": "image/jpeg",
    ".png": "image/png",
    ".tiff": "image/tiff",
    ".tif": "image/tiff",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".doc": "application/msword",
}

# All timestamps in German timezone (Europe/Berlin)
BERLIN = ZoneInfo("Europe/Berlin")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logging.Formatter.converter = lambda *args: datetime.datetime.now(BERLIN).timetuple()
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
    if not result or "access_token" not in result:
        raise Exception(f"Token error: {result.get('error_description') if result else 'No result'}")
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
    # Check top-level folders first
    data = graph_get(token, f"users/{USER_EMAIL}/mailFolders?$top=100")
    for folder in data.get("value", []):
        if folder["displayName"] == folder_name:
            return folder["id"]

    # Search child folders of all top-level folders
    for parent in data.get("value", []):
        children = graph_get(
            token,
            f"users/{USER_EMAIL}/mailFolders/{parent['id']}/childFolders?$top=100"
        )
        for folder in children.get("value", []):
            if folder["displayName"] == folder_name:
                return folder["id"]

    raise Exception(f"Folder '{folder_name}' not found")


def get_or_create_subfolder(token, parent_id, subfolder_name):
    data = graph_get(
        token,
        f"users/{USER_EMAIL}/mailFolders/{parent_id}/childFolders?$top=100"
    )
    for folder in data.get("value", []):
        if folder["displayName"] == subfolder_name:
            return folder["id"]
    log.info(f"Creating subfolder '{subfolder_name}'...")
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


def convert_word_to_pdf(filename, file_bytes):
    """
    Converts a Word document (.doc/.docx) to PDF using LibreOffice headless.
    Returns (pdf_filename, pdf_bytes) or raises an exception on failure.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, filename)
        with open(input_path, "wb") as f:
            f.write(file_bytes)

        result = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, input_path],
            capture_output=True, text=True, timeout=120,
        )
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice conversion failed: {result.stderr.strip()}")

        pdf_name = os.path.splitext(filename)[0] + ".pdf"
        pdf_path = os.path.join(tmpdir, pdf_name)
        if not os.path.exists(pdf_path):
            raise RuntimeError(f"PDF output not found after conversion: {pdf_name}")

        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()

    return pdf_name, pdf_bytes


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
    # Paperless returns the task UUID as a string (possibly with surrounding quotes)
    task_id = r.text.strip().strip('"')
    return task_id


def wait_for_task(task_id):
    """
    Polls a Paperless task until it completes or times out.
    Returns (True, None) on success, or (False, "error message") on failure.
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
            result = task.get("result", "Unknown error")
            return False, result
        else:
            log.info(f"  Task {task_id} status: {status}, waiting...")
            time.sleep(TASK_INTERVAL)
            elapsed += TASK_INTERVAL

    return False, f"Timeout after {TASK_TIMEOUT}s"


def _write_log_entry(entry):
    """Appends a log entry as a JSON line to the import log file."""
    try:
        with open(IMPORT_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except Exception as e:
        log.error(f"Error writing to import log file: {e}")


def _read_log_entries(date_iso):
    """
    Reads all import entries for a given date (YYYY-MM-DD) from the log file.
    Returns two lists: (succeeded, failed).
    """
    succeeded = []
    failed = []
    if not os.path.exists(IMPORT_LOG_FILE):
        return succeeded, failed
    with open(IMPORT_LOG_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                entry = json.loads(line)
            except json.JSONDecodeError:
                continue
            if entry.get("type") != "import":
                continue
            if not entry.get("ts", "").startswith(date_iso):
                continue
            try:
                timestamp = datetime.datetime.fromisoformat(entry["ts"]).strftime("%d.%m.%Y %H:%M")
            except (ValueError, KeyError):
                timestamp = entry.get("ts", "")
            if entry.get("status") == "success":
                succeeded.append({
                    "file": entry.get("file", ""),
                    "subject": entry.get("subject", ""),
                    "timestamp": timestamp,
                })
            elif entry.get("status") == "failed":
                failed.append({
                    "file": entry.get("file", ""),
                    "subject": entry.get("subject", ""),
                    "error": entry.get("error", ""),
                    "timestamp": timestamp,
                })
    return succeeded, failed


def _read_log_entries_since_last_summary():
    """
    Reads all import entries that appear after the last summary_sent entry in
    the log file. Returns two lists: (succeeded, failed).
    This ensures that imports processed after the previous day's summary are
    included in the next summary.
    """
    succeeded = []
    failed = []
    if not os.path.exists(IMPORT_LOG_FILE):
        return succeeded, failed

    # Load all entries
    all_entries = []
    with open(IMPORT_LOG_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                all_entries.append(json.loads(line))
            except json.JSONDecodeError:
                continue

    # Find the index of the last summary_sent entry
    last_summary_idx = -1
    for i, entry in enumerate(all_entries):
        if entry.get("type") == "summary_sent":
            last_summary_idx = i

    # Process only entries after the last summary
    for entry in all_entries[last_summary_idx + 1:]:
        if entry.get("type") != "import":
            continue
        try:
            timestamp = datetime.datetime.fromisoformat(entry["ts"]).strftime("%d.%m.%Y %H:%M")
        except (ValueError, KeyError):
            timestamp = entry.get("ts", "")
        if entry.get("status") == "success":
            succeeded.append({
                "file": entry.get("file", ""),
                "subject": entry.get("subject", ""),
                "timestamp": timestamp,
            })
        elif entry.get("status") == "failed":
            failed.append({
                "file": entry.get("file", ""),
                "subject": entry.get("subject", ""),
                "error": entry.get("error", ""),
                "timestamp": timestamp,
            })
    return succeeded, failed


def _get_last_summary_date():
    """
    Reads the date of the last sent summary from the log file.
    Returns a datetime.date object, or None if no summary has been sent yet.
    """
    if not os.path.exists(IMPORT_LOG_FILE):
        return None
    last = None
    with open(IMPORT_LOG_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                entry = json.loads(line)
            except json.JSONDecodeError:
                continue
            if entry.get("type") == "summary_sent":
                try:
                    last = datetime.date.fromisoformat(entry["date"])
                except (ValueError, KeyError):
                    pass
    return last


def graph_send_mail(token, to_address, subject, html_body):
    """Sends an email via the Graph API (sendMail). Returns no body."""
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


def _analyze_pending_messages(token, folder_id):
    """Analyzes messages still in the source folder and returns a list of
    dicts with subject, sender, and reason why they were not processed."""
    messages = get_messages(token, folder_id)
    pending = []
    for msg in messages:
        subject = msg.get("subject", "(kein Betreff)")
        sender_data = msg.get("from") or {}
        sender = sender_data.get("emailAddress", {}).get("address", "unbekannt")
        attachments = msg.get("attachments", [])

        if not attachments:
            reason = "Kein Anhang gefunden"
        else:
            unsupported = []
            for att in attachments:
                content_type = att.get("contentType", "")
                filename = att.get("name", "unbekannt")
                ext = os.path.splitext(filename)[1].lower()
                resolved = EXTENSION_TYPE_MAP.get(ext)
                if content_type not in SUPPORTED_TYPES and content_type not in WORD_TYPES and not resolved:
                    unsupported.append(f"{filename} ({content_type})")

            if unsupported and len(unsupported) == len(attachments):
                reason = f"Keine unterstützten Anhänge: {', '.join(unsupported)}"
            elif unsupported:
                reason = f"Teilweise nicht unterstützte Anhänge: {', '.join(unsupported)}"
            else:
                reason = "Unbekannter Grund"

        pending.append({
            "subject": subject,
            "sender": sender,
            "reason": reason,
        })
    return pending


def _build_summary_html(succeeded, failed, date_display, pending=None):
    """Builds the HTML body of the daily summary email (content in German)."""
    if pending is None:
        pending = []

    def rows(entries, fields):
        if not entries:
            return "<tr><td colspan='{}' style='color:#888'>keine</td></tr>".format(len(fields))
        result = ""
        for e in entries:
            result += "<tr>" + "".join(f"<td style='padding:4px 8px;border:1px solid #ddd'>{e.get(f,'')}</td>" for f in fields) + "</tr>"
        return result

    pending_section = ""
    if pending:
        pending_section = f"""
        <h3 style='color:#e65100;margin-top:24px'>Nicht verarbeitete Mails im Ordner ({len(pending)})</h3>
        <table style='border-collapse:collapse;width:100%'>
            <thead><tr style='background:#fff3e0'>
                <th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Betreff</th>
                <th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Absender</th>
                <th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Grund</th>
            </tr></thead>
            <tbody>{rows(pending, ['subject','sender','reason'])}</tbody>
        </table>
        """

    html = f"""
    <html><body style='font-family:sans-serif;color:#333'>
    <h2>Paperless Import – Zusammenfassung {date_display}</h2>
    <p>
        <b style='color:green'>Erfolgreich importiert:</b> {len(succeeded)} Datei(en)<br>
        <b style='color:red'>Fehlgeschlagen:</b> {len(failed)} Datei(en)
    </p>

    <h3 style='color:green'>Erfolgreich importierte Dateien</h3>
    <table style='border-collapse:collapse;width:100%'>
        <thead><tr style='background:#e8f5e9'>
            <th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Datei</th>
            <th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Betreff der Mail</th>
            <th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Zeitpunkt</th>
        </tr></thead>
        <tbody>{rows(succeeded, ['file','subject','timestamp'])}</tbody>
    </table>

    <h3 style='color:red;margin-top:24px'>Fehlgeschlagene Importe</h3>
    <table style='border-collapse:collapse;width:100%'>
        <thead><tr style='background:#ffebee'>
            <th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Datei</th>
            <th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Betreff der Mail</th>
            <th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Fehler</th>
            <th style='padding:4px 8px;border:1px solid #ddd;text-align:left'>Zeitpunkt</th>
        </tr></thead>
        <tbody>{rows(failed, ['file','subject','error','timestamp'])}</tbody>
    </table>
    {pending_section}
    </body></html>
    """
    return html


def send_daily_summary(token, folder_id=None):
    """
    Reads all import entries since the last summary from the log file and sends
    the summary email. When folder_id is provided, also lists messages still
    sitting in the source folder together with the reason they were not processed.
    Writes a summary_sent sentinel entry to the log afterwards.
    """
    recipient = SUMMARY_RECIPIENT or USER_EMAIL
    today = datetime.datetime.now(BERLIN).date()
    date_display = today.strftime("%d.%m.%Y")
    subject_line = f"Paperless Import Zusammenfassung – {date_display}"
    # Collect all imports since the last summary (not just today's)
    succeeded, failed = _read_log_entries_since_last_summary()
    # Analyze messages still in the source folder
    pending = []
    if folder_id:
        try:
            pending = _analyze_pending_messages(token, folder_id)
        except Exception as e:
            log.error(f"Error analyzing pending messages: {e}")
    html = _build_summary_html(succeeded, failed, date_display, pending)
    graph_send_mail(token, recipient, subject_line, html)
    # Write sentinel so service restarts detect today's send
    _write_log_entry({"type": "summary_sent", "date": today.isoformat()})
    log.info(
        f"Daily summary sent to {recipient}: "
        f"{len(succeeded)} succeeded, {len(failed)} failed, "
        f"{len(pending)} pending in folder."
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
    """Processes all messages in the source folder and writes results to the log file."""
    messages = get_messages(token, source_folder_id)
    if not messages:
        log.info("No messages found in folder.")
        return

    for msg in messages:
        subject = msg.get("subject", "(no subject)")
        msg_id = msg["id"]
        attachments = msg.get("attachments", [])

        uploaded = 0
        failed = 0

        for att in attachments:
            content_type = att.get("contentType", "")
            filename = att.get("name", "document.pdf")

            # Fallback: derive MIME type from file extension for generic contentType
            if content_type not in SUPPORTED_TYPES and content_type not in WORD_TYPES:
                ext = os.path.splitext(filename)[1].lower()
                resolved = EXTENSION_TYPE_MAP.get(ext)
                if resolved:
                    log.info(f"  Resolved '{filename}' from '{content_type}' to '{resolved}' by extension")
                    content_type = resolved
                else:
                    log.info(f"  Skipping attachment '{filename}' (type: {content_type})")
                    continue

            file_bytes = base64.b64decode(att["contentBytes"])

            # Convert Word documents to PDF before upload
            if content_type in WORD_TYPES:
                try:
                    log.info(f"  Converting '{filename}' to PDF...")
                    filename, file_bytes = convert_word_to_pdf(filename, file_bytes)
                    content_type = "application/pdf"
                    log.info(f"  Converted to '{filename}'")
                except Exception as e:
                    log.error(f"  Word-to-PDF conversion failed for '{filename}': {e}")
                    failed += 1
                    _write_log_entry({
                        "type": "import",
                        "ts": datetime.datetime.now(BERLIN).isoformat(timespec="seconds"),
                        "file": filename,
                        "subject": subject,
                        "status": "failed",
                        "error": f"Word-to-PDF conversion failed: {e}",
                    })
                    continue

            try:
                task_id = upload_to_paperless(filename, file_bytes, content_type)
                log.info(f"  Uploaded: '{filename}' -> task {task_id}")

                success, error = wait_for_task(task_id)
                if success:
                    log.info(f"  Task completed successfully: '{filename}'")
                    uploaded += 1
                    _write_log_entry({
                        "type": "import",
                        "ts": datetime.datetime.now(BERLIN).isoformat(timespec="seconds"),
                        "file": filename,
                        "subject": subject,
                        "status": "success",
                        "error": None,
                    })
                else:
                    log.warning(f"  Task failed for '{filename}': {error}")
                    failed += 1
                    _write_log_entry({
                        "type": "import",
                        "ts": datetime.datetime.now(BERLIN).isoformat(timespec="seconds"),
                        "file": filename,
                        "subject": subject,
                        "status": "failed",
                        "error": error or "Unknown",
                    })
            except Exception as e:
                log.error(f"  Error uploading '{filename}': {e}")
                failed += 1
                _write_log_entry({
                    "type": "import",
                    "ts": datetime.datetime.now(BERLIN).isoformat(timespec="seconds"),
                    "file": filename,
                    "subject": subject,
                    "status": "failed",
                    "error": str(e),
                })

        # Move message based on outcome
        if failed > 0:
            mark_as_read(token, msg_id)
            move_message(token, msg_id, error_folder_id)
            log.warning(f"Message '{subject}' moved to error folder due to failures.")
        elif uploaded > 0:
            mark_as_read(token, msg_id)
            move_message(token, msg_id, done_folder_id)
            log.info(f"Message '{subject}' processed successfully and moved.")
        else:
            log.info(f"Message '{subject}' had no supported attachments, skipping.")


def run():
    log.info("Starting Paperless Mail Consumer...")

    # Read last summary date from log file – survives service restarts
    last_summary_date = _get_last_summary_date()
    if last_summary_date:
        log.info(f"Last summary was sent on {last_summary_date}.")

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

            # Send daily summary once at SUMMARY_HOUR (Berlin time)
            now = datetime.datetime.now(BERLIN)
            if now.hour >= SUMMARY_HOUR and now.date() != last_summary_date:
                try:
                    send_daily_summary(token, source_folder_id)
                    last_summary_date = now.date()
                except Exception as e:
                    log.error(f"Error sending daily summary: {e}")

        except Exception as e:
            log.error(f"Error in main loop: {e}")

        log.info(f"Waiting {POLL_INTERVAL} seconds...")
        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    run()
