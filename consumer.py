import os
import time
import base64
import logging
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
POLL_INTERVAL   = int(os.getenv("POLL_INTERVAL", "300"))
TASK_TIMEOUT    = int(os.getenv("TASK_TIMEOUT", "120"))
TASK_INTERVAL   = int(os.getenv("TASK_INTERVAL", "3"))

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


def mark_as_read(token, message_id):
    graph_patch(token, f"users/{USER_EMAIL}/messages/{message_id}", {"isRead": True})


def move_message(token, message_id, destination_folder_id):
    graph_post(
        token,
        f"users/{USER_EMAIL}/messages/{message_id}/move",
        {"destinationId": destination_folder_id},
    )


def process_messages(token, source_folder_id, done_folder_id, error_folder_id):
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
                else:
                    log.warning(f"  Task fehlgeschlagen für '{filename}': {error}")
                    failed += 1
            except Exception as e:
                log.error(f"  Fehler beim Upload von '{filename}': {e}")
                failed += 1

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

        except Exception as e:
            log.error(f"Fehler im Hauptloop: {e}")

        log.info(f"Warte {POLL_INTERVAL} Sekunden...")
        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    run()
