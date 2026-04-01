"""
Tests für consumer.py

Alle externen API-Aufrufe (Microsoft Graph, Paperless, MSAL) werden mit
unittest.mock gemockt. Die Tests decken Erfolgs- und Fehlerpfade ab.
"""

import base64
import datetime
import json
import os

import pytest

# Umgebungsvariablen vor dem Import setzen, damit die Modul-Konstanten befüllt werden
os.environ.setdefault("AZURE_CLIENT_ID", "test-client-id")
os.environ.setdefault("AZURE_CLIENT_SECRET", "test-client-secret")
os.environ.setdefault("AZURE_TENANT_ID", "test-tenant-id")
os.environ.setdefault("USER_EMAIL", "test@example.com")
os.environ.setdefault("PAPERLESS_URL", "http://localhost:8000")
os.environ.setdefault("PAPERLESS_TOKEN", "test-token")
os.environ.setdefault("MAIL_FOLDER", "_ecoDMS")
os.environ.setdefault("INBOX_TAG_ID", "2")

from unittest.mock import MagicMock, call, patch

import consumer


# ---------------------------------------------------------------------------
# get_token
# ---------------------------------------------------------------------------


class TestGetToken:
    def test_erfolgreicher_token_abruf(self):
        """Gibt den Access-Token zurück wenn MSAL erfolgreich antwortet."""
        mock_app = MagicMock()
        mock_app.acquire_token_for_client.return_value = {"access_token": "mein-token"}

        with patch("msal.ConfidentialClientApplication", return_value=mock_app):
            token = consumer.get_token()

        assert token == "mein-token"

    def test_token_fehler_wirft_exception(self):
        """Wirft Exception wenn kein access_token im MSAL-Ergebnis enthalten ist."""
        mock_app = MagicMock()
        mock_app.acquire_token_for_client.return_value = {
            "error": "invalid_client",
            "error_description": "Client authentication failed",
        }

        with patch("msal.ConfidentialClientApplication", return_value=mock_app):
            with pytest.raises(Exception, match="Token-Fehler"):
                consumer.get_token()


# ---------------------------------------------------------------------------
# graph_get / graph_patch / graph_post
# ---------------------------------------------------------------------------


class TestGraphHelpers:
    def test_graph_get_gibt_json_zurueck(self):
        """graph_get ruft die korrekte URL auf und gibt JSON zurück."""
        mock_response = MagicMock()
        mock_response.json.return_value = {"value": [{"id": "123"}]}

        with patch("requests.get", return_value=mock_response) as mock_get:
            result = consumer.graph_get("token123", "users/test@example.com/mailFolders")

        mock_get.assert_called_once()
        assert "Bearer token123" in mock_get.call_args.kwargs["headers"]["Authorization"]
        assert result == {"value": [{"id": "123"}]}

    def test_graph_get_raise_bei_http_fehler(self):
        """graph_get wirft bei HTTP-Fehler eine Exception."""
        mock_response = MagicMock()
        mock_response.raise_for_status.side_effect = Exception("HTTP 403")

        with patch("requests.get", return_value=mock_response):
            with pytest.raises(Exception, match="HTTP 403"):
                consumer.graph_get("token", "some/path")

    def test_graph_patch_sendet_payload(self):
        """graph_patch sendet den korrekten JSON-Payload."""
        mock_response = MagicMock()

        with patch("requests.patch", return_value=mock_response) as mock_patch:
            consumer.graph_patch("token123", "users/x/messages/id1", {"isRead": True})

        mock_patch.assert_called_once()
        assert mock_patch.call_args.kwargs["json"] == {"isRead": True}

    def test_graph_patch_raise_bei_http_fehler(self):
        """graph_patch wirft bei HTTP-Fehler eine Exception."""
        mock_response = MagicMock()
        mock_response.raise_for_status.side_effect = Exception("HTTP 500")

        with patch("requests.patch", return_value=mock_response):
            with pytest.raises(Exception, match="HTTP 500"):
                consumer.graph_patch("token", "path", {})

    def test_graph_post_gibt_json_zurueck(self):
        """graph_post sendet Payload und gibt JSON-Antwort zurück."""
        mock_response = MagicMock()
        mock_response.json.return_value = {"id": "neuer-ordner-id"}

        with patch("requests.post", return_value=mock_response) as mock_post:
            result = consumer.graph_post("token", "some/path", {"displayName": "Test"})

        assert result == {"id": "neuer-ordner-id"}
        assert mock_post.call_args.kwargs["json"] == {"displayName": "Test"}

    def test_graph_post_raise_bei_http_fehler(self):
        """graph_post wirft bei HTTP-Fehler eine Exception."""
        mock_response = MagicMock()
        mock_response.raise_for_status.side_effect = Exception("HTTP 404")

        with patch("requests.post", return_value=mock_response):
            with pytest.raises(Exception, match="HTTP 404"):
                consumer.graph_post("token", "path", {})


# ---------------------------------------------------------------------------
# get_folder_id
# ---------------------------------------------------------------------------


class TestGetFolderId:
    def test_findet_top_level_ordner(self):
        """Findet einen Ordner direkt auf der obersten Ebene."""
        top_level = {"value": [{"id": "folder-1", "displayName": "_ecoDMS"}]}

        with patch("consumer.graph_get", return_value=top_level):
            folder_id = consumer.get_folder_id("token", "_ecoDMS")

        assert folder_id == "folder-1"

    def test_findet_unterordner(self):
        """Findet einen Ordner als Unterordner eines Top-Level-Ordners."""
        top_level = {"value": [{"id": "inbox-id", "displayName": "Posteingang"}]}
        child_folders = {"value": [{"id": "eco-id", "displayName": "_ecoDMS"}]}

        def graph_get_side_effect(token, path):
            if "childFolders" in path:
                return child_folders
            return top_level

        with patch("consumer.graph_get", side_effect=graph_get_side_effect):
            folder_id = consumer.get_folder_id("token", "_ecoDMS")

        assert folder_id == "eco-id"

    def test_wirft_exception_wenn_nicht_gefunden(self):
        """Wirft Exception wenn der Ordner weder top-level noch als Unterordner gefunden wird."""
        empty = {"value": []}

        with patch("consumer.graph_get", return_value=empty):
            with pytest.raises(Exception, match="nicht gefunden"):
                consumer.get_folder_id("token", "ExistiertNicht")

    def test_top_level_hat_vorrang_vor_unterordner(self):
        """Wenn der Ordner auf beiden Ebenen existiert, wird der Top-Level-Treffer zurückgegeben."""
        top_level = {
            "value": [
                {"id": "top-id", "displayName": "_ecoDMS"},
                {"id": "inbox-id", "displayName": "Posteingang"},
            ]
        }
        child_folders = {"value": [{"id": "child-id", "displayName": "_ecoDMS"}]}

        def graph_get_side_effect(token, path):
            if "childFolders" in path:
                return child_folders
            return top_level

        with patch("consumer.graph_get", side_effect=graph_get_side_effect):
            folder_id = consumer.get_folder_id("token", "_ecoDMS")

        # Top-Level-Ordner hat Vorrang
        assert folder_id == "top-id"


# ---------------------------------------------------------------------------
# get_or_create_subfolder
# ---------------------------------------------------------------------------


class TestGetOrCreateSubfolder:
    def test_gibt_vorhandenen_unterordner_zurueck(self):
        """Gibt die ID eines bereits vorhandenen Unterordners zurück ohne neu anzulegen."""
        existing = {"value": [{"id": "sub-id", "displayName": "verarbeitet"}]}

        with patch("consumer.graph_get", return_value=existing):
            with patch("consumer.graph_post") as mock_post:
                result = consumer.get_or_create_subfolder("token", "parent-id", "verarbeitet")

        assert result == "sub-id"
        mock_post.assert_not_called()

    def test_legt_neuen_unterordner_an(self):
        """Erstellt einen neuen Unterordner wenn er noch nicht existiert."""
        no_children = {"value": []}
        new_folder = {"id": "new-sub-id", "displayName": "fehlerhaft"}

        with patch("consumer.graph_get", return_value=no_children):
            with patch("consumer.graph_post", return_value=new_folder) as mock_post:
                result = consumer.get_or_create_subfolder("token", "parent-id", "fehlerhaft")

        assert result == "new-sub-id"
        mock_post.assert_called_once()
        assert mock_post.call_args.args[2] == {"displayName": "fehlerhaft"}


# ---------------------------------------------------------------------------
# upload_to_paperless
# ---------------------------------------------------------------------------


class TestUploadToPaperless:
    def test_gibt_task_id_zurueck(self):
        """Gibt die Task-UUID zurück – umgebende Anführungszeichen werden entfernt."""
        mock_response = MagicMock()
        mock_response.text = '"abc-123-uuid"'

        with patch("requests.post", return_value=mock_response):
            task_id = consumer.upload_to_paperless("test.pdf", b"pdf-inhalt", "application/pdf")

        assert task_id == "abc-123-uuid"

    def test_sendet_tag_und_datei(self):
        """Übergibt INBOX_TAG_ID und Dateiinhalt korrekt an die Paperless API."""
        mock_response = MagicMock()
        mock_response.text = '"task-uuid"'

        with patch("requests.post", return_value=mock_response) as mock_post:
            consumer.upload_to_paperless("dokument.pdf", b"bytes", "application/pdf")

        call_kwargs = mock_post.call_args
        assert "document" in call_kwargs.kwargs["files"]
        assert consumer.INBOX_TAG_ID in call_kwargs.kwargs["data"]["tags"]

    def test_task_id_ohne_anfuehrungszeichen(self):
        """Funktioniert auch wenn Paperless die UUID ohne Anführungszeichen zurückgibt."""
        mock_response = MagicMock()
        mock_response.text = "abc-456-uuid"

        with patch("requests.post", return_value=mock_response):
            task_id = consumer.upload_to_paperless("test.pdf", b"inhalt", "application/pdf")

        assert task_id == "abc-456-uuid"

    def test_raise_bei_http_fehler(self):
        """Wirft Exception bei HTTP-Fehler vom Paperless-Server."""
        mock_response = MagicMock()
        mock_response.raise_for_status.side_effect = Exception("HTTP 500")

        with patch("requests.post", return_value=mock_response):
            with pytest.raises(Exception, match="HTTP 500"):
                consumer.upload_to_paperless("test.pdf", b"inhalt", "application/pdf")


# ---------------------------------------------------------------------------
# wait_for_task
# ---------------------------------------------------------------------------


class TestWaitForTask:
    def test_erfolg_beim_ersten_versuch(self):
        """Gibt (True, None) zurück wenn der Task sofort SUCCESS meldet."""
        mock_response = MagicMock()
        mock_response.json.return_value = [{"status": "SUCCESS", "result": "ok"}]

        with patch("requests.get", return_value=mock_response):
            with patch("time.sleep"):
                success, error = consumer.wait_for_task("task-123")

        assert success is True
        assert error is None

    def test_failure_bei_duplikat(self):
        """Gibt (False, Fehlermeldung) zurück wenn der Task FAILURE meldet (z.B. Duplikat)."""
        mock_response = MagicMock()
        mock_response.json.return_value = [
            {"status": "FAILURE", "result": "duplicate document found"}
        ]

        with patch("requests.get", return_value=mock_response):
            with patch("time.sleep"):
                success, error = consumer.wait_for_task("task-456")

        assert success is False
        assert "duplicate" in error

    def test_wartet_bei_pending_dann_success(self):
        """Wartet bei PENDING-Status und gibt (True, None) zurück wenn danach SUCCESS folgt."""
        pending_response = MagicMock()
        pending_response.json.return_value = [{"status": "PENDING"}]

        success_response = MagicMock()
        success_response.json.return_value = [{"status": "SUCCESS"}]

        with patch("requests.get", side_effect=[pending_response, success_response]):
            with patch("time.sleep"):
                success, error = consumer.wait_for_task("task-789")

        assert success is True
        assert error is None

    def test_wartet_bei_started_dann_failure(self):
        """Wartet bei STARTED-Status und gibt (False, ...) zurück wenn FAILURE folgt."""
        started_response = MagicMock()
        started_response.json.return_value = [{"status": "STARTED"}]

        failure_response = MagicMock()
        failure_response.json.return_value = [{"status": "FAILURE", "result": "Fehler"}]

        with patch("requests.get", side_effect=[started_response, failure_response]):
            with patch("time.sleep"):
                success, error = consumer.wait_for_task("task-start-fail")

        assert success is False
        assert "Fehler" in error

    def test_timeout_wenn_task_nicht_abgeschlossen(self):
        """Gibt (False, Timeout-Meldung) zurück wenn der Task den Timeout überschreitet."""
        mock_response = MagicMock()
        mock_response.json.return_value = [{"status": "STARTED"}]

        with patch("requests.get", return_value=mock_response):
            with patch("time.sleep"):
                with patch.object(consumer, "TASK_TIMEOUT", 3):
                    with patch.object(consumer, "TASK_INTERVAL", 3):
                        success, error = consumer.wait_for_task("task-timeout")

        assert success is False
        assert "Timeout" in error

    def test_leere_taskliste_wird_erneut_versucht(self):
        """Bei leerer Task-Liste wird nach TASK_INTERVAL erneut angefragt."""
        empty_response = MagicMock()
        empty_response.json.return_value = []

        success_response = MagicMock()
        success_response.json.return_value = [{"status": "SUCCESS"}]

        with patch("requests.get", side_effect=[empty_response, success_response]):
            with patch("time.sleep"):
                success, _ = consumer.wait_for_task("task-empty")

        assert success is True

    def test_failure_ohne_result_liefert_fallback_meldung(self):
        """Bei FAILURE ohne 'result'-Feld wird ein Fallback-Fehlertext zurückgegeben."""
        mock_response = MagicMock()
        mock_response.json.return_value = [{"status": "FAILURE"}]

        with patch("requests.get", return_value=mock_response):
            with patch("time.sleep"):
                success, error = consumer.wait_for_task("task-no-result")

        assert success is False
        assert error is not None


# ---------------------------------------------------------------------------
# mark_as_read / move_message
# ---------------------------------------------------------------------------


class TestMarkAsReadAndMove:
    def test_mark_as_read_ruft_graph_patch_auf(self):
        """mark_as_read ruft graph_patch mit isRead=True für die korrekte Message-ID auf."""
        with patch("consumer.graph_patch") as mock_patch:
            consumer.mark_as_read("token", "msg-id-1")

        mock_patch.assert_called_once_with(
            "token",
            f"users/{consumer.USER_EMAIL}/messages/msg-id-1",
            {"isRead": True},
        )

    def test_move_message_ruft_graph_post_auf(self):
        """move_message ruft graph_post mit destinationId für die korrekte Message-ID auf."""
        with patch("consumer.graph_post") as mock_post:
            consumer.move_message("token", "msg-id-2", "dest-folder-id")

        mock_post.assert_called_once_with(
            "token",
            f"users/{consumer.USER_EMAIL}/messages/msg-id-2/move",
            {"destinationId": "dest-folder-id"},
        )


# ---------------------------------------------------------------------------
# Hilfsfunktion für Anhang-Fixtures
# ---------------------------------------------------------------------------


def _make_attachment(name="rechnung.pdf", content_type="application/pdf", data=b"pdf-inhalt"):
    """Erstellt einen Mock-Anhang im Graph-API-Format."""
    return {
        "name": name,
        "contentType": content_type,
        "contentBytes": base64.b64encode(data).decode(),
    }


# ---------------------------------------------------------------------------
# process_messages
# ---------------------------------------------------------------------------


class TestProcessMessages:
    def test_keine_nachrichten_loggt_meldung(self, caplog):
        """Loggt eine Meldung wenn der Ordner keine Nachrichten enthält."""
        import logging
        with caplog.at_level(logging.INFO, logger="consumer"):
            with patch("consumer.get_messages", return_value=[]):
                consumer.process_messages("token", "src", "done", "err")

        assert "Keine Nachrichten" in caplog.text

    def test_erfolgreicher_upload_verschiebt_nach_verarbeitet(self):
        """Mail mit erfolgreich verarbeitetem Anhang wird nach 'verarbeitet' verschoben."""
        msg = {
            "id": "msg-1",
            "subject": "Testrechnung",
            "attachments": [_make_attachment()],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            with patch("consumer.upload_to_paperless", return_value="task-ok"):
                with patch("consumer.wait_for_task", return_value=(True, None)):
                    with patch("consumer.mark_as_read") as mock_read:
                        with patch("consumer.move_message") as mock_move:
                            with patch("consumer._log_import_result"):
                                consumer.process_messages("token", "src", "done-id", "err-id")

        mock_read.assert_called_once_with("token", "msg-1")
        mock_move.assert_called_once_with("token", "msg-1", "done-id")

    def test_duplikat_task_failure_verschiebt_nach_fehlerhaft(self):
        """Mail deren Task mit FAILURE (Duplikat) endet wird nach 'fehlerhaft' verschoben."""
        msg = {
            "id": "msg-2",
            "subject": "Duplikat",
            "attachments": [_make_attachment()],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            with patch("consumer.upload_to_paperless", return_value="task-fail"):
                with patch("consumer.wait_for_task", return_value=(False, "duplicate document")):
                    with patch("consumer.mark_as_read") as mock_read:
                        with patch("consumer.move_message") as mock_move:
                            with patch("consumer._log_import_result"):
                                consumer.process_messages("token", "src", "done-id", "err-id")

        mock_read.assert_called_once_with("token", "msg-2")
        mock_move.assert_called_once_with("token", "msg-2", "err-id")

    def test_nicht_unterstuetzter_anhang_wird_uebersprungen(self):
        """Anhänge mit nicht unterstütztem MIME-Typ werden ignoriert – Mail bleibt liegen."""
        msg = {
            "id": "msg-3",
            "subject": "Mit ZIP",
            "attachments": [_make_attachment(name="archiv.zip", content_type="application/zip")],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            with patch("consumer.upload_to_paperless") as mock_upload:
                with patch("consumer.mark_as_read") as mock_read:
                    with patch("consumer.move_message") as mock_move:
                        with patch("consumer._log_import_result"):
                            consumer.process_messages("token", "src", "done-id", "err-id")

        mock_upload.assert_not_called()
        mock_read.assert_not_called()
        mock_move.assert_not_called()

    def test_upload_exception_verschiebt_nach_fehlerhaft(self):
        """Bei Exception während des Uploads wird Mail nach 'fehlerhaft' verschoben."""
        msg = {
            "id": "msg-4",
            "subject": "Upload-Fehler",
            "attachments": [_make_attachment()],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            with patch("consumer.upload_to_paperless", side_effect=Exception("Verbindungsfehler")):
                with patch("consumer.mark_as_read") as mock_read:
                    with patch("consumer.move_message") as mock_move:
                        with patch("consumer._log_import_result"):
                            consumer.process_messages("token", "src", "done-id", "err-id")

        mock_read.assert_called_once_with("token", "msg-4")
        mock_move.assert_called_once_with("token", "msg-4", "err-id")

    def test_mehrere_anhaenge_alle_erfolgreich(self):
        """Mail mit mehreren Anhängen wird nur einmal nach 'verarbeitet' verschoben."""
        msg = {
            "id": "msg-5",
            "subject": "Mehrere Anhänge",
            "attachments": [
                _make_attachment("seite1.pdf"),
                _make_attachment("seite2.pdf"),
            ],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            with patch("consumer.upload_to_paperless", return_value="task-ok"):
                with patch("consumer.wait_for_task", return_value=(True, None)):
                    with patch("consumer.mark_as_read") as mock_read:
                        with patch("consumer.move_message") as mock_move:
                            with patch("consumer._log_import_result"):
                                consumer.process_messages("token", "src", "done-id", "err-id")

        # Nur einmal verschoben, trotz zweier Anhänge
        mock_move.assert_called_once_with("token", "msg-5", "done-id")
        mock_read.assert_called_once_with("token", "msg-5")

    def test_gemischtes_ergebnis_geht_nach_fehlerhaft(self):
        """Bei gemischtem Ergebnis (ein Erfolg, ein Fehler) wird Mail nach 'fehlerhaft' verschoben."""
        msg = {
            "id": "msg-6",
            "subject": "Gemischt",
            "attachments": [
                _make_attachment("ok.pdf"),
                _make_attachment("fail.pdf"),
            ],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            with patch("consumer.upload_to_paperless", return_value="task"):
                with patch(
                    "consumer.wait_for_task",
                    side_effect=[(True, None), (False, "duplicate")],
                ):
                    with patch("consumer.mark_as_read"):
                        with patch("consumer.move_message") as mock_move:
                            with patch("consumer._log_import_result"):
                                consumer.process_messages("token", "src", "done-id", "err-id")

        mock_move.assert_called_once_with("token", "msg-6", "err-id")

    def test_mail_ohne_anhaenge_wird_uebersprungen(self):
        """Mail ohne Anhänge wird weder verschoben noch als gelesen markiert."""
        msg = {
            "id": "msg-7",
            "subject": "Ohne Anhang",
            "attachments": [],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            with patch("consumer.mark_as_read") as mock_read:
                with patch("consumer.move_message") as mock_move:
                    with patch("consumer._log_import_result"):
                        consumer.process_messages("token", "src", "done-id", "err-id")

        mock_read.assert_not_called()
        mock_move.assert_not_called()

    def test_mehrere_mails_werden_einzeln_verarbeitet(self):
        """Jede Mail wird unabhängig verarbeitet – Fehler in einer beeinflusst keine andere."""
        msg_ok = {
            "id": "msg-ok",
            "subject": "Erfolg",
            "attachments": [_make_attachment("ok.pdf")],
        }
        msg_fail = {
            "id": "msg-fail",
            "subject": "Fehler",
            "attachments": [_make_attachment("fail.pdf")],
        }

        with patch("consumer.get_messages", return_value=[msg_ok, msg_fail]):
            with patch("consumer.upload_to_paperless", return_value="task"):
                with patch(
                    "consumer.wait_for_task",
                    side_effect=[(True, None), (False, "duplicate")],
                ):
                    with patch("consumer.mark_as_read"):
                        with patch("consumer.move_message") as mock_move:
                            with patch("consumer._log_import_result"):
                                consumer.process_messages("token", "src", "done-id", "err-id")

        assert mock_move.call_count == 2
        mock_move.assert_any_call("token", "msg-ok", "done-id")
        mock_move.assert_any_call("token", "msg-fail", "err-id")

    def test_log_import_bei_erfolg_aufgerufen(self):
        """_log_import_result wird für erfolgreich verarbeitete Anhänge aufgerufen."""
        msg = {
            "id": "msg-log-ok",
            "subject": "Rechnung Q1",
            "attachments": [_make_attachment("rechnung.pdf")],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            with patch("consumer.upload_to_paperless", return_value="task-x"):
                with patch("consumer.wait_for_task", return_value=(True, None)):
                    with patch("consumer.mark_as_read"):
                        with patch("consumer.move_message"):
                            with patch("consumer._log_import_result") as mock_log:
                                consumer.process_messages("token", "src", "done", "err")

        mock_log.assert_called_once()
        eintrag = mock_log.call_args[0][0]
        assert eintrag["datei"] == "rechnung.pdf"
        assert eintrag["betreff"] == "Rechnung Q1"
        assert eintrag["status"] == "erfolgreich"
        assert eintrag["fehler"] is None

    def test_log_import_bei_fehler_aufgerufen(self):
        """_log_import_result wird für fehlgeschlagene Anhänge mit Fehlermeldung aufgerufen."""
        msg = {
            "id": "msg-log-fail",
            "subject": "Duplikat-Mail",
            "attachments": [_make_attachment("duplikat.pdf")],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            with patch("consumer.upload_to_paperless", return_value="task-y"):
                with patch("consumer.wait_for_task", return_value=(False, "duplicate document")):
                    with patch("consumer.mark_as_read"):
                        with patch("consumer.move_message"):
                            with patch("consumer._log_import_result") as mock_log:
                                consumer.process_messages("token", "src", "done", "err")

        mock_log.assert_called_once()
        eintrag = mock_log.call_args[0][0]
        assert eintrag["datei"] == "duplikat.pdf"
        assert eintrag["status"] == "fehlerhaft"
        assert "duplicate" in eintrag["fehler"]

    def test_log_import_bei_upload_exception(self):
        """Exception beim Upload erzeugt einen Log-Eintrag mit Fehlermeldung."""
        msg = {
            "id": "msg-log-exc",
            "subject": "Netzwerkfehler",
            "attachments": [_make_attachment("datei.pdf")],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            with patch("consumer.upload_to_paperless", side_effect=Exception("Timeout")):
                with patch("consumer.mark_as_read"):
                    with patch("consumer.move_message"):
                        with patch("consumer._log_import_result") as mock_log:
                            consumer.process_messages("token", "src", "done", "err")

        mock_log.assert_called_once()
        eintrag = mock_log.call_args[0][0]
        assert eintrag["status"] == "fehlerhaft"
        assert "Timeout" in eintrag["fehler"]


# ---------------------------------------------------------------------------
# graph_send_mail
# ---------------------------------------------------------------------------


class TestGraphSendMail:
    def test_sendet_korrekte_mail_struktur(self):
        """graph_send_mail sendet den korrekten Payload an die sendMail-Endpoint."""
        mock_response = MagicMock()
        mock_response.status_code = 202

        with patch("requests.post", return_value=mock_response) as mock_post:
            consumer.graph_send_mail("token", "empfaenger@test.de", "Betreff", "<p>Inhalt</p>")

        mock_post.assert_called_once()
        url = mock_post.call_args.args[0]
        assert "sendMail" in url
        payload = mock_post.call_args.kwargs["json"]
        assert payload["message"]["subject"] == "Betreff"
        assert payload["message"]["body"]["contentType"] == "HTML"
        assert payload["message"]["body"]["content"] == "<p>Inhalt</p>"
        assert payload["message"]["toRecipients"][0]["emailAddress"]["address"] == "empfaenger@test.de"

    def test_raise_bei_http_fehler(self):
        """graph_send_mail wirft Exception bei HTTP-Fehler."""
        mock_response = MagicMock()
        mock_response.raise_for_status.side_effect = Exception("HTTP 403")

        with patch("requests.post", return_value=mock_response):
            with pytest.raises(Exception, match="HTTP 403"):
                consumer.graph_send_mail("token", "x@y.de", "Betreff", "<p>x</p>")

    def test_auth_header_wird_gesetzt(self):
        """graph_send_mail setzt den Authorization-Header korrekt."""
        mock_response = MagicMock()

        with patch("requests.post", return_value=mock_response) as mock_post:
            consumer.graph_send_mail("mein-token", "x@y.de", "Betreff", "<p>x</p>")

        headers = mock_post.call_args.kwargs["headers"]
        assert headers["Authorization"] == "Bearer mein-token"


# ---------------------------------------------------------------------------
# _build_summary_html
# ---------------------------------------------------------------------------


class TestBuildSummaryHtml:
    def test_enthaelt_datum(self):
        """HTML enthält das übergebene Datum."""
        html = consumer._build_summary_html([], [], "01.04.2026")
        assert "01.04.2026" in html

    def test_enthaelt_anzahl_erfolgreich(self):
        """HTML zeigt die korrekte Anzahl erfolgreich importierter Dateien."""
        erfolgreich = [{"datei": "a.pdf", "betreff": "S", "zeitpunkt": "01.04.2026 09:00"}]
        html = consumer._build_summary_html(erfolgreich, [], "01.04.2026")
        assert "a.pdf" in html

    def test_enthaelt_fehler_details(self):
        """HTML zeigt Dateiname und Fehlermeldung für fehlgeschlagene Importe."""
        fehlerhaft = [
            {"datei": "fehler.pdf", "betreff": "S", "fehler": "duplicate document", "zeitpunkt": "01.04.2026 10:00"}
        ]
        html = consumer._build_summary_html([], fehlerhaft, "01.04.2026")
        assert "fehler.pdf" in html
        assert "duplicate document" in html

    def test_leere_stats_zeigt_keine_eintraege(self):
        """HTML mit leeren Listen enthält 'keine'-Platzhalter."""
        html = consumer._build_summary_html([], [], "01.04.2026")
        assert "keine" in html


# ---------------------------------------------------------------------------
# send_daily_summary
# ---------------------------------------------------------------------------


class TestSendDailySummary:
    def test_versendet_mail_an_summary_recipient(self):
        """Sendet die Zusammenfassung an SUMMARY_RECIPIENT wenn gesetzt."""
        with patch("consumer._read_log_entries", return_value=([], [])):
            with patch("consumer._log_import_result"):
                with patch("consumer.graph_send_mail") as mock_send:
                    with patch.object(consumer, "SUMMARY_RECIPIENT", "chef@firma.de"):
                        consumer.send_daily_summary("token")

        mock_send.assert_called_once()
        assert mock_send.call_args.args[1] == "chef@firma.de"

    def test_faellt_auf_user_email_zurueck(self):
        """Sendet an USER_EMAIL wenn SUMMARY_RECIPIENT nicht gesetzt (None)."""
        with patch("consumer._read_log_entries", return_value=([], [])):
            with patch("consumer._log_import_result"):
                with patch("consumer.graph_send_mail") as mock_send:
                    with patch.object(consumer, "SUMMARY_RECIPIENT", None):
                        consumer.send_daily_summary("token")

        mock_send.assert_called_once()
        assert mock_send.call_args.args[1] == consumer.USER_EMAIL

    def test_betreff_enthaelt_datum(self):
        """Betreff der Zusammenfassungs-Mail enthält das heutige Datum."""
        heute = datetime.date.today().strftime("%d.%m.%Y")

        with patch("consumer._read_log_entries", return_value=([], [])):
            with patch("consumer._log_import_result"):
                with patch("consumer.graph_send_mail") as mock_send:
                    consumer.send_daily_summary("token")

        betreff = mock_send.call_args.args[2]
        assert heute in betreff

    def test_schreibt_summary_sentinel_in_log(self):
        """Nach dem Versand wird ein summary_gesendet-Eintrag in die Log-Datei geschrieben."""
        with patch("consumer._read_log_entries", return_value=([], [])):
            with patch("consumer.graph_send_mail"):
                with patch("consumer._log_import_result") as mock_log:
                    consumer.send_daily_summary("token")

        eintrag = mock_log.call_args.args[0]
        assert eintrag["typ"] == "summary_gesendet"
        assert eintrag["datum"] == datetime.date.today().isoformat()

    def test_wirft_exception_bei_send_fehler(self):
        """Wirft Exception wenn graph_send_mail fehlschlägt."""
        with patch("consumer._read_log_entries", return_value=([], [])):
            with patch("consumer._log_import_result"):
                with patch("consumer.graph_send_mail", side_effect=Exception("Senderfehler")):
                    with pytest.raises(Exception, match="Senderfehler"):
                        consumer.send_daily_summary("token")


# ---------------------------------------------------------------------------
# _log_import_result
# ---------------------------------------------------------------------------


class TestLogImportResult:
    def test_schreibt_json_zeile_in_datei(self, tmp_path):
        """Schreibt den Eintrag als gültiges JSON in eine neue Datei."""
        logfile = tmp_path / "import.log"
        eintrag = {"typ": "import", "ts": "2026-04-01T09:05:00", "datei": "test.pdf", "status": "erfolgreich"}

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            consumer._log_import_result(eintrag)

        lines = logfile.read_text(encoding="utf-8").strip().splitlines()
        assert len(lines) == 1
        assert json.loads(lines[0]) == eintrag

    def test_haengt_mehrere_zeilen_an(self, tmp_path):
        """Mehrere Aufrufe erzeugen mehrere Zeilen (append-Modus)."""
        logfile = tmp_path / "import.log"

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            consumer._log_import_result({"typ": "import", "datei": "a.pdf"})
            consumer._log_import_result({"typ": "import", "datei": "b.pdf"})

        lines = logfile.read_text(encoding="utf-8").strip().splitlines()
        assert len(lines) == 2

    def test_fehler_wird_geloggt_nicht_ausgeloest(self, caplog):
        """Bei nicht beschreibbarem Pfad wird der Fehler geloggt, keine Exception geworfen."""
        import logging
        with caplog.at_level(logging.ERROR, logger="consumer"):
            with patch.object(consumer, "IMPORT_LOG_FILE", "/nicht/existierender/pfad/import.log"):
                consumer._log_import_result({"typ": "import"})

        assert "Fehler" in caplog.text


# ---------------------------------------------------------------------------
# _read_log_entries
# ---------------------------------------------------------------------------


class TestReadLogEntries:
    def test_leere_datei_gibt_leere_listen(self, tmp_path):
        """Leere Log-Datei ergibt zwei leere Listen."""
        logfile = tmp_path / "import.log"
        logfile.write_text("", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            erfolgreich, fehlerhaft = consumer._read_log_entries("2026-04-01")

        assert erfolgreich == []
        assert fehlerhaft == []

    def test_nicht_vorhandene_datei_gibt_leere_listen(self, tmp_path):
        """Fehlende Log-Datei ergibt zwei leere Listen ohne Exception."""
        with patch.object(consumer, "IMPORT_LOG_FILE", str(tmp_path / "nichtda.log")):
            erfolgreich, fehlerhaft = consumer._read_log_entries("2026-04-01")

        assert erfolgreich == []
        assert fehlerhaft == []

    def test_liest_erfolgreich_eintraege(self, tmp_path):
        """Erfolgreich-Einträge werden korrekt in die erste Liste übernommen."""
        logfile = tmp_path / "import.log"
        eintrag = {
            "typ": "import", "ts": "2026-04-01T09:05:00",
            "datei": "ok.pdf", "betreff": "Test", "status": "erfolgreich", "fehler": None,
        }
        logfile.write_text(json.dumps(eintrag) + "\n", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            erfolgreich, fehlerhaft = consumer._read_log_entries("2026-04-01")

        assert len(erfolgreich) == 1
        assert erfolgreich[0]["datei"] == "ok.pdf"
        assert fehlerhaft == []

    def test_liest_fehlerhaft_eintraege(self, tmp_path):
        """Fehlerhaft-Einträge werden korrekt in die zweite Liste übernommen."""
        logfile = tmp_path / "import.log"
        eintrag = {
            "typ": "import", "ts": "2026-04-01T10:00:00",
            "datei": "fail.pdf", "betreff": "Test", "status": "fehlerhaft", "fehler": "duplicate",
        }
        logfile.write_text(json.dumps(eintrag) + "\n", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            erfolgreich, fehlerhaft = consumer._read_log_entries("2026-04-01")

        assert fehlerhaft[0]["fehler"] == "duplicate"
        assert erfolgreich == []

    def test_filtert_andere_daten_heraus(self, tmp_path):
        """Einträge anderer Tage werden nicht in die Ergebnislisten übernommen."""
        logfile = tmp_path / "import.log"
        zeilen = [
            json.dumps({"typ": "import", "ts": "2026-04-01T09:00:00", "datei": "heute.pdf", "status": "erfolgreich", "fehler": None}),
            json.dumps({"typ": "import", "ts": "2026-03-31T09:00:00", "datei": "gestern.pdf", "status": "erfolgreich", "fehler": None}),
        ]
        logfile.write_text("\n".join(zeilen) + "\n", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            erfolgreich, fehlerhaft = consumer._read_log_entries("2026-04-01")

        assert len(erfolgreich) == 1
        assert erfolgreich[0]["datei"] == "heute.pdf"

    def test_ignoriert_summary_eintraege(self, tmp_path):
        """summary_gesendet-Einträge werden nicht in die Import-Listen übernommen."""
        logfile = tmp_path / "import.log"
        logfile.write_text(json.dumps({"typ": "summary_gesendet", "datum": "2026-04-01"}) + "\n", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            erfolgreich, fehlerhaft = consumer._read_log_entries("2026-04-01")

        assert erfolgreich == []
        assert fehlerhaft == []

    def test_ignoriert_ungueltiges_json(self, tmp_path):
        """Zeilen mit ungültigem JSON werden ohne Exception übersprungen."""
        logfile = tmp_path / "import.log"
        logfile.write_text("kein json hier\n", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            erfolgreich, fehlerhaft = consumer._read_log_entries("2026-04-01")

        assert erfolgreich == []


# ---------------------------------------------------------------------------
# _get_last_summary_date
# ---------------------------------------------------------------------------


class TestGetLastSummaryDate:
    def test_gibt_none_zurueck_wenn_datei_fehlt(self, tmp_path):
        """Gibt None zurück wenn die Log-Datei noch nicht existiert."""
        with patch.object(consumer, "IMPORT_LOG_FILE", str(tmp_path / "nichtda.log")):
            assert consumer._get_last_summary_date() is None

    def test_gibt_none_zurueck_ohne_summary_eintrag(self, tmp_path):
        """Gibt None zurück wenn nur Import-Einträge, aber kein summary_gesendet vorhanden."""
        logfile = tmp_path / "import.log"
        logfile.write_text(
            json.dumps({"typ": "import", "ts": "2026-04-01T09:00:00"}) + "\n",
            encoding="utf-8",
        )

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            assert consumer._get_last_summary_date() is None

    def test_gibt_letztes_summary_datum_zurueck(self, tmp_path):
        """Gibt das Datum des letzten summary_gesendet-Eintrags zurück."""
        logfile = tmp_path / "import.log"
        zeilen = [
            json.dumps({"typ": "summary_gesendet", "datum": "2026-03-31"}),
            json.dumps({"typ": "summary_gesendet", "datum": "2026-04-01"}),
        ]
        logfile.write_text("\n".join(zeilen) + "\n", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            result = consumer._get_last_summary_date()

        assert result == datetime.date(2026, 4, 1)

    def test_ignoriert_ungueltiges_datum(self, tmp_path):
        """Einträge mit ungültigem Datumsformat werden ohne Exception übersprungen."""
        logfile = tmp_path / "import.log"
        logfile.write_text(
            json.dumps({"typ": "summary_gesendet", "datum": "kein-datum"}) + "\n",
            encoding="utf-8",
        )

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            assert consumer._get_last_summary_date() is None
