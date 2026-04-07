"""
Tests for consumer.py

All external API calls (Microsoft Graph, Paperless, MSAL) are mocked with
unittest.mock. Tests cover success and failure paths.
"""

import base64
import datetime
import json
import os

import pytest

# Set environment variables before import so module constants are populated
os.environ.setdefault("AZURE_CLIENT_ID", "test-client-id")
os.environ.setdefault("AZURE_CLIENT_SECRET", "test-client-secret")
os.environ.setdefault("AZURE_TENANT_ID", "test-tenant-id")
os.environ.setdefault("USER_EMAIL", "test@example.com")
os.environ.setdefault("PAPERLESS_URL", "http://localhost:8000")
os.environ.setdefault("PAPERLESS_TOKEN", "test-token")
os.environ.setdefault("MAIL_FOLDER", "paperless")
os.environ.setdefault("INBOX_TAG_ID", "2")

from unittest.mock import MagicMock, call, patch

import consumer


# ---------------------------------------------------------------------------
# get_token
# ---------------------------------------------------------------------------


class TestGetToken:
    def test_erfolgreicher_token_abruf(self):
        """Returns the access token when MSAL responds successfully."""
        mock_app = MagicMock()
        mock_app.acquire_token_for_client.return_value = {"access_token": "mein-token"}

        with patch("msal.ConfidentialClientApplication", return_value=mock_app):
            token = consumer.get_token()

        assert token == "mein-token"

    def test_token_fehler_wirft_exception(self):
        """Raises an exception when no access_token is present in the MSAL result."""
        mock_app = MagicMock()
        mock_app.acquire_token_for_client.return_value = {
            "error": "invalid_client",
            "error_description": "Client authentication failed",
        }

        with patch("msal.ConfidentialClientApplication", return_value=mock_app):
            with pytest.raises(Exception, match="Token error"):
                consumer.get_token()


# ---------------------------------------------------------------------------
# graph_get / graph_patch / graph_post
# ---------------------------------------------------------------------------


class TestGraphHelpers:
    def test_graph_get_gibt_json_zurueck(self):
        """graph_get calls the correct URL and returns JSON."""
        mock_response = MagicMock()
        mock_response.json.return_value = {"value": [{"id": "123"}]}

        with patch("requests.get", return_value=mock_response) as mock_get:
            result = consumer.graph_get("token123", "users/test@example.com/mailFolders")

        mock_get.assert_called_once()
        assert "Bearer token123" in mock_get.call_args.kwargs["headers"]["Authorization"]
        assert result == {"value": [{"id": "123"}]}

    def test_graph_get_raise_bei_http_fehler(self):
        """graph_get raises an exception on HTTP error."""
        mock_response = MagicMock()
        mock_response.raise_for_status.side_effect = Exception("HTTP 403")

        with patch("requests.get", return_value=mock_response):
            with pytest.raises(Exception, match="HTTP 403"):
                consumer.graph_get("token", "some/path")

    def test_graph_patch_sendet_payload(self):
        """graph_patch sends the correct JSON payload."""
        mock_response = MagicMock()

        with patch("requests.patch", return_value=mock_response) as mock_patch:
            consumer.graph_patch("token123", "users/x/messages/id1", {"isRead": True})

        mock_patch.assert_called_once()
        assert mock_patch.call_args.kwargs["json"] == {"isRead": True}

    def test_graph_patch_raise_bei_http_fehler(self):
        """graph_patch raises an exception on HTTP error."""
        mock_response = MagicMock()
        mock_response.raise_for_status.side_effect = Exception("HTTP 500")

        with patch("requests.patch", return_value=mock_response):
            with pytest.raises(Exception, match="HTTP 500"):
                consumer.graph_patch("token", "path", {})

    def test_graph_post_gibt_json_zurueck(self):
        """graph_post sends payload and returns JSON response."""
        mock_response = MagicMock()
        mock_response.json.return_value = {"id": "neuer-ordner-id"}

        with patch("requests.post", return_value=mock_response) as mock_post:
            result = consumer.graph_post("token", "some/path", {"displayName": "Test"})

        assert result == {"id": "neuer-ordner-id"}
        assert mock_post.call_args.kwargs["json"] == {"displayName": "Test"}

    def test_graph_post_raise_bei_http_fehler(self):
        """graph_post raises an exception on HTTP error."""
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
        """Finds a folder directly at the top level."""
        top_level = {"value": [{"id": "folder-1", "displayName": "paperless"}]}

        with patch("consumer.graph_get", return_value=top_level):
            folder_id = consumer.get_folder_id("token", "paperless")

        assert folder_id == "folder-1"

    def test_findet_unterordner(self):
        """Finds a folder as a child of a top-level folder."""
        top_level = {"value": [{"id": "inbox-id", "displayName": "Posteingang"}]}
        child_folders = {"value": [{"id": "eco-id", "displayName": "paperless"}]}

        def graph_get_side_effect(token, path):
            if "childFolders" in path:
                return child_folders
            return top_level

        with patch("consumer.graph_get", side_effect=graph_get_side_effect):
            folder_id = consumer.get_folder_id("token", "paperless")

        assert folder_id == "eco-id"

    def test_wirft_exception_wenn_nicht_gefunden(self):
        """Raises an exception when the folder is not found at any level."""
        empty = {"value": []}

        with patch("consumer.graph_get", return_value=empty):
            with pytest.raises(Exception, match="not found"):
                consumer.get_folder_id("token", "DoesNotExist")

    def test_top_level_hat_vorrang_vor_unterordner(self):
        """When the folder exists at both levels, the top-level match is returned."""
        top_level = {
            "value": [
                {"id": "top-id", "displayName": "paperless"},
                {"id": "inbox-id", "displayName": "Posteingang"},
            ]
        }
        child_folders = {"value": [{"id": "child-id", "displayName": "paperless"}]}

        def graph_get_side_effect(token, path):
            if "childFolders" in path:
                return child_folders
            return top_level

        with patch("consumer.graph_get", side_effect=graph_get_side_effect):
            folder_id = consumer.get_folder_id("token", "paperless")

        # Top-level folder takes precedence
        assert folder_id == "top-id"


# ---------------------------------------------------------------------------
# get_or_create_subfolder
# ---------------------------------------------------------------------------


class TestGetOrCreateSubfolder:
    def test_gibt_vorhandenen_unterordner_zurueck(self):
        """Returns the ID of an already existing subfolder without creating a new one."""
        existing = {"value": [{"id": "sub-id", "displayName": "verarbeitet"}]}

        with patch("consumer.graph_get", return_value=existing):
            with patch("consumer.graph_post") as mock_post:
                result = consumer.get_or_create_subfolder("token", "parent-id", "verarbeitet")

        assert result == "sub-id"
        mock_post.assert_not_called()

    def test_legt_neuen_unterordner_an(self):
        """Creates a new subfolder when it does not yet exist."""
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
        """Returns the task UUID, stripping surrounding quotes."""
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
        """Works even when Paperless returns the UUID without quotes."""
        mock_response = MagicMock()
        mock_response.text = "abc-456-uuid"

        with patch("requests.post", return_value=mock_response):
            task_id = consumer.upload_to_paperless("test.pdf", b"inhalt", "application/pdf")

        assert task_id == "abc-456-uuid"

    def test_raise_bei_http_fehler(self):
        """Raises an exception on HTTP error from the Paperless server."""
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
        """Returns (True, None) when the task reports SUCCESS immediately."""
        mock_response = MagicMock()
        mock_response.json.return_value = [{"status": "SUCCESS", "result": "ok"}]

        with patch("requests.get", return_value=mock_response):
            with patch("time.sleep"):
                success, error = consumer.wait_for_task("task-123")

        assert success is True
        assert error is None

    def test_failure_bei_duplikat(self):
        """Returns (False, error message) when the task reports FAILURE (e.g. duplicate)."""
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
        """Waits on PENDING status and returns (True, None) once SUCCESS follows."""
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
        """Waits on STARTED status and returns (False, ...) when FAILURE follows."""
        started_response = MagicMock()
        started_response.json.return_value = [{"status": "STARTED"}]

        failure_response = MagicMock()
        failure_response.json.return_value = [{"status": "FAILURE", "result": "task error"}]

        with patch("requests.get", side_effect=[started_response, failure_response]):
            with patch("time.sleep"):
                success, error = consumer.wait_for_task("task-start-fail")

        assert success is False
        assert "task error" in error

    def test_timeout_wenn_task_nicht_abgeschlossen(self):
        """Returns (False, timeout message) when the task exceeds the timeout."""
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
        """Retries after TASK_INTERVAL when the task list is empty."""
        empty_response = MagicMock()
        empty_response.json.return_value = []

        success_response = MagicMock()
        success_response.json.return_value = [{"status": "SUCCESS"}]

        with patch("requests.get", side_effect=[empty_response, success_response]):
            with patch("time.sleep"):
                success, _ = consumer.wait_for_task("task-empty")

        assert success is True

    def test_failure_ohne_result_liefert_fallback_meldung(self):
        """Returns a fallback error text for FAILURE entries missing the 'result' field."""
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
        """mark_as_read calls graph_patch with isRead=True for the correct message ID."""
        with patch("consumer.graph_patch") as mock_patch:
            consumer.mark_as_read("token", "msg-id-1")

        mock_patch.assert_called_once_with(
            "token",
            f"users/{consumer.USER_EMAIL}/messages/msg-id-1",
            {"isRead": True},
        )

    def test_move_message_ruft_graph_post_auf(self):
        """move_message calls graph_post with destinationId for the correct message ID."""
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
    """Creates a mock attachment in Graph API format."""
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
        """Logs a message when the folder contains no messages."""
        import logging
        with caplog.at_level(logging.INFO, logger="consumer"):
            with patch("consumer.get_messages", return_value=[]):
                consumer.process_messages("token", "src", "done", "err")

        assert "No messages" in caplog.text

    def test_erfolgreicher_upload_verschiebt_nach_verarbeitet(self):
        """Mail with a successfully processed attachment is moved to 'verarbeitet'."""
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
                            with patch("consumer._write_log_entry"):
                                consumer.process_messages("token", "src", "done-id", "err-id")

        mock_read.assert_called_once_with("token", "msg-1")
        mock_move.assert_called_once_with("token", "msg-1", "done-id")

    def test_duplikat_task_failure_verschiebt_nach_fehlerhaft(self):
        """Mail whose task ends with FAILURE (duplicate) is moved to 'fehlerhaft'."""
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
                            with patch("consumer._write_log_entry"):
                                consumer.process_messages("token", "src", "done-id", "err-id")

        mock_read.assert_called_once_with("token", "msg-2")
        mock_move.assert_called_once_with("token", "msg-2", "err-id")

    def test_nicht_unterstuetzter_anhang_wird_uebersprungen(self):
        """Attachments with unsupported MIME types are ignored — mail is left in place."""
        msg = {
            "id": "msg-3",
            "subject": "Mit ZIP",
            "attachments": [_make_attachment(name="archiv.zip", content_type="application/zip")],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            with patch("consumer.upload_to_paperless") as mock_upload:
                with patch("consumer.mark_as_read") as mock_read:
                    with patch("consumer.move_message") as mock_move:
                        with patch("consumer._write_log_entry"):
                            consumer.process_messages("token", "src", "done-id", "err-id")

        mock_upload.assert_not_called()
        mock_read.assert_not_called()
        mock_move.assert_not_called()

    def test_upload_exception_verschiebt_nach_fehlerhaft(self):
        """An exception during upload moves the mail to 'fehlerhaft'."""
        msg = {
            "id": "msg-4",
            "subject": "Upload-Fehler",
            "attachments": [_make_attachment()],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            with patch("consumer.upload_to_paperless", side_effect=Exception("Verbindungsfehler")):
                with patch("consumer.mark_as_read") as mock_read:
                    with patch("consumer.move_message") as mock_move:
                        with patch("consumer._write_log_entry"):
                            consumer.process_messages("token", "src", "done-id", "err-id")

        mock_read.assert_called_once_with("token", "msg-4")
        mock_move.assert_called_once_with("token", "msg-4", "err-id")

    def test_multiple_attachments_all_successful(self):
        """Mail with multiple attachments is only moved to 'processed' once."""
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
                            with patch("consumer._write_log_entry"):
                                consumer.process_messages("token", "src", "done-id", "err-id")

        # Moved only once despite two attachments
        mock_move.assert_called_once_with("token", "msg-5", "done-id")
        mock_read.assert_called_once_with("token", "msg-5")

    def test_mixed_result_moves_to_failed(self):
        """Mail is moved to 'failed' when results are mixed (one success, one failure)."""
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
                            with patch("consumer._write_log_entry"):
                                consumer.process_messages("token", "src", "done-id", "err-id")

        mock_move.assert_called_once_with("token", "msg-6", "err-id")

    def test_mail_without_attachments_is_skipped(self):
        """Mail without attachments is neither moved nor marked as read."""
        msg = {
            "id": "msg-7",
            "subject": "Ohne Anhang",
            "attachments": [],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            with patch("consumer.mark_as_read") as mock_read:
                with patch("consumer.move_message") as mock_move:
                    with patch("consumer._write_log_entry"):
                        consumer.process_messages("token", "src", "done-id", "err-id")

        mock_read.assert_not_called()
        mock_move.assert_not_called()

    def test_multiple_mails_processed_independently(self):
        """Each mail is processed independently – a failure in one does not affect others."""
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
                            with patch("consumer._write_log_entry"):
                                consumer.process_messages("token", "src", "done-id", "err-id")

        assert mock_move.call_count == 2
        mock_move.assert_any_call("token", "msg-ok", "done-id")
        mock_move.assert_any_call("token", "msg-fail", "err-id")

    def test_log_entry_called_on_success(self):
        """_write_log_entry is called for successfully processed attachments."""
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
                            with patch("consumer._write_log_entry") as mock_log:
                                consumer.process_messages("token", "src", "done", "err")

        mock_log.assert_called_once()
        entry = mock_log.call_args[0][0]
        assert entry["file"] == "rechnung.pdf"
        assert entry["subject"] == "Rechnung Q1"
        assert entry["status"] == "success"
        assert entry["error"] is None

    def test_log_entry_called_on_failure(self):
        """_write_log_entry is called for failed attachments with an error message."""
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
                            with patch("consumer._write_log_entry") as mock_log:
                                consumer.process_messages("token", "src", "done", "err")

        mock_log.assert_called_once()
        entry = mock_log.call_args[0][0]
        assert entry["file"] == "duplikat.pdf"
        assert entry["status"] == "failed"
        assert "duplicate" in entry["error"]

    def test_log_entry_called_on_upload_exception(self):
        """An exception during upload creates a log entry with an error message."""
        msg = {
            "id": "msg-log-exc",
            "subject": "Netzwerkfehler",
            "attachments": [_make_attachment("datei.pdf")],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            with patch("consumer.upload_to_paperless", side_effect=Exception("Timeout")):
                with patch("consumer.mark_as_read"):
                    with patch("consumer.move_message"):
                        with patch("consumer._write_log_entry") as mock_log:
                            consumer.process_messages("token", "src", "done", "err")

        mock_log.assert_called_once()
        entry = mock_log.call_args[0][0]
        assert entry["status"] == "failed"
        assert "Timeout" in entry["error"]


# ---------------------------------------------------------------------------
# graph_send_mail
# ---------------------------------------------------------------------------


class TestGraphSendMail:
    def test_sends_correct_mail_structure(self):
        """graph_send_mail sends the correct payload to the sendMail endpoint."""
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

    def test_raises_on_http_error(self):
        """graph_send_mail raises an exception on HTTP error."""
        mock_response = MagicMock()
        mock_response.raise_for_status.side_effect = Exception("HTTP 403")

        with patch("requests.post", return_value=mock_response):
            with pytest.raises(Exception, match="HTTP 403"):
                consumer.graph_send_mail("token", "x@y.de", "Betreff", "<p>x</p>")

    def test_auth_header_is_set(self):
        """graph_send_mail correctly sets the Authorization header."""
        mock_response = MagicMock()

        with patch("requests.post", return_value=mock_response) as mock_post:
            consumer.graph_send_mail("mein-token", "x@y.de", "Betreff", "<p>x</p>")

        headers = mock_post.call_args.kwargs["headers"]
        assert headers["Authorization"] == "Bearer mein-token"


# ---------------------------------------------------------------------------
# _build_summary_html
# ---------------------------------------------------------------------------


class TestBuildSummaryHtml:
    def test_contains_date(self):
        """HTML contains the provided date."""
        html = consumer._build_summary_html([], [], "01.04.2026")
        assert "01.04.2026" in html

    def test_contains_success_count(self):
        """HTML shows the correct count of successfully imported files."""
        succeeded = [{"file": "a.pdf", "subject": "S", "timestamp": "01.04.2026 09:00"}]
        html = consumer._build_summary_html(succeeded, [], "01.04.2026")
        assert "a.pdf" in html

    def test_contains_error_details(self):
        """HTML shows filename and error message for failed imports."""
        failed = [
            {"file": "fehler.pdf", "subject": "S", "error": "duplicate document", "timestamp": "01.04.2026 10:00"}
        ]
        html = consumer._build_summary_html([], failed, "01.04.2026")
        assert "fehler.pdf" in html
        assert "duplicate document" in html

    def test_empty_stats_shows_no_entries(self):
        """HTML with empty lists contains a 'keine' placeholder."""
        html = consumer._build_summary_html([], [], "01.04.2026")
        assert "keine" in html

    def test_pending_table_shown_when_pending(self):
        """HTML contains the pending messages table when pending entries exist."""
        pending = [{"subject": "Ohne Anhang", "sender": "a@b.de", "reason": "Kein Anhang gefunden"}]
        html = consumer._build_summary_html([], [], "01.04.2026", pending=pending)
        assert "Nicht verarbeitete Mails" in html
        assert "Ohne Anhang" in html
        assert "Kein Anhang gefunden" in html
        assert "a@b.de" in html

    def test_pending_table_hidden_when_empty(self):
        """HTML does not contain the pending section when no pending messages exist."""
        html = consumer._build_summary_html([], [], "01.04.2026", pending=[])
        assert "Nicht verarbeitete Mails" not in html

    def test_pending_default_none_hides_section(self):
        """HTML does not contain the pending section when pending is omitted."""
        html = consumer._build_summary_html([], [], "01.04.2026")
        assert "Nicht verarbeitete Mails" not in html


# ---------------------------------------------------------------------------
# send_daily_summary
# ---------------------------------------------------------------------------


class TestSendDailySummary:
    def test_sends_mail_to_summary_recipient(self):
        """Sends the summary to SUMMARY_RECIPIENT when set."""
        with patch("consumer._read_log_entries_since_last_summary", return_value=([], [])):
            with patch("consumer._analyze_pending_messages", return_value=[]):
                with patch("consumer._write_log_entry"):
                    with patch("consumer.graph_send_mail") as mock_send:
                        with patch.object(consumer, "SUMMARY_RECIPIENT", "chef@firma.de"):
                            consumer.send_daily_summary("token", "folder-id")

        mock_send.assert_called_once()
        assert mock_send.call_args.args[1] == "chef@firma.de"

    def test_falls_back_to_user_email(self):
        """Falls back to USER_EMAIL when SUMMARY_RECIPIENT is not set (None)."""
        with patch("consumer._read_log_entries_since_last_summary", return_value=([], [])):
            with patch("consumer._analyze_pending_messages", return_value=[]):
                with patch("consumer._write_log_entry"):
                    with patch("consumer.graph_send_mail") as mock_send:
                        with patch.object(consumer, "SUMMARY_RECIPIENT", None):
                            consumer.send_daily_summary("token", "folder-id")

        mock_send.assert_called_once()
        assert mock_send.call_args.args[1] == consumer.USER_EMAIL

    def test_subject_contains_date(self):
        """Subject of the summary email contains today's date."""
        heute = datetime.date.today().strftime("%d.%m.%Y")

        with patch("consumer._read_log_entries_since_last_summary", return_value=([], [])):
            with patch("consumer._analyze_pending_messages", return_value=[]):
                with patch("consumer._write_log_entry"):
                    with patch("consumer.graph_send_mail") as mock_send:
                        consumer.send_daily_summary("token", "folder-id")

        subject_line = mock_send.call_args.args[2]
        assert heute in subject_line

    def test_writes_summary_sentinel_to_log(self):
        """After sending, a summary_sent entry is written to the log file."""
        with patch("consumer._read_log_entries_since_last_summary", return_value=([], [])):
            with patch("consumer._analyze_pending_messages", return_value=[]):
                with patch("consumer.graph_send_mail"):
                    with patch("consumer._write_log_entry") as mock_log:
                        consumer.send_daily_summary("token", "folder-id")

        entry = mock_log.call_args.args[0]
        assert entry["type"] == "summary_sent"
        assert entry["date"] == datetime.date.today().isoformat()

    def test_raises_exception_on_send_error(self):
        """Raises an exception when graph_send_mail fails."""
        with patch("consumer._read_log_entries_since_last_summary", return_value=([], [])):
            with patch("consumer._analyze_pending_messages", return_value=[]):
                with patch("consumer._write_log_entry"):
                    with patch("consumer.graph_send_mail", side_effect=Exception("Senderfehler")):
                        with pytest.raises(Exception, match="Senderfehler"):
                            consumer.send_daily_summary("token", "folder-id")

    def test_pending_messages_included_in_html(self):
        """Pending messages from the folder are passed to the HTML builder."""
        pending = [{"subject": "Ohne Anhang", "sender": "a@b.de", "reason": "Kein Anhang gefunden"}]

        with patch("consumer._read_log_entries_since_last_summary", return_value=([], [])):
            with patch("consumer._analyze_pending_messages", return_value=pending):
                with patch("consumer._write_log_entry"):
                    with patch("consumer.graph_send_mail") as mock_send:
                        consumer.send_daily_summary("token", "folder-id")

        html_body = mock_send.call_args.args[3]
        assert "Ohne Anhang" in html_body
        assert "Kein Anhang gefunden" in html_body

    def test_no_folder_id_skips_pending_analysis(self):
        """When folder_id is not provided, pending analysis is skipped."""
        with patch("consumer._read_log_entries_since_last_summary", return_value=([], [])):
            with patch("consumer._analyze_pending_messages") as mock_analyze:
                with patch("consumer._write_log_entry"):
                    with patch("consumer.graph_send_mail"):
                        consumer.send_daily_summary("token")

        mock_analyze.assert_not_called()

    def test_pending_analysis_error_does_not_block_summary(self):
        """An error in _analyze_pending_messages does not prevent the summary from being sent."""
        with patch("consumer._read_log_entries_since_last_summary", return_value=([], [])):
            with patch("consumer._analyze_pending_messages", side_effect=Exception("Graph API error")):
                with patch("consumer._write_log_entry"):
                    with patch("consumer.graph_send_mail") as mock_send:
                        consumer.send_daily_summary("token", "folder-id")

        mock_send.assert_called_once()


# ---------------------------------------------------------------------------
# _write_log_entry
# ---------------------------------------------------------------------------


class TestWriteLogEntry:
    def test_writes_json_line_to_file(self, tmp_path):
        """Writes the entry as valid JSON to a new file."""
        logfile = tmp_path / "import.log"
        entry = {"type": "import", "ts": "2026-04-01T09:05:00", "file": "test.pdf", "status": "success"}

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            consumer._write_log_entry(entry)

        lines = logfile.read_text(encoding="utf-8").strip().splitlines()
        assert len(lines) == 1
        assert json.loads(lines[0]) == entry

    def test_appends_multiple_lines(self, tmp_path):
        """Multiple calls produce multiple lines (append mode)."""
        logfile = tmp_path / "import.log"

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            consumer._write_log_entry({"type": "import", "file": "a.pdf"})
            consumer._write_log_entry({"type": "import", "file": "b.pdf"})

        lines = logfile.read_text(encoding="utf-8").strip().splitlines()
        assert len(lines) == 2

    def test_error_is_logged_not_raised(self, caplog):
        """On a non-writable path the error is logged, no exception raised."""
        import logging
        with caplog.at_level(logging.ERROR, logger="consumer"):
            with patch.object(consumer, "IMPORT_LOG_FILE", "/nicht/existierender/pfad/import.log"):
                consumer._write_log_entry({"type": "import"})

        assert "Error" in caplog.text


# ---------------------------------------------------------------------------
# _read_log_entries
# ---------------------------------------------------------------------------


class TestReadLogEntries:
    def test_empty_file_returns_empty_lists(self, tmp_path):
        """Empty log file yields two empty lists."""
        logfile = tmp_path / "import.log"
        logfile.write_text("", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            succeeded, failed = consumer._read_log_entries("2026-04-01")

        assert succeeded == []
        assert failed == []

    def test_missing_file_returns_empty_lists(self, tmp_path):
        """Missing log file yields two empty lists without an exception."""
        with patch.object(consumer, "IMPORT_LOG_FILE", str(tmp_path / "nichtda.log")):
            succeeded, failed = consumer._read_log_entries("2026-04-01")

        assert succeeded == []
        assert failed == []

    def test_reads_success_entries(self, tmp_path):
        """Success entries are correctly added to the first list."""
        logfile = tmp_path / "import.log"
        entry = {
            "type": "import", "ts": "2026-04-01T09:05:00",
            "file": "ok.pdf", "subject": "Test", "status": "success", "error": None,
        }
        logfile.write_text(json.dumps(entry) + "\n", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            succeeded, failed = consumer._read_log_entries("2026-04-01")

        assert len(succeeded) == 1
        assert succeeded[0]["file"] == "ok.pdf"
        assert failed == []

    def test_reads_failed_entries(self, tmp_path):
        """Failed entries are correctly added to the second list."""
        logfile = tmp_path / "import.log"
        entry = {
            "type": "import", "ts": "2026-04-01T10:00:00",
            "file": "fail.pdf", "subject": "Test", "status": "failed", "error": "duplicate",
        }
        logfile.write_text(json.dumps(entry) + "\n", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            succeeded, failed = consumer._read_log_entries("2026-04-01")

        assert failed[0]["error"] == "duplicate"
        assert succeeded == []

    def test_filters_out_other_dates(self, tmp_path):
        """Entries from other days are not included in the result lists."""
        logfile = tmp_path / "import.log"
        lines = [
            json.dumps({"type": "import", "ts": "2026-04-01T09:00:00", "file": "heute.pdf", "status": "success", "error": None}),
            json.dumps({"type": "import", "ts": "2026-03-31T09:00:00", "file": "gestern.pdf", "status": "success", "error": None}),
        ]
        logfile.write_text("\n".join(lines) + "\n", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            succeeded, failed = consumer._read_log_entries("2026-04-01")

        assert len(succeeded) == 1
        assert succeeded[0]["file"] == "heute.pdf"

    def test_ignores_summary_entries(self, tmp_path):
        """summary_sent entries are not included in the import lists."""
        logfile = tmp_path / "import.log"
        logfile.write_text(json.dumps({"type": "summary_sent", "date": "2026-04-01"}) + "\n", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            succeeded, failed = consumer._read_log_entries("2026-04-01")

        assert succeeded == []
        assert failed == []

    def test_ignores_invalid_json(self, tmp_path):
        """Lines with invalid JSON are skipped without raising an exception."""
        logfile = tmp_path / "import.log"
        logfile.write_text("kein json hier\n", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            succeeded, failed = consumer._read_log_entries("2026-04-01")

        assert succeeded == []


# ---------------------------------------------------------------------------
# _read_log_entries_since_last_summary
# ---------------------------------------------------------------------------


class TestReadLogEntriesSinceLastSummary:
    def test_empty_file_returns_empty_lists(self, tmp_path):
        """Empty log file yields two empty lists."""
        logfile = tmp_path / "import.log"
        logfile.write_text("", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            succeeded, failed = consumer._read_log_entries_since_last_summary()

        assert succeeded == []
        assert failed == []

    def test_missing_file_returns_empty_lists(self, tmp_path):
        """Missing log file yields two empty lists without an exception."""
        with patch.object(consumer, "IMPORT_LOG_FILE", str(tmp_path / "nichtda.log")):
            succeeded, failed = consumer._read_log_entries_since_last_summary()

        assert succeeded == []
        assert failed == []

    def test_all_entries_read_without_summary_sentinel(self, tmp_path):
        """Without a prior summary_sent entry all import entries are returned."""
        logfile = tmp_path / "import.log"
        entry = {
            "type": "import", "ts": "2026-04-01T09:00:00",
            "file": "a.pdf", "subject": "Test", "status": "success", "error": None,
        }
        logfile.write_text(json.dumps(entry) + "\n", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            succeeded, failed = consumer._read_log_entries_since_last_summary()

        assert len(succeeded) == 1
        assert succeeded[0]["file"] == "a.pdf"

    def test_entries_before_last_summary_are_excluded(self, tmp_path):
        """Import entries before the last summary_sent are not returned."""
        logfile = tmp_path / "import.log"
        lines = [
            json.dumps({"type": "import", "ts": "2026-04-01T08:00:00", "file": "alt.pdf", "status": "success", "error": None}),
            json.dumps({"type": "summary_sent", "date": "2026-04-01"}),
            json.dumps({"type": "import", "ts": "2026-04-01T20:00:00", "file": "neu.pdf", "status": "success", "error": None}),
        ]
        logfile.write_text("\n".join(lines) + "\n", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            succeeded, failed = consumer._read_log_entries_since_last_summary()

        assert len(succeeded) == 1
        assert succeeded[0]["file"] == "neu.pdf"

    def test_entries_after_last_summary_are_included(self, tmp_path):
        """The exact bug scenario: imports on the previous day after the summary was sent."""
        logfile = tmp_path / "import.log"
        lines = [
            json.dumps({"type": "summary_sent", "date": "2026-04-02"}),
            json.dumps({"type": "import", "ts": "2026-04-02T19:54:04+02:00", "file": "inv1.pdf", "subject": "Rechnung 1", "status": "success", "error": None}),
            json.dumps({"type": "import", "ts": "2026-04-02T20:29:58+02:00", "file": "inv2.pdf", "subject": "Rechnung 2", "status": "success", "error": None}),
        ]
        logfile.write_text("\n".join(lines) + "\n", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            succeeded, failed = consumer._read_log_entries_since_last_summary()

        assert len(succeeded) == 2
        assert succeeded[0]["file"] == "inv1.pdf"
        assert succeeded[1]["file"] == "inv2.pdf"

    def test_failed_entries_are_correctly_captured(self, tmp_path):
        """Failed imports after the last summary are added to the failed list."""
        logfile = tmp_path / "import.log"
        lines = [
            json.dumps({"type": "summary_sent", "date": "2026-04-01"}),
            json.dumps({"type": "import", "ts": "2026-04-01T22:00:00", "file": "dup.pdf", "subject": "Duplikat", "status": "failed", "error": "duplicate"}),
        ]
        logfile.write_text("\n".join(lines) + "\n", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            succeeded, failed = consumer._read_log_entries_since_last_summary()

        assert succeeded == []
        assert len(failed) == 1
        assert failed[0]["error"] == "duplicate"

    def test_ignores_invalid_json(self, tmp_path):
        """Lines with invalid JSON are skipped without raising an exception."""
        logfile = tmp_path / "import.log"
        logfile.write_text("kein json hier\n", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            succeeded, failed = consumer._read_log_entries_since_last_summary()

        assert succeeded == []
        assert failed == []


# ---------------------------------------------------------------------------
# _get_last_summary_date
# ---------------------------------------------------------------------------


class TestGetLastSummaryDate:
    def test_returns_none_when_file_missing(self, tmp_path):
        """Returns None when the log file does not yet exist."""
        with patch.object(consumer, "IMPORT_LOG_FILE", str(tmp_path / "nichtda.log")):
            assert consumer._get_last_summary_date() is None

    def test_returns_none_without_summary_entry(self, tmp_path):
        """Returns None when only import entries are present but no summary_sent entry."""
        logfile = tmp_path / "import.log"
        logfile.write_text(
            json.dumps({"type": "import", "ts": "2026-04-01T09:00:00"}) + "\n",
            encoding="utf-8",
        )

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            assert consumer._get_last_summary_date() is None

    def test_returns_last_summary_date(self, tmp_path):
        """Returns the date of the last summary_sent entry."""
        logfile = tmp_path / "import.log"
        lines = [
            json.dumps({"type": "summary_sent", "date": "2026-03-31"}),
            json.dumps({"type": "summary_sent", "date": "2026-04-01"}),
        ]
        logfile.write_text("\n".join(lines) + "\n", encoding="utf-8")

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            result = consumer._get_last_summary_date()

        assert result == datetime.date(2026, 4, 1)

    def test_ignores_invalid_date(self, tmp_path):
        """Entries with an invalid date format are skipped without raising an exception."""
        logfile = tmp_path / "import.log"
        logfile.write_text(
            json.dumps({"type": "summary_sent", "date": "kein-datum"}) + "\n",
            encoding="utf-8",
        )

        with patch.object(consumer, "IMPORT_LOG_FILE", str(logfile)):
            assert consumer._get_last_summary_date() is None


# ---------------------------------------------------------------------------
# _analyze_pending_messages
# ---------------------------------------------------------------------------


class TestAnalyzePendingMessages:
    def test_empty_folder_returns_empty_list(self):
        """Returns an empty list when the folder contains no messages."""
        with patch("consumer.get_messages", return_value=[]):
            result = consumer._analyze_pending_messages("token", "folder-id")

        assert result == []

    def test_message_without_attachments(self):
        """Message without attachments is listed with 'Kein Anhang gefunden'."""
        msg = {
            "id": "msg-1",
            "subject": "Nur Text",
            "from": {"emailAddress": {"address": "sender@test.de"}},
            "attachments": [],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            result = consumer._analyze_pending_messages("token", "folder-id")

        assert len(result) == 1
        assert result[0]["subject"] == "Nur Text"
        assert result[0]["sender"] == "sender@test.de"
        assert "Kein Anhang" in result[0]["reason"]

    def test_message_with_unsupported_attachment(self):
        """Message with only unsupported attachments lists them."""
        msg = {
            "id": "msg-2",
            "subject": "ZIP Mail",
            "from": {"emailAddress": {"address": "sender@test.de"}},
            "attachments": [
                {"name": "archiv.zip", "contentType": "application/zip", "contentBytes": ""},
            ],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            result = consumer._analyze_pending_messages("token", "folder-id")

        assert len(result) == 1
        assert "Keine unterstützten Anhänge" in result[0]["reason"]
        assert "archiv.zip" in result[0]["reason"]

    def test_message_with_supported_attachment_shows_unknown(self):
        """Message with supported attachments that remains in the folder shows 'Unbekannter Grund'."""
        msg = {
            "id": "msg-3",
            "subject": "PDF Mail",
            "from": {"emailAddress": {"address": "sender@test.de"}},
            "attachments": [
                {"name": "rechnung.pdf", "contentType": "application/pdf", "contentBytes": ""},
            ],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            result = consumer._analyze_pending_messages("token", "folder-id")

        assert len(result) == 1
        assert result[0]["reason"] == "Unbekannter Grund"

    def test_message_with_mixed_attachments(self):
        """Message with mixed supported and unsupported attachments lists partial reason."""
        msg = {
            "id": "msg-4",
            "subject": "Gemischt",
            "from": {"emailAddress": {"address": "sender@test.de"}},
            "attachments": [
                {"name": "rechnung.pdf", "contentType": "application/pdf", "contentBytes": ""},
                {"name": "archiv.zip", "contentType": "application/zip", "contentBytes": ""},
            ],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            result = consumer._analyze_pending_messages("token", "folder-id")

        assert len(result) == 1
        assert "Teilweise" in result[0]["reason"]
        assert "archiv.zip" in result[0]["reason"]

    def test_missing_from_field(self):
        """Message without 'from' field defaults sender to 'unbekannt'."""
        msg = {
            "id": "msg-5",
            "subject": "Anonym",
            "attachments": [],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            result = consumer._analyze_pending_messages("token", "folder-id")

        assert result[0]["sender"] == "unbekannt"

    def test_extension_fallback_counts_as_supported(self):
        """Attachment with generic contentType but supported extension is treated as supported."""
        msg = {
            "id": "msg-6",
            "subject": "Octet Stream PDF",
            "from": {"emailAddress": {"address": "sender@test.de"}},
            "attachments": [
                {"name": "scan.pdf", "contentType": "application/octet-stream", "contentBytes": ""},
            ],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            result = consumer._analyze_pending_messages("token", "folder-id")

        assert len(result) == 1
        # Extension-based fallback means it's recognised as supported
        assert result[0]["reason"] == "Unbekannter Grund"

    def test_word_attachment_counts_as_supported(self):
        """Word attachment (.docx) is treated as supported in pending analysis."""
        msg = {
            "id": "msg-7",
            "subject": "Word Mail",
            "from": {"emailAddress": {"address": "sender@test.de"}},
            "attachments": [
                {
                    "name": "dokument.docx",
                    "contentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "contentBytes": "",
                },
            ],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            result = consumer._analyze_pending_messages("token", "folder-id")

        assert len(result) == 1
        assert result[0]["reason"] == "Unbekannter Grund"


# ---------------------------------------------------------------------------
# convert_word_to_pdf
# ---------------------------------------------------------------------------


class TestConvertWordToPdf:
    def test_successful_conversion(self, tmp_path):
        """Returns PDF filename and bytes after successful LibreOffice conversion."""
        pdf_content = b"%PDF-1.4 test content"

        def fake_run(cmd, **kwargs):
            # Simulate LibreOffice creating the output PDF
            outdir = cmd[cmd.index("--outdir") + 1]
            pdf_path = os.path.join(outdir, "test.pdf")
            with open(pdf_path, "wb") as f:
                f.write(pdf_content)
            mock_result = MagicMock()
            mock_result.returncode = 0
            return mock_result

        with patch("subprocess.run", side_effect=fake_run):
            pdf_name, pdf_bytes = consumer.convert_word_to_pdf("test.docx", b"word-data")

        assert pdf_name == "test.pdf"
        assert pdf_bytes == pdf_content

    def test_conversion_failure_raises(self):
        """Raises RuntimeError when LibreOffice returns non-zero exit code."""
        mock_result = MagicMock()
        mock_result.returncode = 1
        mock_result.stderr = "LibreOffice error"

        with patch("subprocess.run", return_value=mock_result):
            with pytest.raises(RuntimeError, match="LibreOffice conversion failed"):
                consumer.convert_word_to_pdf("test.docx", b"word-data")

    def test_missing_output_raises(self):
        """Raises RuntimeError when the PDF output file is not created."""
        mock_result = MagicMock()
        mock_result.returncode = 0

        with patch("subprocess.run", return_value=mock_result):
            with pytest.raises(RuntimeError, match="PDF output not found"):
                consumer.convert_word_to_pdf("test.docx", b"word-data")

    def test_timeout_raises(self):
        """subprocess.TimeoutExpired is propagated when LibreOffice hangs."""
        import subprocess
        with patch("subprocess.run", side_effect=subprocess.TimeoutExpired("libreoffice", 120)):
            with pytest.raises(subprocess.TimeoutExpired):
                consumer.convert_word_to_pdf("test.docx", b"word-data")


# ---------------------------------------------------------------------------
# process_messages – Word conversion
# ---------------------------------------------------------------------------


class TestProcessMessagesWordConversion:
    def test_word_attachment_converted_and_uploaded(self):
        """Word attachment is converted to PDF and then uploaded successfully."""
        msg = {
            "id": "msg-word-1",
            "subject": "Word Dokument",
            "attachments": [_make_attachment(
                name="Lastenheft.docx",
                content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                data=b"word-content",
            )],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            with patch("consumer.convert_word_to_pdf", return_value=("Lastenheft.pdf", b"pdf-bytes")) as mock_convert:
                with patch("consumer.upload_to_paperless", return_value="task-ok") as mock_upload:
                    with patch("consumer.wait_for_task", return_value=(True, None)):
                        with patch("consumer.mark_as_read") as mock_read:
                            with patch("consumer.move_message") as mock_move:
                                with patch("consumer._write_log_entry"):
                                    consumer.process_messages("token", "src", "done-id", "err-id")

        mock_convert.assert_called_once()
        mock_upload.assert_called_once_with("Lastenheft.pdf", b"pdf-bytes", "application/pdf")
        mock_read.assert_called_once_with("token", "msg-word-1")
        mock_move.assert_called_once_with("token", "msg-word-1", "done-id")

    def test_word_conversion_failure_moves_to_error(self):
        """Word conversion failure is logged and mail is moved to error folder."""
        msg = {
            "id": "msg-word-2",
            "subject": "Kaputtes Word",
            "attachments": [_make_attachment(
                name="broken.docx",
                content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                data=b"broken",
            )],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            with patch("consumer.convert_word_to_pdf", side_effect=RuntimeError("conversion error")):
                with patch("consumer.upload_to_paperless") as mock_upload:
                    with patch("consumer.mark_as_read") as mock_read:
                        with patch("consumer.move_message") as mock_move:
                            with patch("consumer._write_log_entry") as mock_log:
                                consumer.process_messages("token", "src", "done-id", "err-id")

        mock_upload.assert_not_called()
        mock_read.assert_called_once_with("token", "msg-word-2")
        mock_move.assert_called_once_with("token", "msg-word-2", "err-id")
        entry = mock_log.call_args[0][0]
        assert entry["status"] == "failed"
        assert "conversion" in entry["error"].lower()

    def test_doc_extension_fallback_triggers_conversion(self):
        """A .doc file with generic contentType is resolved and converted via extension fallback."""
        msg = {
            "id": "msg-doc-1",
            "subject": "Altes Word",
            "attachments": [_make_attachment(
                name="alt.doc",
                content_type="application/octet-stream",
                data=b"old-word",
            )],
        }

        with patch("consumer.get_messages", return_value=[msg]):
            with patch("consumer.convert_word_to_pdf", return_value=("alt.pdf", b"pdf-bytes")) as mock_convert:
                with patch("consumer.upload_to_paperless", return_value="task-ok"):
                    with patch("consumer.wait_for_task", return_value=(True, None)):
                        with patch("consumer.mark_as_read"):
                            with patch("consumer.move_message"):
                                with patch("consumer._write_log_entry"):
                                    consumer.process_messages("token", "src", "done-id", "err-id")

        mock_convert.assert_called_once()
