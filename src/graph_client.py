"""Microsoft Graph API client — SharePoint file watching + Outlook email sending."""

import os
import json
import time
import tempfile
import requests
import msal


class GraphClient:
    GRAPH_URL = "https://graph.microsoft.com/v1.0"
    SCOPES = ["https://graph.microsoft.com/.default"]

    def __init__(self):
        self._app = msal.ConfidentialClientApplication(
            client_id=os.environ["AZURE_CLIENT_ID"],
            client_credential=os.environ["AZURE_CLIENT_SECRET"],
            authority=f"https://login.microsoftonline.com/{os.environ['AZURE_TENANT_ID']}",
        )
        self._token: str | None = None
        self._token_expiry: float = 0

    def _get_token(self) -> str:
        if self._token and time.time() < self._token_expiry - 60:
            return self._token
        result = self._app.acquire_token_for_client(scopes=self.SCOPES)
        if "access_token" not in result:
            raise RuntimeError(f"Impossible d'obtenir un token Graph: {result.get('error_description')}")
        self._token = result["access_token"]
        self._token_expiry = time.time() + result.get("expires_in", 3600)
        return self._token

    def _headers(self) -> dict:
        return {"Authorization": f"Bearer {self._get_token()}", "Content-Type": "application/json"}

    def _get(self, path: str, **kwargs) -> dict:
        resp = requests.get(f"{self.GRAPH_URL}{path}", headers=self._headers(), **kwargs)
        resp.raise_for_status()
        return resp.json()

    def list_transcript_files(self) -> list[dict]:
        """List .vtt and .docx files in the configured SharePoint transcript folder."""
        site_id = os.environ["SHAREPOINT_SITE_ID"]
        drive_id = os.environ["SHAREPOINT_DRIVE_ID"]
        folder = os.environ["SHAREPOINT_FOLDER_PATH"].strip("/")

        path = f"/sites/{site_id}/drives/{drive_id}/root:/{folder}:/children"
        data = self._get(path)

        return [
            item for item in data.get("value", [])
            if item.get("name", "").lower().endswith((".vtt", ".docx"))
        ]

    def download_file(self, download_url: str, suffix: str) -> str:
        """Download a file to a temp path and return that path."""
        resp = requests.get(download_url, headers={"Authorization": f"Bearer {self._get_token()}"})
        resp.raise_for_status()
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(resp.content)
        tmp.close()
        return tmp.name

    def send_email(self, to_address: str, subject: str, html_body: str) -> None:
        """Send an email via the configured sender using Microsoft Graph."""
        sender = os.environ["SENDER_EMAIL"]
        payload = {
            "message": {
                "subject": subject,
                "body": {"contentType": "HTML", "content": html_body},
                "toRecipients": [{"emailAddress": {"address": to_address}}],
            }
        }
        resp = requests.post(
            f"{self.GRAPH_URL}/users/{sender}/sendMail",
            headers=self._headers(),
            json=payload,
        )
        resp.raise_for_status()

    def list_inbox_messages(self, user_email: str, top: int = 50) -> list[dict]:
        """Fetch recent inbox messages for a user."""
        data = self._get(
            f"/users/{user_email}/mailFolders/inbox/messages",
            params={
                "$top": top,
                "$orderby": "receivedDateTime desc",
                "$select": "id,subject,from,receivedDateTime,isRead,bodyPreview,importance,conversationId",
            },
        )
        return data.get("value", [])

    def get_message_body(self, user_email: str, message_id: str) -> str:
        """Fetch the full body of a specific message."""
        data = self._get(
            f"/users/{user_email}/messages/{message_id}",
            params={"$select": "body,subject,from,toRecipients"},
        )
        return data.get("body", {}).get("content", "")

    def create_draft(self, user_email: str, to_address: str, subject: str, body: str, reply_to_id: str | None = None) -> str:
        """Create a draft reply in Outlook. NEVER sends. Returns the draft message ID."""
        if reply_to_id:
            # Create as a reply draft to keep thread context
            resp = requests.post(
                f"{self.GRAPH_URL}/users/{user_email}/messages/{reply_to_id}/createReply",
                headers=self._headers(),
                json={},
            )
            resp.raise_for_status()
            draft_id = resp.json()["id"]
            # Update the draft body
            resp2 = requests.patch(
                f"{self.GRAPH_URL}/users/{user_email}/messages/{draft_id}",
                headers=self._headers(),
                json={"body": {"contentType": "Text", "content": body}},
            )
            resp2.raise_for_status()
            return draft_id
        else:
            payload = {
                "subject": subject,
                "isDraft": True,
                "body": {"contentType": "Text", "content": body},
                "toRecipients": [{"emailAddress": {"address": to_address}}],
            }
            resp = requests.post(
                f"{self.GRAPH_URL}/users/{user_email}/messages",
                headers=self._headers(),
                json=payload,
            )
            resp.raise_for_status()
            return resp.json()["id"]

    def resolve_participant_email(self, display_name: str) -> str | None:
        """Try to find the email address for a display name in the directory."""
        try:
            data = self._get(
                f"/users",
                params={"$filter": f"displayName eq '{display_name}'", "$select": "mail,displayName"},
            )
            users = data.get("value", [])
            if users:
                return users[0].get("mail")
        except Exception:
            pass
        return None
