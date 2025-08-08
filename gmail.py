import base64
import re
from typing import List, Optional, Tuple

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

# ---- TODO: set these ----
SPREADSHEET_ID = "1tjg8LSjxGgVBXVJlrkE54Uf8sayarHc3Nkf_hdSGLso"
SHEET_RANGE = "Sheet1!A:C"  # Subject | From | Body (plain text)

def get_gmail_service():
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
    creds = flow.run_local_server(port=0)
    return build("gmail", "v1", credentials=creds)

def get_sheets_service():
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
    creds = flow.run_local_server(port=0)
    return build("sheets", "v4", credentials=creds)

def get_label_id(gmail, label_name: str) -> Optional[str]:
    """Find the Gmail label ID for a given label name (case-sensitive)."""
    labels = gmail.users().labels().list(userId="me").execute().get("labels", [])
    for lab in labels:
        if lab.get("name") == label_name:
            return lab.get("id")
    return None

def list_message_ids(gmail, label_id: str) -> List[str]:
    """Return all message IDs under the given label (handles pagination)."""
    ids = []
    page_token = None
    while True:
        resp = gmail.users().messages().list(
            userId="me", labelIds=[label_id], pageToken=page_token, maxResults=500
        ).execute()
        ids.extend([m["id"] for m in resp.get("messages", [])])
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return ids

def get_header(payload_headers, name: str) -> str:
    for h in payload_headers:
        if h.get("name", "").lower() == name.lower():
            return h.get("value", "")
    return ""

def decode_part(data_b64: str) -> str:
    # Gmail uses URL-safe base64
    return base64.urlsafe_b64decode(data_b64.encode("utf-8")).decode("utf-8", errors="ignore")

def strip_html(html: str) -> str:
    # quick-and-dirty HTML -> text
    text = re.sub(r"(?is)<(script|style).*?>.*?(</\1>)", "", html)
    text = re.sub(r"(?s)<br\s*/?>", "\n", text)
    text = re.sub(r"(?s)</p\s*>", "\n\n", text)
    text = re.sub(r"(?s)<.*?>", "", text)
    return re.sub(r"[ \t]+\n", "\n", text).strip()

def extract_plain_text_from_payload(payload: dict) -> str:
    """
    Try to return a reasonable plain-text body.
    Priority: text/plain -> text/html -> fallback to snippet.
    Handles multipart messages recursively.
    """
    mime_type = payload.get("mimeType", "")
    body = payload.get("body", {})
    data = body.get("data")

    if mime_type == "text/plain" and data:
        return decode_part(data).strip()

    if mime_type == "text/html" and data:
        return strip_html(decode_part(data))

    # multipart: walk parts
    if "parts" in payload:
        plain_candidate = ""
        html_candidate = ""
        for part in payload["parts"]:
            mt = part.get("mimeType", "")
            if mt.startswith("multipart/"):
                sub = extract_plain_text_from_payload(part)
                if sub:
                    # return the first decent sub-body we find
                    return sub
            elif mt == "text/plain" and part.get("body", {}).get("data"):
                plain_candidate = decode_part(part["body"]["data"]).strip() or plain_candidate
            elif mt == "text/html" and part.get("body", {}).get("data"):
                html_candidate = strip_html(decode_part(part["body"]["data"])) or html_candidate

        if plain_candidate:
            return plain_candidate
        if html_candidate:
            return html_candidate

    return ""  # caller can fallback to snippet

def fetch_message_fields(gmail, msg_id: str) -> Tuple[str, str, str]:
    """Return (subject, from, body_text) for a given message id."""
    msg = gmail.users().messages().get(userId="me", id=msg_id, format="full").execute()
    payload = msg.get("payload", {})
    headers = payload.get("headers", [])

    subject = get_header(headers, "Subject")
    sender = get_header(headers, "From")
    body_text = extract_plain_text_from_payload(payload).strip()
    if not body_text:
        body_text = (msg.get("snippet") or "").strip()

    # Optional: keep body short to avoid gigantic rows
    # body_text = body_text[:5000]

    return subject, sender, body_text

def append_rows_to_sheet(sheets, rows: List[List[str]]):
    if not rows:
        return
    body = {"values": rows}
    sheets.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_RANGE,
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body=body,
    ).execute()

def main():
    gmail = get_gmail_service()
    sheets = get_sheets_service()

    label_id = get_label_id(gmail, "internships")
    if not label_id:
        raise RuntimeError('Could not find Gmail label named "internships".')

    msg_ids = list_message_ids(gmail, label_id)
    print(f"Found {len(msg_ids)} messages under label 'internships'.")

    rows = []
    for mid in msg_ids:
        try:
            subject, sender, body = fetch_message_fields(gmail, mid)
            rows.append([subject, sender, body])
        except HttpError as e:
            print(f"Skipping message {mid} due to API error: {e}")

    append_rows_to_sheet(sheets, rows)
    print("Done.")

if __name__ == "__main__":
    main()