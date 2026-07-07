"""
Phase 1 — Gmail Fetcher
Connects to energymonitoring.glc@greatlakes.edu.in via Gmail API,
finds emails from sodexo@greatlakes.edu.in with Excel attachments,
and downloads them to the GLIM folder.
"""

import os
import sys
import base64
import json
import logging
from datetime import datetime
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ── Config ────────────────────────────────────────────────────────────────────
BASE_DIR        = Path(__file__).parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE       = BASE_DIR / "token.json"
DOWNLOAD_DIR     = BASE_DIR / "downloads"
PROCESSED_IDS    = BASE_DIR / "processed_ids.json"
SCOPES           = ["https://www.googleapis.com/auth/gmail.modify"]

# All senders whose Excel attachments should be fetched.
# Electrical/STP report comes from sodexo; WTP water report may come
# from a different address — add it here if the sender changes.
SENDER_FILTERS = [
    "sodexo@greatlakes.edu.in",
    # WTP report sender — add if they email from a distinct address:
    # "wtp@greatlakes.edu.in",
]

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(BASE_DIR / "fetcher.log", encoding="utf-8"),
        logging.StreamHandler(stream=open(sys.stdout.fileno(), mode="w", encoding="utf-8", closefd=False)),
    ],
)
log = logging.getLogger(__name__)


def authenticate() -> object:
    """OAuth flow — opens browser on first run, reuses token.json after."""
    creds = None

    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            log.info("Refreshing expired token …")
            creds.refresh(Request())
        else:
            log.info("Opening browser for first-time authorization …")
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE), SCOPES
            )
            creds = flow.run_local_server(port=0)

        TOKEN_FILE.write_text(creds.to_json())
        log.info("Token saved → %s", TOKEN_FILE)

    return build("gmail", "v1", credentials=creds)


def load_processed_ids() -> set:
    if PROCESSED_IDS.exists():
        return set(json.loads(PROCESSED_IDS.read_text()))
    return set()

def save_processed_ids(ids: set):
    PROCESSED_IDS.write_text(json.dumps(list(ids)))

def search_emails(service) -> list:
    """Return unprocessed message IDs from all configured senders (last 60 days)."""
    from_clause = " OR ".join(f"from:{s}" for s in SENDER_FILTERS)
    query = f"({from_clause}) has:attachment newer_than:60d"
    result = service.users().messages().list(userId="me", q=query).execute()
    messages = result.get("messages", [])
    processed = load_processed_ids()
    new_messages = [m for m in messages if m["id"] not in processed]
    log.info("Found %d email(s) matching sender filters, %d new (unprocessed)", len(messages), len(new_messages))
    return new_messages


def download_attachments(service, msg_id: str) -> list[str]:
    """Download all .xlsx / .xls attachments from a single message."""
    DOWNLOAD_DIR.mkdir(exist_ok=True)
    saved = []

    msg = service.users().messages().get(userId="me", id=msg_id).execute()
    subject = next(
        (h["value"] for h in msg["payload"]["headers"] if h["name"] == "Subject"),
        "(no subject)",
    )
    log.info("Processing email: %s", subject)

    parts = msg["payload"].get("parts", [])

    def extract_parts(parts):
        for part in parts:
            # recurse into multipart
            if part.get("parts"):
                extract_parts(part["parts"])
            fname = part.get("filename", "")
            if fname.lower().endswith((".xlsx", ".xls")):
                att_id = part["body"].get("attachmentId")
                if not att_id:
                    continue
                att = service.users().messages().attachments().get(
                    userId="me", messageId=msg_id, id=att_id
                ).execute()
                data = base64.urlsafe_b64decode(att["data"])
                stem = Path(fname).stem
                suffix = Path(fname).suffix
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_path = DOWNLOAD_DIR / f"{stem}_{ts}{suffix}"
                out_path.write_bytes(data)
                log.info("  Downloaded → %s (%.1f KB)", fname, len(data) / 1024)
                saved.append(str(out_path))

    extract_parts(parts)
    return saved


def mark_as_processed(msg_id: str):
    """Record message ID so it isn't processed again, regardless of read/unread state."""
    ids = load_processed_ids()
    ids.add(msg_id)
    save_processed_ids(ids)


def run():
    log.info("═" * 60)
    log.info("GL Dashboard — Gmail Fetcher started")
    log.info("Sender filters: %s", ", ".join(SENDER_FILTERS))
    log.info("Download dir  : %s", DOWNLOAD_DIR)

    service = authenticate()
    messages = search_emails(service)

    if not messages:
        log.info("No new emails — nothing to do.")
        return []

    all_files = []
    for msg in messages:
        files = download_attachments(service, msg["id"])
        if files:
            mark_as_processed(msg["id"])
            all_files.extend(files)

    log.info("─" * 60)
    log.info("Done. %d file(s) downloaded:", len(all_files))
    for f in all_files:
        log.info("  %s", f)
    log.info("═" * 60)
    return all_files


if __name__ == "__main__":
    run()
