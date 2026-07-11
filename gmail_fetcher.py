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
    "kalaimughilan@greatlakes.edu.in",  # Digii Tickets (service request export)
    "itsupport@greatlakes.edu.in",      # Digii Tickets — also sent from IT Support's shared inbox
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


class AuthNeedsHumanError(RuntimeError):
    """
    Raised when the stored token can't be silently refreshed and completing
    auth would require a human to click through a browser consent screen.
    On an unattended host there is nobody to do that — callers should treat
    this as distinct from a generic failure (e.g. skip retrying the same way
    an API hiccup would be retried, and word the alert email accordingly).
    """


def authenticate() -> object:
    """
    OAuth flow — reuses token.json's refresh token silently when possible.
    Never opens a browser without a bound timeout: on a headless/unattended
    host there is no one to complete an interactive consent screen, and the
    underlying run_local_server() call has no timeout of its own (blocks
    forever by default) — so a dead/missing token must fail fast and loud
    rather than hang the whole pipeline indefinitely.
    """
    creds = None

    if TOKEN_FILE.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
        except ValueError as e:
            # Corrupted/incomplete token.json (e.g. missing refresh_token
            # field) — treat exactly like "no token" rather than crashing
            # with a raw ValueError that looks like a code bug.
            log.warning("token.json is unreadable/incomplete (%s) — treating as missing", e)
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            log.info("Refreshing expired token …")
            creds.refresh(Request())
        else:
            log.error(
                "No valid token.json and no refresh_token available — "
                "interactive browser consent is required, which is not "
                "possible on an unattended run."
            )
            raise AuthNeedsHumanError(
                "Gmail OAuth needs a human to re-authorize: token.json is missing, "
                "invalid, or has no usable refresh_token. Run gmail_fetcher.py "
                "interactively once on a machine with a browser to mint a fresh "
                "token.json, then redeploy it to this host."
            )

        TOKEN_FILE.write_text(creds.to_json())
        log.info("Token saved → %s", TOKEN_FILE)

    return build("gmail", "v1", credentials=creds)


def bootstrap_token(timeout_seconds: int = 180) -> None:
    """
    One-time INTERACTIVE setup: opens a real browser for OAuth consent and
    writes token.json. Run this manually (`python -c "import gmail_fetcher;
    gmail_fetcher.bootstrap_token()"`) on a machine with a browser before the
    very first unattended run — authenticate() itself never does this, so a
    scheduled/headless run can't accidentally trigger a hanging browser popup.
    """
    flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
    creds = flow.run_local_server(port=0, timeout_seconds=timeout_seconds)
    TOKEN_FILE.write_text(creds.to_json())
    log.info("Token saved → %s (bootstrap complete)", TOKEN_FILE)


def load_processed_ids() -> set:
    if PROCESSED_IDS.exists():
        return set(json.loads(PROCESSED_IDS.read_text()))
    return set()

def save_processed_ids(ids: set):
    # Atomic write — a process kill mid-write must never leave this file
    # truncated/corrupt, since a JSONDecodeError on the next run's
    # load_processed_ids() would crash Gmail fetching entirely.
    tmp = PROCESSED_IDS.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(list(ids)))
    os.replace(str(tmp), str(PROCESSED_IDS))

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
