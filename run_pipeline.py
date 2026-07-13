"""
Phase 4 — Master Pipeline Orchestrator
Chains Gmail fetch → Excel parse → Dashboard inject into a single automated run.
The pipeline's own logic (paths, no OS-specific calls) is portable, but the
current scheduling mechanism is Windows Task Scheduler only (see
setup_scheduler.ps1) — there is no Cloud Run/cron entrypoint, Dockerfile, or
HTTP handler in this repo yet. See README.md's "First-run / new-host
bootstrap" section before running this anywhere other than the current
lab PC.
"""

import sys
import re
import logging
import traceback
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).parent
LOG_FILE = BASE_DIR / "pipeline.log"

# ── Logging setup ─────────────────────────────────────────────────────────────
def setup_logging():
    fmt = "%(asctime)s  %(levelname)-8s  %(message)s"
    datefmt = "%Y-%m-%d %H:%M:%S"
    handlers = [
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(
            stream=open(sys.stdout.fileno(), mode="w", encoding="utf-8", closefd=False)
        ),
    ]
    logging.basicConfig(level=logging.INFO, format=fmt, datefmt=datefmt, handlers=handlers)

log = logging.getLogger(__name__)

SEPARATOR = "=" * 65
DIVIDER   = "-" * 65


ALERT_FALLBACK_FILE = BASE_DIR / "ALERT_UNDELIVERED.txt"
LAST_REPORT_SENT_FILE = BASE_DIR / "LAST_REPORT_SENT.txt"

# Hours (24h, local time) of the guaranteed sends — each one fires even if
# nothing changed since the last email, so there's always a dashboard copy
# at these two checkpoints regardless of how many (or how few) mid-day
# change-triggered sends happened in between. Matches the scheduler's
# 8AM-11PM hourly window (setup_scheduler.ps1) — both hours must fall
# inside that window to actually get a pipeline run to check them.
GUARANTEED_SEND_HOURS = [11, 20]  # 11 AM, 8 PM


def _significant_signal(data: dict) -> dict:
    """
    The small subset of `data` that counts as a "new day of real data landed"
    rather than noise. The vendor's source Excel is edited throughout the
    day (corrections to old rows, AMC/chair-count/street-light snapshots
    changing), so hashing the WHOLE parsed data dict made almost every
    hourly run look "changed" and fired an email far more often than
    intended. Only the day-count actually advancing means genuinely new
    electrical/EB&DG data arrived — that's the only thing that should
    trigger a mid-day send; everything else (including tickets) waits for
    the guaranteed end-of-day send.

    ticketTotal was tried here too but dropped: the Digii ticket export's
    file size swings wildly between hourly fetches (26 KB, then 490 KB, then
    5.8 KB, ...) seemingly due to how the vendor's export/attachment process
    works, not real ticket volume — so ticketTotal jumped by thousands
    between consecutive runs and fired an email almost every hour even
    though nothing meaningful had changed.
    """
    days = data.get("days") or []
    return {
        "maxDay": max(days) if days else None,
        "dayCount": len(days),
    }


def _data_fingerprint(data: dict) -> str:
    """
    Stable hash of the SIGNIFICANT subset of parsed source data (see
    _significant_signal) — changes only when a new day's data actually
    lands, not on every minor edit elsewhere in the source file.
    """
    import hashlib
    import json
    signal = _significant_signal(data)
    return hashlib.sha256(json.dumps(signal, sort_keys=True, default=str).encode("utf-8")).hexdigest()


def _load_report_sent_state() -> dict:
    """
    {'date': 'YYYY-MM-DD', 'fingerprint': '...', 'guaranteedSent': [11, 20]}
    of the last successful send, or {}. guaranteedSent tracks which of
    GUARANTEED_SEND_HOURS have already sent the full dashboard TODAY — each
    guaranteed hour is independent and unconditional, so 8 PM still sends
    its own copy even though 11 AM already sent one earlier the same day.
    """
    import json
    if not LAST_REPORT_SENT_FILE.exists():
        return {}
    try:
        return json.loads(LAST_REPORT_SENT_FILE.read_text(encoding="utf-8"))
    except (ValueError, OSError):
        return {}


def _should_send_report(data: dict) -> tuple:
    """
    Decide whether this run should email the full dashboard HTML, and why
    (for logging). Returns (action, reason) where action is either
    "dashboard" or None:
      "dashboard" — send the full dashboard HTML. Fires when either the data
          changed since the last send (a new day's electrical/EB&DG data
          landing mid-day), or the current hour is one of
          GUARANTEED_SEND_HOURS and that specific slot hasn't already sent
          today — this fires unconditionally, even if nothing changed, so
          11 AM and 8 PM always deliver a real dashboard copy.
      None        — nothing to send (already covered by an earlier run this
          hour, or outside a guaranteed hour with no data change).
    """
    state = _load_report_sent_state()
    today = datetime.now().strftime("%Y-%m-%d")
    fingerprint = _data_fingerprint(data)

    if fingerprint != state.get("fingerprint"):
        return "dashboard", "data changed since last send"

    guaranteed_sent = state.get("guaranteedSent", []) if state.get("date") == today else []
    now_hour = datetime.now().hour
    for slot in GUARANTEED_SEND_HOURS:
        if now_hour == slot and slot not in guaranteed_sent:
            return "dashboard", f"guaranteed send for {slot}:00 (not yet sent this slot today)"

    return None, "no data change and no guaranteed slot due right now"


def _mark_report_sent(data: dict, slot_hour: int = None) -> None:
    import os
    import json
    today = datetime.now().strftime("%Y-%m-%d")
    prev = _load_report_sent_state()
    guaranteed_sent = prev.get("guaranteedSent", []) if prev.get("date") == today else []
    if slot_hour is not None and slot_hour not in guaranteed_sent:
        guaranteed_sent.append(slot_hour)
    state = {
        "date": today,
        "time": datetime.now().strftime("%I:%M %p"),
        "fingerprint": _data_fingerprint(data),
        "guaranteedSent": guaranteed_sent,
    }
    tmp = LAST_REPORT_SENT_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(state), encoding="utf-8")
    os.replace(str(tmp), str(LAST_REPORT_SENT_FILE))


def _send_alert_email(subject: str, body: str) -> None:
    """
    Shared alert-email sender, reused by both crash alerts and freshness-check
    alerts. This is the pipeline's only notification channel — there is no
    SMTP/webhook fallback — so if it can't send (most likely because the same
    Gmail token that broke the pipeline is also needed to send the alert
    about it), the alert is written to a fallback file on disk instead of
    just a log line, so it's discoverable without anyone reading pipeline.log.
    """
    # Imported in its own try, separate from the imports below — those can
    # fail for unrelated reasons (e.g. a missing 'google' package on a fresh
    # host), and if AuthNeedsHumanError were imported in that same block, an
    # unrelated import failure would leave the name unbound, crashing the
    # `except AuthNeedsHumanError` clause below with an UnboundLocalError
    # instead of falling through to the generic handler as intended.
    try:
        from gmail_fetcher import AuthNeedsHumanError
    except Exception:
        AuthNeedsHumanError = None

    try:
        from send_report import get_gmail_service, FROM_EMAIL, TO_EMAILS
        from email.mime.text import MIMEText
        import base64

        msg = MIMEText(body, 'plain')
        msg['Subject'] = subject
        msg['From']    = FROM_EMAIL
        msg['To']      = ', '.join(TO_EMAILS)

        service = get_gmail_service()
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId='me', body={'raw': raw}).execute()
        log.info("Alert emailed to: %s", ', '.join(TO_EMAILS))
        # A successful send means the token is healthy again — clear any
        # stale fallback so it doesn't look like an unresolved alert forever.
        if ALERT_FALLBACK_FILE.exists():
            ALERT_FALLBACK_FILE.unlink()
    except Exception as e:
        if AuthNeedsHumanError is not None and isinstance(e, AuthNeedsHumanError):
            log.error("Could not send alert email — Gmail auth itself needs a human: %s", e)
            _write_alert_fallback(subject, body, reason=str(e))
        else:
            log.error("Failed to send alert email (non-fatal):\n%s", traceback.format_exc())
            _write_alert_fallback(subject, body, reason=traceback.format_exc())


def _write_alert_fallback(subject: str, body: str, reason: str) -> None:
    """Last-resort, file-based record of an alert that couldn't be emailed."""
    try:
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        ALERT_FALLBACK_FILE.write_text(
            f"[{now}] ALERT COULD NOT BE EMAILED\n\n"
            f"Subject: {subject}\n\n{body}\n\n"
            f"---\nWhy the email failed to send:\n{reason}\n",
            encoding="utf-8",
        )
        log.error("Alert written to fallback file instead: %s", ALERT_FALLBACK_FILE)
    except Exception:
        log.error("Could not even write the alert fallback file:\n%s", traceback.format_exc())


def send_failure_alert(error_text: str) -> None:
    """
    Notify by email when the pipeline crashes. Task Scheduler runs this
    unattended, so an unhandled exception here would otherwise only ever
    show up in pipeline.log — silently going stale until someone happens
    to check. This is the tripwire for that.

    An "AuthNeedsHumanError" in the traceback gets a distinct subject line
    so it doesn't get triaged the same way as an ordinary bug — it means
    someone needs to physically re-authorize Gmail, not debug code.
    """
    now = datetime.now().strftime('%d %b %Y, %I:%M %p')
    is_auth_issue = "AuthNeedsHumanError" in error_text
    subject = (
        'GL Dashboard — ACTION NEEDED: Gmail re-authorization required'
        if is_auth_issue else
        'GL Dashboard — Pipeline FAILED'
    )
    _send_alert_email(
        subject,
        f"The GL Campus Intelligence pipeline failed at {now}.\n\n"
        f"The dashboard was NOT updated this run — it is showing stale data "
        f"until this is fixed and the pipeline re-run succeeds.\n\n"
        + ("This specific failure means the Gmail OAuth token needs a human to "
           "re-authorize it (run gmail_fetcher.bootstrap_token() on a machine "
           "with a browser, then redeploy the resulting token.json).\n\n"
           if is_auth_issue else "")
        + f"Error:\n{error_text}\n\n"
        f"Full log: {LOG_FILE}"
    )


def verify_dashboard_freshness(data: dict, out_path: str) -> list[str]:
    """
    Catch the "pipeline exited cleanly but the dashboard is actually wrong/stale"
    class of bug — the kind that doesn't crash, so send_failure_alert never fires,
    but still leaves you looking at outdated numbers. Returns a list of problem
    descriptions (empty list = all clear). This is a sanity check on the OUTPUT,
    not on whether the code ran without exceptions.
    """
    problems = []
    path = Path(out_path)

    # 1. The file must have actually been touched in the last few minutes.
    if not path.exists():
        problems.append(f"Dashboard file does not exist at {path}")
        return problems
    age_seconds = datetime.now().timestamp() - path.stat().st_mtime
    if age_seconds > 300:
        problems.append(f"Dashboard file was NOT updated just now (last modified {age_seconds/60:.1f} min ago)")

    html = path.read_text(encoding="utf-8")

    # 2. Day coverage should roughly match how far into the month we actually are.
    #    A big gap (e.g. today is the 6th but only 2 days of data) means a stale
    #    cached Excel file was used instead of the latest email attachment.
    today = datetime.now()
    if data.get("month") == today.month and data.get("year") == today.year:
        n_days = len(data.get("days", []) or [])
        active_days = sum(1 for v in (data.get("junEB") or []) if v)
        if active_days and active_days < today.day - 2:
            problems.append(
                f"Only {active_days} active day(s) of data for {today.strftime('%B')} "
                f"but today is day {today.day} — likely an outdated source file was used."
            )

    # 3. Injection placeholders that should have been filled must not still be empty.
    for var_name, label in [
        ("CHAIR_COUNTS", "Chair Count"),
        ("BUILDINGS", "Buildings"),
        ("REPORT_ARCHIVE", "Report Archive"),
    ]:
        m = re.search(rf"const\s+{var_name}\s*=\s*(\{{\}}|\[\]);", html)
        if m:
            problems.append(f"{label} section is still empty ({var_name}={m.group(1)}) — injection may have silently failed")

    # 4. Any parser section that failed this run (see excel_parser._safe_section) —
    #    a renamed/reworked sheet degrades gracefully instead of crashing the
    #    whole run, but that degradation must still be visible to a human,
    #    not just a line in pipeline.log nobody is watching.
    for err in (data.get("sectionErrors") or []):
        problems.append(f"A dashboard section failed to parse this run and was skipped: {err}")

    return problems


def send_freshness_alert(problems: list[str]) -> None:
    now = datetime.now().strftime('%d %b %Y, %I:%M %p')
    _send_alert_email(
        'GL Dashboard — Data Looks Stale or Incomplete',
        f"The pipeline ran successfully at {now} (no crash), but an automatic "
        f"check found the dashboard may not reflect today's real data:\n\n"
        + "\n".join(f"  - {p}" for p in problems)
        + f"\n\nThe dashboard file was still updated — please double-check it before "
        f"relying on it. Full log: {LOG_FILE}"
    )


def send_vercel_deploy_confirmation(production_url: str) -> None:
    """
    Short confirmation that Step 5's Vercel redeploy succeeded and the
    public link now reflects this run's data — sent on every successful
    deploy (not just once/day), since a deploy failing silently would
    otherwise leave the public URL stale with no visible signal at all.
    """
    now = datetime.now().strftime('%d %b %Y, %I:%M %p')
    _send_alert_email(
        'GL Campus Intelligence — Public Dashboard Redeployed',
        f"The public dashboard was successfully redeployed at {now} with "
        f"this run's latest data.\n\n"
        f"Live link: {production_url}\n\n"
        f"Full log: {LOG_FILE}"
    )


# ── Pipeline steps ────────────────────────────────────────────────────────────

def step1_fetch_emails() -> list[str]:
    """Phase 1: Fetch Excel attachments from Gmail."""
    log.info("STEP 1 — Gmail Fetch")
    from gmail_fetcher import run as fetch_run
    files = fetch_run()
    if not files:
        log.info("  No new files downloaded — pipeline complete (nothing to update).")
    else:
        log.info("  %d file(s) fetched:", len(files))
        for f in files:
            log.info("    %s", Path(f).name)
    return files or []


def step2_parse_excel(file_path: str) -> dict:
    """Phase 2a: Parse the Electrical/STP/AC Excel file."""
    log.info("STEP 2a — Electrical Parse: %s", Path(file_path).name)
    from excel_parser import parse_electrical_file
    data = parse_electrical_file(file_path)
    log.info(
        "  Parsed %d days of data for %d/%d",
        len(data.get("days", [])),
        data.get("month"),
        data.get("year"),
    )
    return data


def step2b_parse_wtp(file_path: str) -> dict:
    """Phase 2b: Parse the WTP water Excel file."""
    log.info("STEP 2b — WTP Parse: %s", Path(file_path).name)
    from excel_parser import parse_wtp_file
    data = parse_wtp_file(file_path)
    log.info(
        "  WTP parsed %d days for %d/%d",
        len(data.get("wtpDays", [])),
        data.get("month"),
        data.get("year"),
    )
    return data


def step2c_parse_cblock(file_path: str) -> dict:
    """Phase 2c: Parse the C-Block / Bramahputra panel reading Excel file."""
    log.info("STEP 2c — C-Block Parse: %s", Path(file_path).name)
    from excel_parser import parse_cblock_file
    data = parse_cblock_file(file_path)
    log.info(
        "  C-Block parsed for %s/%s: %s",
        data.get("month"), data.get("year"),
        ", ".join(data.get("buildings", {}).keys()) or "none",
    )
    return data


def step2d_parse_chair_count(file_path: str) -> dict:
    """Phase 2d: Parse the Chair Count Details Excel file."""
    log.info("STEP 2d — Chair Count Parse: %s", Path(file_path).name)
    from excel_parser import parse_chair_count_file
    data = parse_chair_count_file(file_path)
    log.info(
        "  Chair Count parsed: %s",
        ", ".join(data.get("chairCounts", {}).keys()) or "none",
    )
    return data


def step2e_parse_tickets(file_path: str) -> dict:
    """Phase 2e: Parse the Digii Tickets (CHC_Service_Report) Excel file."""
    log.info("STEP 2e — Tickets Parse: %s", Path(file_path).name)
    from excel_parser import parse_ticket_file
    data = parse_ticket_file(file_path)
    log.info("  Tickets parsed: %d total", data.get("ticketTotal", 0))
    return data


def step3_inject_dashboard(data: dict) -> str:
    """Phase 3: Inject parsed data into the HTML dashboard."""
    log.info("STEP 3 — Dashboard Inject")
    from dashboard_injector import inject
    out_path = inject(data)
    log.info("  Dashboard updated: %s", Path(out_path).name)
    return out_path


# ── Pick the most relevant Excel file ─────────────────────────────────────────

def pick_file(files: list[str], keywords: list[str]) -> str | None:
    """Return the most recently downloaded file matching any keyword (last = newest email)."""
    matches = [f for f in files if any(kw in Path(f).name.lower() for kw in keywords)]
    return matches[-1] if matches else None


def _pick_most_recent(files: list[str], keywords: list[str], label: str) -> str | None:
    """
    Return the most recently MODIFIED file matching any keyword — never the
    largest. File size does not correlate with recency (an older report with
    more sheets/history can outweigh a fresh but smaller one), which silently
    regressed the dashboard to stale data in the past. Modification time is
    the only reliable signal for "which file is actually current."
    """
    matches = [f for f in files if any(kw in Path(f).name.lower() for kw in keywords)]
    if matches:
        return max(matches, key=lambda f: Path(f).stat().st_mtime)

    download_dir = BASE_DIR / "downloads"
    xlsx_files = list(download_dir.glob("*.xlsx"))
    cached = [f for f in xlsx_files if any(kw in f.name.lower() for kw in keywords)]
    if cached:
        best = max(cached, key=lambda p: p.stat().st_mtime)
        log.warning("  No %s report in this run's fetch — using most recently modified cached file: %s",
                    label, best.name)
        return str(best)

    return None


def pick_electrical_file(files: list[str]) -> str | None:
    """Return the most recently modified Electrical/STP/AC report."""
    return _pick_most_recent(files, ["electrical", "ac", "carpentry", "stp", "plumbing"], "electrical")


def pick_cblock_file(files: list[str]) -> str | None:
    """Return the most recently modified C-Block/Bramahputra panel reading Excel."""
    return _pick_most_recent(files, ["c block", "c-block", "bramahputra", "bramhaputra"], "C-Block")


def pick_chair_count_file(files: list[str]) -> str | None:
    """Return the most recently modified Chair Count Details Excel."""
    return _pick_most_recent(files, ["chair count", "chair"], "Chair Count")


def pick_wtp_file(files: list[str]) -> str | None:
    """Return the most recently modified WTP water Excel."""
    return _pick_most_recent(files, ["wtp", "water treatment", "water reading", "ro,canteen", "ro water"], "WTP")


TICKET_CACHE_MAX_AGE_DAYS = 7  # matches search_emails' fetch cadence — see pick_ticket_file


def pick_ticket_file(files: list[str]) -> str | None:
    """
    Return the Digii Tickets (CHC_Service_Report) Excel with the most rows of
    history. Sent daily from either kalaimughilan@greatlakes.edu.in or the
    itsupport@greatlakes.edu.in shared inbox — same export, either sender.
    Whichever day's sender emails twice, the earlier send is a small, partial
    export — the later/larger one is the real full history, and file size
    (not recency) is the correct signal here since both arrive same-day.

    The cache-fallback path below is a DIFFERENT situation: downloads/ is
    never cleaned up, so it can hold files spanning months. Size alone is
    the wrong signal there — an old export that accumulated more history by
    the time it was sent can out-size a smaller recent one, silently
    injecting stale ticket data with nothing to catch it. So the fallback
    only considers files downloaded within TICKET_CACHE_MAX_AGE_DAYS, and
    refuses (rather than guessing) if nothing that recent exists.
    """
    keywords = ["chc_service_report", "digii ticket", "service_request", "service request"]
    matches = [f for f in files if any(kw in Path(f).name.lower() for kw in keywords)]
    if matches:
        return max(matches, key=lambda f: Path(f).stat().st_size)

    download_dir = BASE_DIR / "downloads"
    cached = [f for f in download_dir.glob("*.xlsx") if any(kw in f.name.lower() for kw in keywords)]
    cutoff = datetime.now().timestamp() - TICKET_CACHE_MAX_AGE_DAYS * 86400
    recent_cached = [f for f in cached if f.stat().st_mtime >= cutoff]
    if recent_cached:
        best = max(recent_cached, key=lambda p: p.stat().st_size)
        log.warning("  No Tickets report in this run's fetch — using largest cached file "
                    "from the last %d day(s): %s", TICKET_CACHE_MAX_AGE_DAYS, best.name)
        return str(best)

    if cached:
        newest_age_days = (datetime.now().timestamp() - max(f.stat().st_mtime for f in cached)) / 86400
        log.warning("  No Tickets report in this run's fetch, and the newest cached ticket file "
                    "is %.0f day(s) old (older than the %d-day freshness window) — refusing to "
                    "inject stale ticket data. Tickets section will not be updated this run.",
                    newest_age_days, TICKET_CACHE_MAX_AGE_DAYS)

    return None


# ── Main pipeline ─────────────────────────────────────────────────────────────

def _run_section(label: str, fn):
    """
    Run one file's parse step in isolation. Without this, a crash in ANY one
    of the 5 source files (e.g. a renamed sheet, an unparseable date column)
    propagates to the pipeline's outer try/except and aborts step3_inject_
    dashboard entirely — discarding every OTHER file's data that parsed fine
    in the same run, not just the one that failed. Returns the parsed dict,
    or None (+ a logged warning) if this section's file couldn't be parsed —
    the caller must treat None as "skip this section, keep going."
    """
    try:
        return fn()
    except Exception:
        log.warning("%s failed to parse — that section will not be updated this run "
                    "(other sections continue normally):\n%s", label, traceback.format_exc())
        return None


def run_pipeline(force_file: str = None, send_report_email: bool = False, deploy_to_vercel: bool = False) -> bool:
    """
    Run the full pipeline.
    force_file: skip email fetch and use this Excel path directly (useful for testing).
    send_report_email: if True, sends the monthly report email after injecting the dashboard.
    deploy_to_vercel: if True, pushes the updated dashboard HTML to Vercel for the public link
        after a successful injection. Requires the Vercel CLI installed + logged in on this
        machine (see README.md's "Public hosting (Vercel)" section). Failure is non-fatal.
    Returns True on success, False on any error.
    """
    start = datetime.now()
    log.info(SEPARATOR)
    log.info("GL Dashboard — Pipeline started at %s", start.strftime("%Y-%m-%d %H:%M:%S"))
    log.info(SEPARATOR)

    try:
        # ── Step 1: Gmail fetch (skip if force_file provided) ─────────────────
        if force_file:
            log.info("STEP 1 — Skipped (using provided file: %s)", Path(force_file).name)
            elec_file   = force_file
            wtp_file    = pick_wtp_file([])        # check downloads dir only
            cblock_file = pick_cblock_file([])     # check downloads dir only
            chair_file  = pick_chair_count_file([])  # check downloads dir only
            ticket_file = pick_ticket_file([])       # check downloads dir only
        else:
            fetched = step1_fetch_emails()
            log.info(DIVIDER)

            elec_file   = pick_electrical_file(fetched)
            wtp_file    = pick_wtp_file(fetched)
            cblock_file = pick_cblock_file(fetched)
            chair_file  = pick_chair_count_file(fetched)
            ticket_file = pick_ticket_file(fetched)

            if not elec_file and not wtp_file and not cblock_file and not chair_file and not ticket_file:
                log.info("No Excel file to process. Pipeline done.")
                return True

        # ── Step 2a: Parse Electrical Excel ───────────────────────────────────
        data = {}
        if elec_file:
            log.info(DIVIDER)
            data = _run_section("Electrical", lambda: step2_parse_excel(elec_file)) or {}

        # ── Step 2b: Parse WTP Excel and merge ────────────────────────────────
        if wtp_file:
            log.info(DIVIDER)
            wtp_data = _run_section("WTP", lambda: step2b_parse_wtp(wtp_file))
            if wtp_data:
                elec_unmatched = data.get("unmatchedSheets", [])
                data.update(wtp_data)  # merge all WTP fields including new water arrays
                data["unmatchedSheets"] = elec_unmatched + wtp_data.get("unmatchedSheets", [])
                if not data.get("month"):
                    data["month"] = wtp_data.get("month")
                    data["year"]  = wtp_data.get("year")
        else:
            log.warning("WTP file not found — water section will not be updated this run")

        # ── Step 2c: Parse C-Block/Bramahputra Excel and merge ─────────────────
        if cblock_file:
            log.info(DIVIDER)
            cblock_data = _run_section("C-Block", lambda: step2c_parse_cblock(cblock_file))
            if cblock_data:
                # Merge buildings dicts rather than overwrite — electrical file and
                # C-Block file each contribute distinct building keys.
                data.setdefault("buildings", {}).update(cblock_data.get("buildings", {}))
                data["unmatchedSheets"] = data.get("unmatchedSheets", []) + cblock_data.get("unmatchedSheets", [])
                if not data.get("month"):
                    data["month"] = cblock_data.get("month")
                    data["year"]  = cblock_data.get("year")
        else:
            log.warning("C-Block file not found — C-Block/Bramahputra section will not be updated this run")

        # ── Step 2d: Parse Chair Count Excel and merge ─────────────────────────
        if chair_file:
            log.info(DIVIDER)
            chair_data = _run_section("Chair Count", lambda: step2d_parse_chair_count(chair_file))
            if chair_data:
                data["chairCounts"] = chair_data.get("chairCounts", {})
                if not data.get("month"):
                    # Chair Count sheets don't carry a single "report month" the way
                    # electrical/WTP do (inventory, not a period total) — fall back
                    # to today's date only if nothing else has set month/year yet.
                    today = datetime.now()
                    data["month"] = today.month
                    data["year"]  = today.year
        else:
            log.warning("Chair Count file not found — Facilities/Chair Count section will not be updated this run")

        # ── Step 2e: Parse Digii Tickets Excel and merge ────────────────────────
        if ticket_file:
            log.info(DIVIDER)
            ticket_data = _run_section("Tickets", lambda: step2e_parse_tickets(ticket_file))
            if ticket_data:
                for key in ["ticketTotal", "ticketPending", "ticketAvgTatMin",
                            "ticketsByDept", "ticketsByLevel", "ticketsDeptByLevel", "ticketRows",
                            "ticketRecurring", "ticketRecurringByDept",
                            "ticketMonths", "ticketThisMonth", "ticketThisMonthKey"]:
                    data[key] = ticket_data.get(key)
                data["sectionErrors"] = data.get("sectionErrors", []) + ticket_data.get("ticketSectionErrors", [])
                if not data.get("month"):
                    today = datetime.now()
                    data["month"] = today.month
                    data["year"]  = today.year
        else:
            log.warning("Tickets file not found — Tickets section will not be updated this run")

        # ── Step 2f: Monthly Report cutover check ───────────────────────────────
        # Must run BEFORE injection — once the new month's daily arrays are
        # written in, the previous month's report can no longer be rebuilt.
        if data.get("month") and data.get("year"):
            log.info(DIVIDER)
            log.info("STEP 2f — Report Cutover Check")
            try:
                from report_archiver import check_and_rollover
                check_and_rollover(data["month"], data["year"], datetime.now())
            except Exception:
                log.warning("Report cutover check failed (non-fatal):\n%s", traceback.format_exc())

        # ── Step 3: Inject dashboard ──────────────────────────────────────────
        # inject() requires the electrical backbone (days/month/year/
        # rmSolarDaily) unconditionally — if that section failed to parse
        # above (data == {}) there is nothing safe to inject this run. Skip
        # rather than let inject() KeyError on missing required fields, which
        # would look like a code bug rather than "one file didn't parse."
        if not (data.get("days") and data.get("month") and data.get("year")):
            log.warning("No usable electrical/day data this run (parse failed or file missing) — "
                        "skipping dashboard injection entirely; dashboard remains on last good data.")
            log.info(SEPARATOR)
            log.info("Pipeline COMPLETE (no injection) in %.1fs", (datetime.now() - start).total_seconds())
            log.info(SEPARATOR)
            return True

        log.info(DIVIDER)
        out_path = step3_inject_dashboard(data)

        # ── Step 3b: Freshness check — catch silent "ran fine but wrong" bugs ──
        log.info(DIVIDER)
        log.info("STEP 3b — Dashboard Freshness Check")
        problems = verify_dashboard_freshness(data, out_path)
        unmatched = data.get("unmatchedSheets", [])
        if unmatched:
            problems.append(
                f"Excel file has {len(unmatched)} sheet(s) not recognized by any parser "
                f"(new or renamed sheet, data NOT extracted): {', '.join(unmatched)}"
            )
        if problems:
            for p in problems:
                log.warning("  FRESHNESS CHECK: %s", p)
            send_freshness_alert(problems)
        else:
            log.info("  Freshness check passed — dashboard reflects current data.")

        # ── Step 4: Send dashboard email (optional — on data change, or ─────────
        # unconditionally at each hour in GUARANTEED_SEND_HOURS) ────────────────
        if send_report_email:
            action, reason = _should_send_report(data)
            if action is None:
                log.info("Dashboard email skipped — %s.", reason)
            else:
                log.info(DIVIDER)
                log.info("STEP 4 — Dashboard Email (%s)", reason)
                try:
                    from send_report import load_dashboard_data, compute_insights, build_pdf, send_report as _send
                    d   = load_dashboard_data()
                    ins = compute_insights(d)
                    pdf_bytes = build_pdf(ins)
                    _send(pdf_bytes, None, ins['month'])
                    now_hour = datetime.now().hour
                    slot_hour = now_hour if now_hour in GUARANTEED_SEND_HOURS else None
                    _mark_report_sent(data, slot_hour=slot_hour)
                except Exception:
                    log.warning("Dashboard email failed (non-fatal):\n%s", traceback.format_exc())

        # ── Step 5: Deploy dashboard to Vercel (optional public link) ──────────
        # Non-fatal by design: a failed deploy just means the public link is
        # stale until next run, not that this run's local dashboard update
        # (already written above) is invalid.
        if deploy_to_vercel:
            log.info(DIVIDER)
            log.info("STEP 5 — Vercel Deploy")
            try:
                from vercel_deploy import deploy as _deploy_vercel, PRODUCTION_URL
                ok, _deployed_url = _deploy_vercel(Path(out_path))
                if not ok:
                    log.warning("Vercel deploy did not succeed (non-fatal) — see log above.")
                else:
                    try:
                        send_vercel_deploy_confirmation(PRODUCTION_URL)
                    except Exception:
                        log.warning("Vercel deploy confirmation email failed (non-fatal):\n%s", traceback.format_exc())
            except Exception:
                log.warning("Vercel deploy failed (non-fatal):\n%s", traceback.format_exc())

        elapsed = (datetime.now() - start).total_seconds()
        log.info(SEPARATOR)
        log.info("Pipeline COMPLETE in %.1fs", elapsed)
        log.info(SEPARATOR)
        return True

    except Exception:
        err = traceback.format_exc()
        log.error("Pipeline FAILED:")
        log.error(err)
        send_failure_alert(err)
        log.info(SEPARATOR)
        return False


# ── CLI entry point ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    setup_logging()

    # Dashboard email is sent by default (once/day, guarded by
    # _report_already_sent_today) so the scheduled task doesn't need any
    # special flag. Pass --no-send-report to disable it for a one-off/test run.
    #
    # Usage: python run_pipeline.py
    #   or:  python run_pipeline.py "E:\GLIM\downloads\electrical.xlsx"
    #   or:  python run_pipeline.py --no-send-report
    #   or:  python run_pipeline.py --deploy
    force       = next((a for a in sys.argv[1:] if not a.startswith('--')), None)
    send_report = '--no-send-report' not in sys.argv
    deploy      = '--deploy' in sys.argv

    success = run_pipeline(force_file=force, send_report_email=send_report, deploy_to_vercel=deploy)
    sys.exit(0 if success else 1)
