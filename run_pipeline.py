"""
Phase 4 — Master Pipeline Orchestrator
Chains Gmail fetch → Excel parse → Dashboard inject into a single automated run.
Designed for both local scheduling (Task Scheduler) and cloud deployment (Cloud Run).
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


def _send_alert_email(subject: str, body: str) -> None:
    """Shared alert-email sender, reused by both crash alerts and freshness-check alerts."""
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
    except Exception:
        log.error("Failed to send alert email (non-fatal):\n%s", traceback.format_exc())


def send_failure_alert(error_text: str) -> None:
    """
    Notify by email when the pipeline crashes. Task Scheduler runs this
    unattended, so an unhandled exception here would otherwise only ever
    show up in pipeline.log — silently going stale until someone happens
    to check. This is the tripwire for that.
    """
    now = datetime.now().strftime('%d %b %Y, %I:%M %p')
    _send_alert_email(
        'GL Dashboard — Pipeline FAILED',
        f"The GL Campus Intelligence pipeline failed at {now}.\n\n"
        f"The dashboard was NOT updated this run — it is showing stale data "
        f"until this is fixed and the pipeline re-run succeeds.\n\n"
        f"Error:\n{error_text}\n\n"
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


# ── Main pipeline ─────────────────────────────────────────────────────────────

def run_pipeline(force_file: str = None, send_report_email: bool = False) -> bool:
    """
    Run the full pipeline.
    force_file: skip email fetch and use this Excel path directly (useful for testing).
    send_report_email: if True, sends the monthly report email after injecting the dashboard.
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
        else:
            fetched = step1_fetch_emails()
            log.info(DIVIDER)

            elec_file   = pick_electrical_file(fetched)
            wtp_file    = pick_wtp_file(fetched)
            cblock_file = pick_cblock_file(fetched)
            chair_file  = pick_chair_count_file(fetched)

            if not elec_file and not wtp_file and not cblock_file and not chair_file:
                log.info("No Excel file to process. Pipeline done.")
                return True

        # ── Step 2a: Parse Electrical Excel ───────────────────────────────────
        data = {}
        if elec_file:
            log.info(DIVIDER)
            data = step2_parse_excel(elec_file)

        # ── Step 2b: Parse WTP Excel and merge ────────────────────────────────
        if wtp_file:
            log.info(DIVIDER)
            wtp_data = step2b_parse_wtp(wtp_file)
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
            cblock_data = step2c_parse_cblock(cblock_file)
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
            chair_data = step2d_parse_chair_count(chair_file)
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

        # ── Step 2e: Monthly Report cutover check ───────────────────────────────
        # Must run BEFORE injection — once the new month's daily arrays are
        # written in, the previous month's report can no longer be rebuilt.
        if data.get("month") and data.get("year"):
            log.info(DIVIDER)
            log.info("STEP 2e — Report Cutover Check")
            try:
                from report_archiver import check_and_rollover
                check_and_rollover(data["month"], data["year"], datetime.now())
            except Exception:
                log.warning("Report cutover check failed (non-fatal):\n%s", traceback.format_exc())

        # ── Step 3: Inject dashboard ──────────────────────────────────────────
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

        # ── Step 4: Send report email (optional — only on last day of month) ──
        if send_report_email:
            log.info(DIVIDER)
            log.info("STEP 4 — Report Email")
            try:
                from send_report import load_dashboard_data, compute_insights, build_report_html, send_report as _send
                d   = load_dashboard_data()
                ins = compute_insights(d)
                ins['eb'] = d['eb']
                rpt = build_report_html(ins)
                _send(rpt, None, ins['month'])
            except Exception:
                log.warning("Report email failed (non-fatal):\n%s", traceback.format_exc())

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

    # Usage: python run_pipeline.py
    #   or:  python run_pipeline.py "E:\GLIM\downloads\electrical.xlsx"
    #   or:  python run_pipeline.py --send-report
    force       = next((a for a in sys.argv[1:] if not a.startswith('--')), None)
    send_report = '--send-report' in sys.argv

    success = run_pipeline(force_file=force, send_report_email=send_report)
    sys.exit(0 if success else 1)
