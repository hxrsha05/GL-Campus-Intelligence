"""
Phase 5 — Monthly Report Archiver
Handles the "Monthly Report" cutover: the live report stays on a given
month through the 15th of the following month, then freezes a PDF/HTML
snapshot of it into archive/ and hands the live slot to the new month.

This must run BEFORE the pipeline injects the new month's data into the
dashboard, since injection overwrites the daily arrays the old month's
report is built from — there is no "look back" once that happens.
"""

import json
import logging
import re
from pathlib import Path
from datetime import datetime

BASE_DIR     = Path(__file__).parent
ARCHIVE_DIR  = BASE_DIR / "archive"
INDEX_FILE   = ARCHIVE_DIR / "report_archive.json"
DASHBOARD_FILE = BASE_DIR / "GL_Dashboard_v4_July2026.html"

MONTH_NAMES = {
    1:"January",2:"February",3:"March",4:"April",5:"May",6:"June",
    7:"July",8:"August",9:"September",10:"October",11:"November",12:"December"
}

log = logging.getLogger(__name__)


def _read_dashboard_month_year(html: str) -> tuple[int, int] | None:
    """Read the month/year the dashboard is CURRENTLY showing, before this run's injection."""
    days_m = re.search(r"const JUN_DAYS\s*=\s*\[\s*'?\d+-([A-Za-z]{3})'", html)
    yr_m   = re.search(r"const\s+CUR_YEAR\s*=\s*(\d+);", html)
    if not days_m or not yr_m:
        return None
    abbr_to_num = {v.lower(): k for k, v in
                   {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
                    7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}.items()}
    mon_num = abbr_to_num.get(days_m.group(1).lower())
    if not mon_num:
        return None
    return mon_num, int(yr_m.group(1))


def _load_index() -> list:
    if INDEX_FILE.exists():
        return json.loads(INDEX_FILE.read_text(encoding="utf-8"))
    return []


def _save_index(entries: list) -> None:
    ARCHIVE_DIR.mkdir(exist_ok=True)
    INDEX_FILE.write_text(json.dumps(entries, indent=2), encoding="utf-8")


def archive_current_report(month_num: int, year: int) -> dict | None:
    """
    Freeze the dashboard's CURRENT live report (before this run's new-month
    injection happens) into a static PDF, and record it in the archive index.
    Returns the archive entry, or None on failure.
    """
    try:
        from send_report import load_dashboard_data, compute_insights, build_pdf
        data = load_dashboard_data()
        ins  = compute_insights(data)
        pdf_bytes = build_pdf(ins)
    except Exception:
        log.exception("Failed to build archive snapshot for %s %d", MONTH_NAMES.get(month_num), year)
        return None

    ARCHIVE_DIR.mkdir(exist_ok=True)
    abbr = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
            7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}[month_num]
    pdf_name = f"GL_Campus_Report_{abbr}_{year}.pdf"
    pdf_path = ARCHIVE_DIR / pdf_name
    pdf_path.write_bytes(pdf_bytes)

    entry = {
        "month": month_num,
        "year": year,
        "label": f"{MONTH_NAMES[month_num]} {year}",
        "abbr": abbr,
        "file": pdf_name,
        "archivedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "placeholder": False,
    }

    entries = _load_index()
    entries = [e for e in entries if not (e["month"] == month_num and e["year"] == year)]
    entries.append(entry)
    entries.sort(key=lambda e: (e["year"], e["month"]))
    _save_index(entries)

    log.info("Archived report: %s -> %s", entry["label"], pdf_name)
    return entry


def add_placeholder_months(year: int, up_to_month_exclusive: int) -> None:
    """
    Register "name sake" placeholder entries for months before real archiving
    began (e.g. Jan-May 2026) — no PDF, just a name/date shown as archived
    with no downloadable data, since no daily source data exists for them.
    """
    entries = _load_index()
    existing = {(e["month"], e["year"]) for e in entries}
    changed = False
    for m in range(1, up_to_month_exclusive):
        if (m, year) in existing:
            continue
        entries.append({
            "month": m, "year": year,
            "label": f"{MONTH_NAMES[m]} {year}",
            "abbr": {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
                     7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}[m],
            "file": None,
            "archivedAt": None,
            "placeholder": True,
        })
        changed = True
    if changed:
        entries.sort(key=lambda e: (e["year"], e["month"]))
        _save_index(entries)
        log.info("Registered placeholder report entries for %s %d through month %d",
                  year, year, up_to_month_exclusive - 1)


def check_and_rollover(new_month_num: int, new_year: int, today: datetime) -> None:
    """
    Core cutover rule: the live Monthly Report stays on a given month through
    day 15 of the FOLLOWING month. On day 16+ of the new month, if the
    dashboard is still showing the previous month, freeze it into the
    archive and let this run's injection promote the new month to live.

    Must be called BEFORE dashboard_injector.inject() overwrites the old
    month's daily arrays.
    """
    if not DASHBOARD_FILE.exists():
        return
    html = DASHBOARD_FILE.read_text(encoding="utf-8")
    current = _read_dashboard_month_year(html)
    if current is None:
        return
    cur_month, cur_year = current

    if (cur_month, cur_year) == (new_month_num, new_year):
        return  # same month already live — nothing to roll over

    is_next_month = (new_year == cur_year and new_month_num == cur_month + 1) or \
                     (new_year == cur_year + 1 and cur_month == 12 and new_month_num == 1)
    if not is_next_month:
        # Not a simple one-month step forward (e.g. backfill run) — skip
        # the automatic rollover rather than guess; archive manually if needed.
        log.info("Dashboard shows %s %d but new data is %s %d — not a direct month "
                  "step, skipping automatic archive.",
                  MONTH_NAMES.get(cur_month), cur_year, MONTH_NAMES.get(new_month_num), new_year)
        return

    if today.day < 16:
        log.info("New data is for %s %d but today is day %d of %s — keeping %s %d live "
                  "until day 16 (per report cutover rule).",
                  MONTH_NAMES.get(new_month_num), new_year, today.day,
                  MONTH_NAMES.get(new_month_num), MONTH_NAMES.get(cur_month), cur_year)
        return

    log.info("Day %d of %s reached — archiving %s %d and promoting %s %d to live Monthly Report.",
              today.day, MONTH_NAMES.get(new_month_num),
              MONTH_NAMES.get(cur_month), cur_year, MONTH_NAMES.get(new_month_num), new_year)
    archive_current_report(cur_month, cur_year)
