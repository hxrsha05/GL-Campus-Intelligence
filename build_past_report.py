"""
Phase 5b — Past Month Report Builder
Builds a real PDF report for a historical month (Jan-May 2026) directly
from the source Excel files, bypassing the dashboard HTML entirely (the
dashboard only ever holds the CURRENT month's daily arrays — it cannot be
scraped for past months the way send_report.py's load_dashboard_data() does).

Only includes data that genuinely exists for that month in the source
Excel: Energy (EB/DG/solar/power cuts) has full history back to 2023.
Water/work-orders/manpower/AMC do NOT have 2026 Jan-May history in the
current source files (work order sheets reset monthly, WTP totals/RO
readings only resume from June 2026) — those sections are honestly
omitted rather than shown as zero.

Usage:
    python build_past_report.py 1 2026   # January 2026
    python build_past_report.py 5 2026   # May 2026
"""

import sys
import logging
from pathlib import Path
from datetime import datetime

logging.basicConfig(level=logging.INFO, format="%(levelname)s  %(message)s")
log = logging.getLogger(__name__)

BASE_DIR = Path(__file__).parent
ELECTRICAL_FILE = BASE_DIR / "Electrical, AC, Carpentry, STP & Plumbing Daily Report July 2026.xlsx"
WTP_FILE = BASE_DIR / "downloads" / "WTP 1 to 3, RO,canteen Water reading and new well water reading details July- 2026_20260705_215223.xlsx"

MONTH_NAMES = {1:"January",2:"February",3:"March",4:"April",5:"May",6:"June",
               7:"July",8:"August",9:"September",10:"October",11:"November",12:"December"}
MONTH_ABBR = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
              7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}


def build_report_data(month: int, year: int) -> dict:
    """
    Build the same shaped dict send_report.compute_insights() expects,
    straight from parsed Excel data — no dashboard HTML involved.
    """
    from excel_parser import parse_electrical_file, parse_wtp_file

    elec = parse_electrical_file(str(ELECTRICAL_FILE), target_month=month, target_year=year)
    days = elec.get("days", [])
    day_labels = [f"{d:02d}-{MONTH_ABBR[month]}" for d in days]

    wtp_daily = {}
    laundry_w, new_well = [], []
    ro_cans, tds_ro2000 = [], []
    ac_outlet_total, ac_outlet_blocks = [], []
    if WTP_FILE.exists():
        try:
            wtp = parse_wtp_file(str(WTP_FILE), target_month=month, target_year=year)
            wtp_daily = wtp.get("wtpDaily", {})
            laundry_w = wtp.get("laundryDaily", [])
            new_well  = wtp.get("newWellJunKL", [])
            ro_cans   = wtp.get("roCansDaily", [])
            tds_ro2000 = wtp.get("tdsRo2000", [])
            ac_outlet_total  = wtp.get("acOutletTotal", [])
            ac_outlet_blocks = wtp.get("acOutletBlocks", [])
        except Exception:
            log.warning("WTP data unavailable for %s %d — omitting water section", MONTH_NAMES[month], year)

    return {
        'html': '',  # no dashboard HTML for historical months — regex lookups (work orders, AMC) safely find nothing
        'month': f"{MONTH_ABBR[month]} {year}",  # bare abbr+year, matches live-path convention (see load_dashboard_data)
        'day_labels': day_labels, 'n_days': len(day_labels),
        'eb': elec.get('junEB', []),
        'dg': elec.get('junDG', []),
        'diesel': elec.get('junDiesel', []),
        'stock': elec.get('junStock', []),
        'power_cuts': elec.get('powerCuts', []),
        'ev': elec.get('evDailyKWh', []),
        'nescafe': elec.get('nescafeDailyKWh', []),
        'cvb': elec.get('cvbDailyKWh', []),
        'tea': elec.get('teaDailyKWh', []),
        'yummy': elec.get('yummyDailyKWh', []),
        'laundry_e': elec.get('laundryElecKWh', []),
        'wtp_daily': wtp_daily,
        'laundry_w': laundry_w,
        'new_well': new_well,
        'ac_outlet': ac_outlet_total,
        'ac_blocks': ac_outlet_blocks,
        'ro_cans': ro_cans,
        'tds_ro2000': tds_ro2000,
        'buildings': elec.get('buildings', {}),
    }


def build_past_report_pdf(month: int, year: int) -> bytes:
    from send_report import compute_insights, build_pdf
    data = build_report_data(month, year)
    ins  = compute_insights(data)
    return build_pdf(ins)


def main():
    if len(sys.argv) < 3:
        print("Usage: python build_past_report.py <month> <year>")
        sys.exit(1)
    month, year = int(sys.argv[1]), int(sys.argv[2])

    log.info("Building report for %s %d from source Excel...", MONTH_NAMES[month], year)
    pdf_bytes = build_past_report_pdf(month, year)

    archive_dir = BASE_DIR / "archive"
    archive_dir.mkdir(exist_ok=True)
    out_path = archive_dir / f"GL_Campus_Report_{MONTH_ABBR[month]}_{year}.pdf"
    out_path.write_bytes(pdf_bytes)
    log.info("Saved: %s (%d bytes)", out_path, len(pdf_bytes))

    # Update the archive index
    import json
    index_file = archive_dir / "report_archive.json"
    entries = json.loads(index_file.read_text(encoding="utf-8")) if index_file.exists() else []
    entries = [e for e in entries if not (e["month"] == month and e["year"] == year)]
    entries.append({
        "month": month, "year": year,
        "label": f"{MONTH_NAMES[month]} {year}",
        "abbr": MONTH_ABBR[month],
        "file": out_path.name,
        "archivedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "placeholder": False,
    })
    entries.sort(key=lambda e: (e["year"], e["month"]))
    index_file.write_text(json.dumps(entries, indent=2), encoding="utf-8")
    log.info("Archive index updated.")


if __name__ == "__main__":
    main()
