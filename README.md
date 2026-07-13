# GL Campus Intelligence

Turn raw facility Excel reports into a live campus operations dashboard — automatically, with zero manual data entry.

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Gmail API](https://img.shields.io/badge/Gmail%20API-Fetch-EA4335?logo=gmail&logoColor=white)](https://developers.google.com/gmail/api)
[![openpyxl](https://img.shields.io/badge/openpyxl-Excel%20Parsing-217346?logo=microsoft-excel&logoColor=white)](https://openpyxl.readthedocs.io/)
[![Chart.js](https://img.shields.io/badge/Chart.js-Visualization-FF6384?logo=chartdotjs&logoColor=white)](https://www.chartjs.org/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](#license)

## Overview

GL Campus Intelligence is a fully automated reporting pipeline built for Great Lakes Institute of Management. Every day, facility teams email in Excel reports covering electrical load, solar generation, diesel generators, water treatment, canteen utilities, AMC status, and building operations. This pipeline fetches those reports straight from Gmail, parses them, and injects the real numbers into a single-file HTML dashboard — no manual copy-paste, no stale data, no hardcoded placeholders.

```
Gmail Inbox → Excel Fetch → Parse & Validate → Data Injection → Live Dashboard
```

## Features

### Core Pipeline
- **Automated Ingestion** — polls Gmail for the latest facility report attachments and downloads them
- **Multi-Sheet Parsing** — extracts electrical, solar, diesel, water (WTP/RO/canteen), and building panel data from differently-structured Excel workbooks
- **Schema Drift Detection** — flags any sheet in an incoming workbook that no parser recognizes, so silent data loss never goes unnoticed
- **Layout-Change Resilience** — detects mid-sheet header re-declarations (a real vendor quirk) and re-maps columns instead of misattributing data
- **Historical Merge Support** — reconciles renamed meters/panels into one continuous series instead of splitting history

### Dashboard
- Single-file HTML — no build step, no framework, works offline once generated
- Month-over-month and year-over-year comparison widgets that roll forward automatically as new data arrives
- Drill-down views per zone/department/building with live Chart.js visualizations
- Null-vs-zero aware rendering — a genuine zero reading is never confused with missing data

### Reporting
- Dashboard HTML emailed automatically: guaranteed at 11 AM and 8 PM every day regardless of whether anything changed (8 PM exists as a safety net in case the 11 AM email is delayed or missed), plus immediately whenever a genuinely new day of electrical/EB&DG data lands at any other hour
- Monthly PDF archival for long-term recordkeeping
- Freshness and anomaly alerts sent automatically if expected data doesn't arrive

### Public Hosting
- The dashboard is auto-deployed to Vercel after every successful pipeline run, so the public link always reflects the latest data with no manual step
- A short confirmation email is sent on every successful redeploy, linking the live dashboard

## Tech Stack

| Layer | Technology |
|---|---|
| Ingestion | Gmail API, google-auth-oauthlib |
| Parsing | openpyxl |
| Orchestration | Python 3.10+ |
| Dashboard | HTML5, CSS3, Chart.js |
| Reporting | ReportLab (PDF), Gmail API (email — no SMTP fallback) |
| Scheduling | Windows Task Scheduler |
| Public Hosting | Vercel (static deploy via CLI) |

## Project Structure

```
GL-Campus-Intelligence/
├── run_pipeline.py         # Orchestrator — chains fetch → parse → inject
├── gmail_fetcher.py        # Downloads report attachments from Gmail
├── excel_parser.py         # Parses electrical, solar, water, building sheets
├── dashboard_injector.py   # Injects parsed data into the dashboard HTML
├── send_report.py          # Emails the generated dashboard/report
├── report_archiver.py      # Archives monthly PDF snapshots
├── vercel_deploy.py        # Pushes the dashboard to Vercel for the public link
├── build_day1_docs.py      # Generates project documentation PDFs
├── build_past_report.py    # Backfills historical reports
├── setup_scheduler.ps1     # Registers the hourly automation task
└── requirements.txt
```

## Getting Started

### Prerequisites
- Python 3.10+
- A Google Cloud project with the Gmail API enabled
- OAuth credentials (`credentials.json`) placed in the project root

### Setup

```bash
git clone https://github.com/hxrsha05/GL-Campus-Intelligence.git
cd GL-Campus-Intelligence

pip install -r requirements.txt
```

Neither `credentials.json` nor `token.json` is included in this repo — provide your own (see **First-run / new-host bootstrap** below). `run_pipeline.py` itself never opens a browser — on an unattended/headless run, a missing or dead token raises immediately instead of hanging.

### First-run / new-host bootstrap

Before `run_pipeline.py` can succeed even once on a fresh checkout or a new host, three things must exist that are **not** in git (all gitignored on purpose — they're secrets or contain real institutional data):

1. **`credentials.json`** — your Google Cloud OAuth client secret. Download it from the Google Cloud Console for the project with the Gmail API enabled.
2. **`token.json`** — minted by an interactive, one-time authorization on a machine *with a browser*:
   ```bash
   python -c "import gmail_fetcher; gmail_fetcher.bootstrap_token()"
   ```
   This is the *only* place a browser opens in this codebase. Copy the resulting `token.json` to wherever the pipeline will actually run unattended (lab PC, cloud host, container).
3. **`GL_Dashboard_v4_*.html`** — the dashboard template/shell itself. There is no code path that generates this from scratch; `dashboard_injector.py` only ever string-replaces values into an existing file. If you don't have a prior copy, ask the project owner for the current template.

Re-authorization is needed again whenever the refresh token dies (Google can expire it after ~7 days for an OAuth app still in "Testing" publish status, or after 6 months of inactivity, or on manual revocation) — re-run step 2 and redeploy the new `token.json`. The pipeline emails an "ACTION NEEDED: Gmail re-authorization required" alert when this happens; if even the alert can't send (same dead token), it's written to `ALERT_UNDELIVERED.txt` in the project root instead.

### Public hosting (Vercel)

The dashboard can optionally be pushed to a public Vercel URL after every pipeline run. One-time setup on whichever machine will run the deploy:

```bash
npm install -g vercel
vercel login
```

Then pass `--deploy` when running the pipeline (or add it to the scheduled task's arguments):

```bash
python run_pipeline.py --deploy
```

`vercel_deploy.py` copies the current dashboard into `vercel_deploy/index.html` and runs `vercel deploy --prod` from there. The production alias (e.g. `https://gl-campus-intelligence.vercel.app`) always points at the latest deploy, even though each individual deploy also gets its own unique throwaway URL. A failed deploy is non-fatal — it just means the public link is stale until the next successful run.

## Notes

- The generated dashboard (`GL_Dashboard_v4_*.html`) and all source Excel reports are excluded from this repo since they contain real institutional operational data
- Schema drift alerts reuse the existing freshness-check email — no separate alert channel to maintain
- Designed to run unattended via Task Scheduler (hourly, 8 AM–11 PM), with crash and freshness alerts emailed automatically. There is currently no Linux/cloud-cron scheduling script — only `setup_scheduler.ps1` (Windows Task Scheduler)
- A renamed/reworked sheet in a source Excel file degrades gracefully — that one section is skipped and flagged in the freshness-check alert, the rest of the dashboard still updates
- The dashboard email only re-sends mid-day when a genuinely new day of electrical data appears, not on every minor edit elsewhere in the source file — this avoids near-duplicate emails from a source workbook that gets touched frequently during the day

## Author

Sri Harshavardhan Palaniswamy J
