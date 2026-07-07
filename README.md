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
- Daily/weekly automated email delivery of the generated report
- Monthly PDF archival for long-term recordkeeping
- Freshness and anomaly alerts sent automatically if expected data doesn't arrive

## Tech Stack

| Layer | Technology |
|---|---|
| Ingestion | Gmail API, google-auth-oauthlib |
| Parsing | openpyxl |
| Orchestration | Python 3.10+ |
| Dashboard | HTML5, CSS3, Chart.js |
| Reporting | ReportLab (PDF), smtplib/Gmail API (email) |
| Scheduling | Windows Task Scheduler |

## Project Structure

```
GL-Campus-Intelligence/
├── run_pipeline.py         # Orchestrator — chains fetch → parse → inject
├── gmail_fetcher.py        # Downloads report attachments from Gmail
├── excel_parser.py         # Parses electrical, solar, water, building sheets
├── dashboard_injector.py   # Injects parsed data into the dashboard HTML
├── send_report.py          # Emails the generated report
├── report_archiver.py      # Archives monthly PDF snapshots
├── build_day1_docs.py      # Generates project documentation PDFs
├── build_past_report.py    # Backfills historical reports
├── setup_scheduler.ps1     # Registers the daily automation task
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

python run_pipeline.py
```

The first run opens a browser window for Gmail OAuth consent; a `token.json` is cached locally afterward. Neither `credentials.json` nor `token.json` is included in this repo — provide your own.

## Notes

- The generated dashboard (`GL_Dashboard_v4_*.html`) and all source Excel reports are excluded from this repo since they contain real institutional operational data
- Schema drift alerts reuse the existing freshness-check email — no separate alert channel to maintain
- Designed to run unattended via Task Scheduler, with crash and freshness alerts emailed automatically

## Author

Sri Harshavardhan Palaniswamy J
