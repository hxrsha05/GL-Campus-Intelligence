"""
GL Campus Intelligence — Monthly Report Sender
Generates a polished PDF report from dashboard data and emails it as an attachment.

Usage:
    python send_report.py                    # send to default recipients
    python send_report.py --dry-run          # generate PDF only, no email
    python send_report.py --to extra@x.com  # add extra recipient
"""

import re
import sys
import logging
import base64
import io
from pathlib import Path
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white, black
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import Flowable

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

BASE_DIR         = Path(__file__).parent
DASHBOARD_FILE   = BASE_DIR / "GL_Dashboard_v4_July2026.html"
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE       = BASE_DIR / "token.json"
SCOPES           = ["https://www.googleapis.com/auth/gmail.send",
                    "https://www.googleapis.com/auth/gmail.modify"]

FROM_EMAIL = "energymonitoring.glc@greatlakes.edu.in"
TO_EMAILS  = [
    "harshavardhan.j@greatlakes.edu.in",
    "maheshkumaar.r@greatlakes.edu.in"
]

logging.basicConfig(level=logging.INFO, format="%(levelname)s  %(message)s")
log = logging.getLogger(__name__)

# ── Colours ───────────────────────────────────────────────────────────────────
BLUE   = HexColor("#1B4FD8")
BLUE2  = HexColor("#3B82F6")
TEAL   = HexColor("#0891B2")
GREEN  = HexColor("#059669")
YELLOW = HexColor("#D97706")
RED    = HexColor("#DC2626")
ORANGE = HexColor("#EA580C")
PURPLE = HexColor("#7C3AED")

INK    = HexColor("#111827")
INK2   = HexColor("#374151")
INK3   = HexColor("#6B7280")
INK4   = HexColor("#9CA3AF")
BG     = HexColor("#F7F8FA")
BORDER = HexColor("#E5E7EB")
SURFACE= white


# ── Data extraction ────────────────────────────────────────────────────────────

def extract_js_array(html, var_name):
    m = re.search(rf'const\s+{re.escape(var_name)}\s*=\s*\[([^\]]*)\]', html)
    if not m: return []
    result = []
    for tok in m.group(1).split(','):
        tok = tok.strip()
        if tok in ('null', ''): result.append(None)
        else:
            try: result.append(float(tok))
            except ValueError: result.append(None)
    return result


def extract_nested_array(html, var_name):
    m = re.search(rf'const\s+{re.escape(var_name)}\s*=\s*\[', html)
    if not m: return []
    start = m.end() - 1
    depth, pos = 0, start
    while pos < len(html):
        if html[pos] == '[': depth += 1
        elif html[pos] == ']':
            depth -= 1
            if depth == 0: break
        pos += 1
    inner = html[start+1:pos]
    result = []
    for sub in re.finditer(r'\[([^\]]*)\]', inner):
        row = []
        for tok in sub.group(1).split(','):
            try: row.append(float(tok.strip()))
            except ValueError: row.append(0.0)
        result.append(row)
    return result


def extract_power_cuts(html):
    pattern = r"\{date:'([^']+)',dur:(\d+),events:'([^']+)',dg:'([^']+)'\}"
    return [{'date': m[0], 'dur': int(m[1]), 'events': m[2], 'dg': m[3]}
            for m in re.findall(pattern, html)]


def extract_wtp_daily(html):
    m = re.search(r'const wtpDaily\s*=\s*\{([\s\S]*?)\};', html)
    if not m: return {}
    block = m.group(1)
    result = {}
    for key in ['wtp1', 'wtp2', 'wtp3']:
        km = re.search(rf'{key}:\s*\[([^\]]*)\]', block)
        if km:
            result[key] = [float(v) for v in km.group(1).split(',') if v.strip()]
    return result


def extract_buildings(html):
    """
    Parse the BUILDINGS const (per-building sub-meter panels) out of the
    dashboard HTML. Returns {key: {'days': [...], 'panels': [{'name','daily'}]}}.
    Uses a balanced-brace scan since the object is arbitrarily nested.
    """
    m = re.search(r'const\s+BUILDINGS\s*=\s*\{', html)
    if not m:
        return {}
    start = m.end() - 1
    depth, pos = 0, start
    while pos < len(html):
        if html[pos] == '{': depth += 1
        elif html[pos] == '}':
            depth -= 1
            if depth == 0: break
        pos += 1
    body = html[start:pos + 1]

    buildings = {}
    for bm in re.finditer(r"(\w+):\{days:\[([^\]]*)\],panels:\[", body):
        key = bm.group(1)
        days = [int(d) for d in bm.group(2).split(',') if d.strip()]
        # scan panels array from just after 'panels:[' to its matching ']'
        p_start = bm.end() - 1
        depth2, ppos = 0, p_start
        while ppos < len(body):
            if body[ppos] == '[': depth2 += 1
            elif body[ppos] == ']':
                depth2 -= 1
                if depth2 == 0: break
            ppos += 1
        panels_block = body[p_start:ppos + 1]
        panels = []
        for pm in re.finditer(r"name:'([^']*)',daily:\[([^\]]*)\]", panels_block):
            name = pm.group(1).replace("\\'", "'")
            daily = [float(v) for v in pm.group(2).split(',') if v.strip()]
            panels.append({'name': name, 'daily': daily})
        buildings[key] = {'days': days, 'panels': panels}
    return buildings


def extract_labels_vals(html, var_name):
    """Extract a `const NAME = {labels:[...], vals:[...]};` shaped object."""
    m = re.search(rf'const\s+{re.escape(var_name)}\s*=\s*\{{labels:\[([^\]]*)\],vals:\[([^\]]*)\]\}}', html)
    if not m:
        return {'labels': [], 'vals': []}
    labels = [t.strip().strip("'") for t in m.group(1).split(',') if t.strip()]
    vals = [float(t) for t in m.group(2).split(',') if t.strip()]
    return {'labels': labels, 'vals': vals}


def extract_js_scalar(html, var_name):
    m = re.search(rf'const\s+{re.escape(var_name)}\s*=\s*(-?\d+(?:\.\d+)?|null);', html)
    if not m or m.group(1) == 'null':
        return None
    return float(m.group(1))


def load_dashboard_data():
    html = DASHBOARD_FILE.read_text(encoding='utf-8')
    days_m = re.search(r"const CUR_DAYS\s*=\s*\[([^\]]+)\]", html)
    day_labels = []
    if days_m:
        day_labels = [d.strip().strip("'") for d in days_m.group(1).split(',')]
    month_abbr = day_labels[0].split('-')[1] if day_labels else 'Jun'
    yr_m = re.search(r"const\s+CUR_YEAR\s*=\s*(\d+);", html)
    year = int(yr_m.group(1)) if yr_m else datetime.now().year
    month_label = f"{month_abbr} {year}"
    return {
        'html': html, 'month': month_label,
        'day_labels': day_labels, 'n_days': len(day_labels),
        'eb': extract_js_array(html, 'curEB'),
        'dg': extract_js_array(html, 'curDG'),
        'diesel': extract_js_array(html, 'curDiesel'),
        'stock': extract_js_array(html, 'curStock'),
        'power_cuts': extract_power_cuts(html),
        'ev': extract_js_array(html, 'evDailyKWh'),
        'nescafe': extract_js_array(html, 'nescafeDailyKWh'),
        'cvb': extract_js_array(html, 'cvbDailyKWh'),
        'tea': extract_js_array(html, 'teaDailyKWh'),
        'yummy': extract_js_array(html, 'yummyDailyKWh'),
        'laundry_e': extract_js_array(html, 'laundryElecKWh'),
        'wtp_daily': extract_wtp_daily(html),
        'laundry_w': extract_js_array(html, 'laundryDaily'),
        'new_well': extract_js_array(html, 'newWellJunKL'),
        'ac_outlet': extract_js_array(html, 'acOutletTotal'),
        'ac_blocks': extract_nested_array(html, 'acOutletBlocks'),
        'ro_cans': extract_js_array(html, 'roCansDaily'),
        'tds_ro2000': extract_js_array(html, 'tdsRo2000'),
        'buildings': extract_buildings(html),
        'tickets_by_dept': extract_labels_vals(html, 'ticketsByDept'),
        'ticket_total': extract_js_scalar(html, 'ticketTotal'),
        'ticket_pending': extract_js_scalar(html, 'ticketPending'),
        'ticket_avg_tat_min': extract_js_scalar(html, 'ticketAvgTatMin'),
        'ticket_recurring_by_dept': extract_labels_vals(html, 'ticketRecurringByDept'),
    }


# ── Analytics ─────────────────────────────────────────────────────────────────

def s(arr):   return sum(v for v in (arr or []) if v)
def nz(arr):  return [v for v in (arr or []) if v and v > 0]
def avg(arr): vals = nz(arr); return sum(vals)/len(vals) if vals else 0
def fmt_min_hm(minutes):
    m = int(round(minutes))
    return f"{m // 60}h {m % 60}m"
def peak_day(arr, labels):
    pairs = [(v, labels[i]) for i, v in enumerate(arr or []) if v and i < len(labels)]
    return max(pairs, key=lambda x: x[0]) if pairs else (0, 'N/A')


def compute_insights(d):
    html   = d['html']
    labels = d['day_labels']
    eb, dg, die = d['eb'], d['dg'], d['diesel']

    eb_total  = s(eb);  dg_total = s(dg);  die_total = s(die)
    total_pwr = eb_total + dg_total
    dg_pct    = (dg_total / total_pwr * 100) if total_pwr else 0

    eb_peak_val, eb_peak_day = peak_day(eb, labels)
    dg_peak_val, dg_peak_day = peak_day(dg, labels)
    diesel_cost = die_total * 92

    rm_solar = 13608
    rm_m = re.search(r'rmSolarMonthly\s*=\s*\{[^}]*totals:\[([^\]]*)\]', html)
    if rm_m:
        totals = [float(v) for v in rm_m.group(1).split(',') if v.strip()]
        if len(totals) >= 6: rm_solar = totals[5]
    solar_pct = (rm_solar / eb_total * 100) if eb_total else 0

    cuts = d['power_cuts']
    cut_count   = len(cuts)
    cut_min     = sum(c['dur'] for c in cuts)
    worst_cut   = max(cuts, key=lambda c: c['dur']) if cuts else None

    wtp = d['wtp_daily']
    wtp1t = s(wtp.get('wtp1',[])); wtp2t = s(wtp.get('wtp2',[])); wtp3t = s(wtp.get('wtp3',[]))
    wtp_total = wtp1t + wtp2t + wtp3t

    lw = d['laundry_w']
    laundry_w_total  = s(lw);  laundry_active = len(nz(lw))
    nw = d['new_well']
    new_well_total = s(nw);  new_well_active = len(nz(nw))

    ac_total = s(d['ac_outlet'])
    ac_peak_val, ac_peak_day = peak_day(d['ac_outlet'], labels)

    ro_cans = d['ro_cans']
    ro_total  = s(ro_cans);  ro_avg = avg(ro_cans)
    has_ro_data = any(v and v > 0 for v in ro_cans)

    tds = [v for v in d['tds_ro2000'] if v and v > 0]
    tds_avg = avg(tds); tds_max = max(tds) if tds else 0; tds_min = min(tds) if tds else 0

    ev_t  = s(d['ev']); nesc_t = s(d['nescafe']); cvb_t = s(d['cvb'])
    tea_t = s(d['tea']); yum_t = s(d['yummy']); le_t  = s(d['laundry_e'])
    tenant_total = ev_t + nesc_t + cvb_t + tea_t + yum_t + le_t

    # ── Digii Tickets (all-time cumulative, not month-scoped like everything
    # else — the source file is always the full history) ──
    tickets_by_dept = d['tickets_by_dept']
    ticket_total = int(d['ticket_total'] or 0)
    ticket_pending = int(d['ticket_pending'] or 0)
    ticket_avg_tat_min = d['ticket_avg_tat_min']
    recurring_by_dept = d['ticket_recurring_by_dept']
    ticket_recurring_total = int(sum(recurring_by_dept['vals'])) if recurring_by_dept['vals'] else 0
    has_ticket_data = ticket_total > 0

    # ── Building sub-meter panels ──
    BUILDING_LABELS = {
        'substation': 'Substation', 'mainAcademic': 'Main Academic',
        'nalanda': 'Nalanda Block', 'oldHostel': 'Old Hostel',
        'pgdmHostel': 'PGDM Hostel', 'canteen': 'Canteen',
        'canteenAddOn': 'Canteen Add-On', 'arrBlock': 'ARR Block',
        'thiruvalluvar': 'Thiruvalluvar Block', 'saraswati': 'Saraswati Block',
        'cBlock': 'C-Block', 'bramahputra': 'Bramahputra',
    }
    buildings_raw = d.get('buildings', {})
    # Merge Canteen + Canteen Add-On into one row, matching the dashboard UI
    building_rows = []
    canteen_total = s(sum((p['daily'] for p in buildings_raw.get('canteen', {}).get('panels', [])), []))
    addon_total = s(sum((p['daily'] for p in buildings_raw.get('canteenAddOn', {}).get('panels', [])), []))
    for key, b in buildings_raw.items():
        if key == 'canteenAddOn':
            continue
        total = s(sum((p['daily'] for p in b['panels']), []))
        if key == 'canteen':
            total += addon_total
        building_rows.append({
            'label': BUILDING_LABELS.get(key, key),
            'panels': len(b['panels']) + (len(buildings_raw.get('canteenAddOn', {}).get('panels', [])) if key == 'canteen' else 0),
            'total': total,
        })
    buildings_total = sum(r['total'] for r in building_rows)

    # active EB days
    eb_active = [v for v in eb if v and v > 0]
    eb_avg_day = sum(eb_active)/len(eb_active) if eb_active else 0

    flags = []
    if dg_pct > 3:
        flags.append(('alert', f'DG accounted for {dg_pct:.1f}% of total power — review EB supply reliability for next month'))
    if cut_min > 1000:
        flags.append(('alert', f'{cut_min} min of outages in {d["month"]} ({cut_count} events) — highest risk: {worst_cut["date"]} ({worst_cut["dur"]} min)'))
    elif cut_count > 10:
        flags.append(('warn', f'{cut_count} power cut events — frequency trending up, plan DG fuel buffer'))
    if solar_pct < 8:
        flags.append(('warn', f'RM Solar offset only {solar_pct:.1f}% of EB consumption — check panel cleaning schedule'))
    if tds_avg > 28:
        flags.append(('alert', f'RO-2000 TDS avg {tds_avg:.1f} ppm — above 28 ppm threshold, check membrane'))
    elif tds_avg > 0:
        flags.append(('ok', f'RO-2000 water quality good — TDS avg {tds_avg:.1f} ppm (range {tds_min}–{tds_max} ppm)'))
    if ro_avg > 100:
        flags.append(('warn', f'RO can demand avg {ro_avg:.0f} cans/day — ensure adequate RO-2000 uptime'))
    amc_m = re.findall(r"\{system:'([^']+)'[^}]*status:'(Overdue|Critical)'", html)
    for sys_name, status in amc_m:
        flags.append(('alert' if status=='Critical' else 'warn',
                      f'AMC {status}: {sys_name} — schedule service immediately'))
    if d['ac_blocks'] and any(s(b) > 0 for b in d['ac_blocks']):
        block_totals = [s(b) for b in d['ac_blocks']]
        top_idx = block_totals.index(max(block_totals))
        blabels = ['Block 1','Block 2','Block 3','Block 5','Block 6',
                   'ARR Block','Nalanda','Faculty','Bharathi','Gym','Manimegalai']
        flags.append(('info', f'Highest AC outlet consumer: {blabels[top_idx] if top_idx<len(blabels) else "Block "+str(top_idx+1)} ({block_totals[top_idx]:.0f} L in {d["month"]})'))
    if new_well_total > 0:
        flags.append(('info', f'New well pump delivered {new_well_total:.0f} KL over {new_well_active} active days'))
    if laundry_active < 15:
        flags.append(('warn', f'Laundromat active only {laundry_active}/{len(lw)} days — investigate downtime'))

    # ── Dynamic outlook: next-month action items driven by current data ──────
    from calendar import month_name
    try:
        cur_month_num = datetime.strptime(d['month'], "%b %Y").month
    except Exception:
        cur_month_num = datetime.now().month
    next_month_num  = (cur_month_num % 12) + 1
    next_month_name = month_name[next_month_num]

    # month-on-month EB trend (compare last 7 days avg to first 7 days avg)
    eb_nz = nz(eb)
    recent_avg  = sum(eb_nz[-7:]) / len(eb_nz[-7:])  if len(eb_nz) >= 7 else eb_avg_day
    early_avg   = sum(eb_nz[:7])  / len(eb_nz[:7])   if len(eb_nz) >= 7 else eb_avg_day
    eb_trending = "up" if recent_avg > early_avg * 1.05 else ("down" if recent_avg < early_avg * 0.95 else "stable")

    outlook = []

    # 1. Energy budget
    if eb_trending == "up":
        outlook.append(
            f"EB consumption trended UP this month (recent avg {recent_avg:,.0f} kWh/day vs early avg {early_avg:,.0f} kWh/day) — "
            f"budget {recent_avg*1.05:,.0f} kWh/day for {next_month_name} and audit AC usage to contain growth."
        )
    elif eb_trending == "down":
        outlook.append(
            f"EB consumption trended DOWN this month ({recent_avg:,.0f} kWh/day recently) — "
            f"target {recent_avg:,.0f} kWh/day for {next_month_name} and document what drove the saving."
        )
    else:
        outlook.append(
            f"EB daily avg stable at {eb_avg_day:,.0f} kWh — budget similarly for {next_month_name}; "
            f"monsoon/seasonal load shifts may affect AC demand."
        )

    # 2. DG / power cuts
    if dg_pct > 5:
        outlook.append(
            f"DG share was {dg_pct:.1f}% this month ({cut_count} cut events, {cut_min} min total) — "
            f"maintain minimum 600 L diesel stock for {next_month_name} and escalate EB supply reliability with TANGEDCO."
        )
    elif cut_count > 0:
        outlook.append(
            f"{cut_count} power cut event(s) this month — maintain min 400 L diesel buffer; "
            f"inspect DG coolant and battery before {next_month_name}."
        )

    # 3. Solar
    if solar_pct < 8:
        outlook.append(
            f"Solar offset was low at {solar_pct:.1f}% — schedule panel cleaning and check inverter logs before {next_month_name} "
            f"to recover generation; target >10% offset."
        )
    elif solar_pct >= 12:
        outlook.append(
            f"Solar performing well at {solar_pct:.1f}% offset — continue monitoring; check for shading or soiling monthly."
        )

    # 4. RO water (only when the source Excel actually had RO can data for this month)
    if has_ro_data:
        if tds_avg > 28:
            outlook.append(
                f"RO-2000 TDS averaged {tds_avg:.1f} ppm (above 28 ppm threshold) — replace or service membrane before {next_month_name}; "
                f"pre-stock RO cans as backup while membrane is serviced."
            )
        else:
            ro_buffer = round(ro_avg * 3)
            outlook.append(
                f"RO demand averaged {ro_avg:.0f} cans/day — pre-stock {ro_buffer} cans at {next_month_name} start "
                f"(3-day buffer) to avoid shortages during peak demand days."
            )

    # 5. Digii Tickets (all-time cumulative — insight reflects the whole
    # history to date, not just this reporting month)
    if has_ticket_data:
        if tickets_by_dept['labels']:
            top_dept = tickets_by_dept['labels'][0]
            top_count = int(tickets_by_dept['vals'][0])
            outlook.append(
                f"{top_dept} has raised the most Digii tickets to date ({top_count:,}) — conduct a root-cause "
                f"review and schedule preventive maintenance to reduce reactive calls."
            )
        if ticket_recurring_total > 0 and recurring_by_dept['labels']:
            top_recur_dept = recurring_by_dept['labels'][0]
            top_recur_count = int(recurring_by_dept['vals'][0])
            outlook.append(
                f"{ticket_recurring_total} tickets were re-raised by the same person for the same problem within "
                f"3 days of being marked resolved — {top_recur_dept} accounts for the most ({top_recur_count}). "
                f"Fixes here may not be holding; recommend a follow-up inspection pass."
            )
        if ticket_pending > 0:
            outlook.append(
                f"{ticket_pending} Digii ticket(s) remain pending — follow up before {next_month_name} to keep the backlog from growing."
            )

    # 6. Water / well
    if new_well_active > 0:
        days_in_month = d['n_days'] or 30
        well_uptime = new_well_active / days_in_month * 100
        if well_uptime < 60:
            outlook.append(
                f"New well pump was active only {new_well_active}/{days_in_month} days ({well_uptime:.0f}% uptime) — "
                f"inspect pump and check borewell yield before {next_month_name}."
            )
        else:
            outlook.append(
                f"New well delivered {new_well_total:,.0f} KL over {new_well_active} days — "
                f"continue monitoring pump health; check motor amperage monthly."
            )

    # 7. Laundromat
    if laundry_active < 20:
        outlook.append(
            f"Laundromat operated only {laundry_active} days this month — verify machine availability and "
            f"plan for full-month operation in {next_month_name} to avoid backlog."
        )

    # 8. AMC items from flags (pull critical/overdue systems)
    amc_critical = [msg for sev, msg in flags if sev in ('alert', 'warn') and 'AMC' in msg]
    if amc_critical:
        systems = '; '.join(m.split('AMC')[1].split('—')[0].strip() for m in amc_critical)
        outlook.append(
            f"Pending AMC actions: {systems} — confirm service vendor appointments before end of {next_month_name}."
        )

    return dict(
        month=d['month'], n_days=d['n_days'],
        next_month=next_month_name,
        eb_total=eb_total, dg_total=dg_total, die_total=die_total,
        eb_peak_val=eb_peak_val, eb_peak_day=eb_peak_day,
        dg_peak_val=dg_peak_val, dg_peak_day=dg_peak_day,
        dg_pct=dg_pct, diesel_cost=diesel_cost,
        rm_solar=rm_solar, solar_pct=solar_pct,
        cut_count=cut_count, cut_min=cut_min, worst_cut=worst_cut, cuts=cuts,
        wtp_total=wtp_total, wtp1t=wtp1t, wtp2t=wtp2t, wtp3t=wtp3t,
        laundry_w_total=laundry_w_total, laundry_active=laundry_active,
        new_well_total=new_well_total, new_well_active=new_well_active,
        ac_total=ac_total, ac_peak_val=ac_peak_val, ac_peak_day=ac_peak_day,
        ro_total=ro_total, ro_avg=ro_avg,
        tds_avg=tds_avg, tds_min=tds_min, tds_max=tds_max,
        ev_t=ev_t, nesc_t=nesc_t, cvb_t=cvb_t,
        tea_t=tea_t, yum_t=yum_t, le_t=le_t, tenant_total=tenant_total,
        ticket_total=ticket_total, ticket_pending=ticket_pending,
        ticket_avg_tat_min=ticket_avg_tat_min, tickets_by_dept=tickets_by_dept,
        ticket_recurring_total=ticket_recurring_total, recurring_by_dept=recurring_by_dept,
        has_ticket_data=has_ticket_data,
        eb_avg_day=eb_avg_day, eb_trending=eb_trending,
        building_rows=building_rows, buildings_total=buildings_total,
        flags=flags, outlook=outlook,
    )


# ── PDF Custom Flowables ───────────────────────────────────────────────────────

class ColorRect(Flowable):
    """A solid-filled rectangle, used for header band."""
    def __init__(self, width, height, color):
        super().__init__()
        self.width  = width
        self.height = height
        self.color  = color

    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.rect(0, 0, self.width, self.height, fill=1, stroke=0)


class KPIRow(Flowable):
    """A row of KPI cards drawn directly on canvas."""
    def __init__(self, items, page_width, padding):
        super().__init__()
        self.items     = items   # list of (label, value, unit, color_hex)
        self.n         = len(items)
        usable         = page_width - 2*padding
        self.card_w    = usable / self.n
        self.height    = 52
        self.width     = usable

    def draw(self):
        c = self.canv
        gap = 4
        for i, (label, value, unit, color) in enumerate(self.items):
            x = i * self.card_w
            w = self.card_w - gap

            # Card background
            c.setFillColor(HexColor("#F9FAFB"))
            c.setStrokeColor(BORDER)
            c.setLineWidth(0.5)
            c.roundRect(x, 0, w, self.height, 4, fill=1, stroke=1)

            # Top colour stripe
            c.setFillColor(HexColor(color))
            c.rect(x, self.height-3, w, 3, fill=1, stroke=0)

            # Label
            c.setFillColor(INK4)
            c.setFont("Helvetica", 7)
            c.drawString(x+8, self.height-14, label.upper())

            # Value
            c.setFillColor(INK)
            c.setFont("Helvetica-Bold", 16)
            c.drawString(x+8, 20, value)

            # Unit
            c.setFillColor(INK3)
            c.setFont("Helvetica", 8)
            c.drawString(x+8, 8, unit)


class HBarChart(Flowable):
    """Simple horizontal bar chart for tenant sub-meters."""
    def __init__(self, items, page_width, padding):
        # items: list of (label, value, color_hex)
        super().__init__()
        self.items   = items
        self.width   = page_width - 2*padding
        self.height  = len(items) * 22
        self.max_val = max(v for _, v, _ in items) if items else 1

    def draw(self):
        c   = self.canv
        bar_start = 110
        bar_width = self.width - bar_start - 60
        row_h = 22

        for i, (label, value, color) in enumerate(self.items):
            y = self.height - (i+1)*row_h + 4

            # Label
            c.setFillColor(INK2)
            c.setFont("Helvetica", 8.5)
            c.drawString(0, y+5, label)

            # Background track
            c.setFillColor(HexColor("#F3F4F6"))
            c.roundRect(bar_start, y+2, bar_width, 12, 3, fill=1, stroke=0)

            # Value bar
            fill_w = int(bar_width * value / self.max_val)
            if fill_w > 0:
                c.setFillColor(HexColor(color))
                c.roundRect(bar_start, y+2, fill_w, 12, 3, fill=1, stroke=0)

            # Value text
            c.setFillColor(INK)
            c.setFont("Helvetica-Bold", 8)
            c.drawRightString(bar_start + bar_width + 55, y+5, f"{value:,.0f} kWh")


class FlagRow(Flowable):
    """One insight/flag row with coloured dot and text."""
    def __init__(self, severity, message, width):
        super().__init__()
        self.severity = severity
        self.message  = message
        self.width    = width
        self.height   = 28

    def draw(self):
        c = self.canv
        color_map = {'alert': '#DC2626', 'warn': '#D97706', 'ok': '#059669', 'info': '#1B4FD8'}
        bg_map    = {'alert': '#FEF2F2', 'warn': '#FFFBEB', 'ok': '#F0FDF4', 'info': '#EFF6FF'}
        icon_map  = {'alert': '!', 'warn': '~', 'ok': 'OK', 'info': 'i'}

        col = HexColor(color_map.get(self.severity, '#6B7280'))
        bg  = HexColor(bg_map.get(self.severity, '#F9FAFB'))

        # Background
        c.setFillColor(bg)
        c.setStrokeColor(col)
        c.setLineWidth(0.4)
        c.roundRect(0, 2, self.width, self.height-4, 4, fill=1, stroke=1)

        # Dot
        c.setFillColor(col)
        c.circle(12, self.height/2, 5, fill=1, stroke=0)
        c.setFillColor(white)
        c.setFont("Helvetica-Bold", 6)
        c.drawCentredString(12, self.height/2 - 2.5, icon_map.get(self.severity,'•'))

        # Message text — wrap if needed
        c.setFillColor(INK2)
        c.setFont("Helvetica", 8.5)
        msg = self.message
        if len(msg) > 110:
            mid = msg[:110].rfind(' ')
            c.drawString(24, self.height/2 + 2, msg[:mid])
            c.drawString(24, self.height/2 - 9, msg[mid+1:])
        else:
            c.drawString(24, self.height/2 - 3, msg)


class OutlookRow(Flowable):
    """One outlook action item — purple left accent bar, white card, wrapped text."""
    ACCENT = HexColor("#7C3AED")
    BG     = HexColor("#F5F3FF")
    BORDER_C = HexColor("#DDD6FE")

    def __init__(self, message, width):
        super().__init__()
        self.message = message
        self.width   = width
        # estimate height: ~75 chars per line at 8.5pt, min 30
        chars_per_line = int((width - 30) / 5.1)
        lines = max(1, -(-len(message) // chars_per_line))  # ceiling div
        self.height = max(30, lines * 13 + 12)

    def draw(self):
        c   = self.canv
        h   = self.height
        w   = self.width

        # Card background
        c.setFillColor(self.BG)
        c.setStrokeColor(self.BORDER_C)
        c.setLineWidth(0.4)
        c.roundRect(0, 0, w, h, 4, fill=1, stroke=1)

        # Left accent bar
        c.setFillColor(self.ACCENT)
        c.rect(0, 0, 4, h, fill=1, stroke=0)

        # Bullet circle
        cy = h / 2
        c.setFillColor(self.ACCENT)
        c.circle(16, cy, 4, fill=1, stroke=0)
        c.setFillColor(white)
        c.setFont("Helvetica-Bold", 7)
        c.drawCentredString(16, cy - 2.5, "")

        # Text — wrap across lines
        c.setFillColor(HexColor("#1E1B4B"))
        c.setFont("Helvetica", 8.5)
        text_x    = 28
        text_w    = w - text_x - 8
        chars_per_line = int(text_w / 5.1)
        words     = self.message.split()
        lines     = []
        cur       = ""
        for word in words:
            test = (cur + " " + word).strip()
            if len(test) <= chars_per_line:
                cur = test
            else:
                if cur:
                    lines.append(cur)
                cur = word
        if cur:
            lines.append(cur)

        total_h = len(lines) * 13
        start_y = (h + total_h) / 2 - 10
        for i, line in enumerate(lines):
            c.drawString(text_x, start_y - i * 13, line)


# ── PDF Builder ────────────────────────────────────────────────────────────────

def build_pdf(ins: dict) -> bytes:
    buf   = io.BytesIO()
    W, H  = A4
    PAD   = 18*mm
    TW    = W - 2*PAD   # text width

    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=PAD, rightMargin=PAD,
        topMargin=14*mm, bottomMargin=14*mm,
        title=f"GL Campus Intelligence — {ins['month']}",
        author="GL Campus Intelligence System",
    )

    styles = getSampleStyleSheet()

    def style(name, **kw):
        return ParagraphStyle(name, **kw)

    S = {
        'h1':    style('h1',    fontName='Helvetica-Bold',   fontSize=18, textColor=white,   leading=22, spaceAfter=2),
        'sub':   style('sub',   fontName='Helvetica',        fontSize=10, textColor=HexColor('#CBD5E1'), leading=14),
        'sec':   style('sec',   fontName='Helvetica-Bold',   fontSize=11, textColor=BLUE,    leading=16, spaceBefore=14, spaceAfter=4),
        'body':  style('body',  fontName='Helvetica',        fontSize=8.5,textColor=INK2,    leading=13),
        'small': style('small', fontName='Helvetica',        fontSize=7.5,textColor=INK3,    leading=11),
        'bold':  style('bold',  fontName='Helvetica-Bold',   fontSize=8.5,textColor=INK,     leading=13),
        'th':    style('th',    fontName='Helvetica-Bold',   fontSize=7.5,textColor=INK3,    leading=11, alignment=TA_LEFT),
        'td':    style('td',    fontName='Helvetica',        fontSize=8,  textColor=INK2,    leading=12, alignment=TA_LEFT),
        'tdc':   style('tdc',   fontName='Helvetica',        fontSize=8,  textColor=INK2,    leading=12, alignment=TA_CENTER),
        'num':   style('num',   fontName='Helvetica-Bold',   fontSize=8,  textColor=INK,     leading=12, alignment=TA_RIGHT),
        'right': style('right', fontName='Helvetica',        fontSize=8,  textColor=INK3,    leading=12, alignment=TA_RIGHT),
        'label': style('label', fontName='Helvetica',        fontSize=7,  textColor=INK4,    leading=10, alignment=TA_CENTER),
        'foot':  style('foot',  fontName='Helvetica',        fontSize=7,  textColor=INK4,    leading=10, alignment=TA_CENTER),
    }

    story = []
    now   = datetime.now().strftime('%d %b %Y, %I:%M %p')
    month = ins['month']

    # ── HEADER BAND ────────────────────────────────────────────────────────────
    header_data = [[
        Paragraph(f"<b>GL Campus Intelligence Report</b><br/>"
                  f"<font size='10' color='#CBD5E1'>{month} &nbsp;·&nbsp; Manamai Campus</font>",
                  style('hdr', fontName='Helvetica-Bold', fontSize=18, textColor=white,
                        leading=24, spaceAfter=0)),
        Paragraph(f"<font size='7' color='#94A3B8'>GENERATED</font><br/>"
                  f"<font size='9' color='white'><b>{now}</b></font><br/>"
                  f"<font size='7' color='#94A3B8'>Great Lakes Institute of Management</font>",
                  style('hdr_r', fontName='Helvetica', fontSize=9, textColor=white,
                        leading=14, alignment=TA_RIGHT)),
    ]]
    hdr_tbl = Table(header_data, colWidths=[TW*0.6, TW*0.4])
    hdr_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), BLUE),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 14),
        ('BOTTOMPADDING',(0,0),(-1,-1), 14),
        ('LEFTPADDING', (0,0),(0,-1), 16),
        ('RIGHTPADDING',(-1,0),(-1,-1),16),
        ('ROUNDEDCORNERS', [6]),
    ]))
    story.append(hdr_tbl)
    story.append(Spacer(1, 10))

    # ── EXECUTIVE KPIs ─────────────────────────────────────────────────────────
    kpi_items = [
        ('EB Consumed',   f"{ins['eb_total']:,.0f}", 'kWh',     '#1B4FD8'),
        ('DG Units',      f"{ins['dg_total']:,.0f}", 'kWh',     '#EA580C'),
        ('Water Produced',f"{ins['wtp_total']:,.0f}",'KL',      '#0891B2'),
    ]
    if ins.get('has_ticket_data'):
        kpi_items.append(('Digii Tickets', f"{ins['ticket_total']:,}", 'all-time', '#059669'))
    kpi_items.append(('Power Outages', f"{ins['cut_count']}", f"{ins['cut_min']} min total", '#DC2626'))
    story.append(KPIRow(kpi_items, W, PAD))
    story.append(Spacer(1, 12))

    # ── SECTION: ENERGY ────────────────────────────────────────────────────────
    story.append(HRFlowable(width=TW, thickness=0.5, color=BORDER))
    story.append(Paragraph("⚡  Energy", S['sec']))

    # Energy summary table
    e_data = [
        [Paragraph('Metric', S['th']),            Paragraph('Value', S['th']),          Paragraph('Notes', S['th'])],
        [Paragraph('Total EB Consumption', S['td']),  Paragraph(f"{ins['eb_total']:,.0f} kWh", S['num']), Paragraph(f"Avg {ins['eb_avg_day']:,.0f} kWh/day", S['small'])],
        [Paragraph('Total DG Generation', S['td']),   Paragraph(f"{ins['dg_total']:,.0f} kWh", S['num']), Paragraph(f"{ins['dg_pct']:.2f}% of total power", S['small'])],
        [Paragraph('Diesel Consumed', S['td']),       Paragraph(f"{ins['die_total']:,.1f} L", S['num']),  Paragraph(f"Est. cost Rs.{ins['diesel_cost']:,.0f} @ Rs.92/L", S['small'])],
        [Paragraph('RM Solar Generation', S['td']),   Paragraph(f"{ins['rm_solar']:,.0f} kWh", S['num']), Paragraph(f"{ins['solar_pct']:.1f}% offset of EB grid draw", S['small'])],
        [Paragraph('EB Peak Day', S['td']),           Paragraph(f"{ins['eb_peak_day']}", S['num']),       Paragraph(f"{ins['eb_peak_val']:,.0f} kWh", S['small'])],
        [Paragraph('Highest DG Day', S['td']),        Paragraph(f"{ins['dg_peak_day']}", S['num']),       Paragraph(f"{ins['dg_peak_val']:,.0f} kWh via generator", S['small'])],
    ]
    e_tbl = Table(e_data, colWidths=[TW*0.38, TW*0.28, TW*0.34])
    e_tbl.setStyle(TableStyle([
        ('BACKGROUND',   (0,0),(-1,0),  HexColor('#F3F4F6')),
        ('BACKGROUND',   (0,2),(-1,2),  HexColor('#FAFAFA')),
        ('BACKGROUND',   (0,4),(-1,4),  HexColor('#FAFAFA')),
        ('BACKGROUND',   (0,6),(-1,6),  HexColor('#FAFAFA')),
        ('GRID',         (0,0),(-1,-1), 0.4, BORDER),
        ('TOPPADDING',   (0,0),(-1,-1), 5),
        ('BOTTOMPADDING',(0,0),(-1,-1), 5),
        ('LEFTPADDING',  (0,0),(-1,-1), 8),
        ('RIGHTPADDING', (0,0),(-1,-1), 8),
        ('VALIGN',       (0,0),(-1,-1), 'MIDDLE'),
    ]))
    story.append(e_tbl)
    story.append(Spacer(1, 10))

    # Power cuts top 5
    story.append(Paragraph("Power Outages — Top 5 by Duration", S['bold']))
    story.append(Spacer(1, 4))
    sorted_cuts = sorted(ins['cuts'], key=lambda c: c['dur'], reverse=True)[:5]
    c_data = [
        [Paragraph('Date', S['th']), Paragraph('Time Window', S['th']),
         Paragraph('Duration', S['th']), Paragraph('DG Used', S['th']), Paragraph('Severity', S['th'])],
    ]
    for c in sorted_cuts:
        sev = 'High' if c['dur']>200 else ('Medium' if c['dur']>60 else 'Low')
        sev_col = RED if c['dur']>200 else (YELLOW if c['dur']>60 else GREEN)
        c_data.append([
            Paragraph(c['date'], S['td']),
            Paragraph(c['events'], S['td']),
            Paragraph(f"{c['dur']} min", style('cut_dur', fontName='Helvetica-Bold', fontSize=8, textColor=sev_col, leading=12)),
            Paragraph(c['dg'], S['td']),
            Paragraph(sev, style('sev', fontName='Helvetica-Bold', fontSize=8, textColor=sev_col, leading=12)),
        ])
    c_tbl = Table(c_data, colWidths=[TW*0.14, TW*0.28, TW*0.16, TW*0.20, TW*0.22])
    c_tbl.setStyle(TableStyle([
        ('BACKGROUND',   (0,0),(-1,0),  HexColor('#F3F4F6')),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[white, HexColor('#FAFAFA')]),
        ('GRID',         (0,0),(-1,-1), 0.4, BORDER),
        ('TOPPADDING',   (0,0),(-1,-1), 5),
        ('BOTTOMPADDING',(0,0),(-1,-1), 5),
        ('LEFTPADDING',  (0,0),(-1,-1), 8),
        ('RIGHTPADDING', (0,0),(-1,-1), 8),
        ('VALIGN',       (0,0),(-1,-1), 'MIDDLE'),
    ]))
    story.append(c_tbl)
    story.append(Spacer(1, 10))

    # Tenant sub-meters
    story.append(Paragraph("Tenant Sub-Meter Consumption", S['bold']))
    story.append(Spacer(1, 4))
    tenant_items = [
        ('CVB Beverages',    ins['cvb_t'],   '#1B4FD8'),
        ('Laundromat (Elec)',ins['le_t'],    '#0891B2'),
        ('Nescafe',          ins['nesc_t'],  '#D97706'),
        ('EV Charging',      ins['ev_t'],    '#059669'),
        ("Yummy's",          ins['yum_t'],   '#7C3AED'),
        ('Tea Wheeler',      ins['tea_t'],   '#9CA3AF'),
    ]
    story.append(HBarChart(tenant_items, W, PAD))
    story.append(Spacer(1, 4))
    story.append(Paragraph(f"Total tenant consumption: {ins['tenant_total']:,.0f} kWh for {month}", S['small']))
    story.append(Spacer(1, 8))

    # ── WATER + WORK ORDERS + INSIGHTS + OUTLOOK + FOOTER — one KeepTogether
    # Keeps all remaining content together so pages fill top-to-bottom cleanly.
    tail = []

    tail.append(HRFlowable(width=TW, thickness=0.5, color=BORDER))
    tail.append(Paragraph("Water", S['sec']))

    w_data = [
        [Paragraph('Source / Metric', S['th']),      Paragraph('Value', S['th']),                          Paragraph('Notes', S['th'])],
        [Paragraph('WTP-1 Production', S['td']),     Paragraph(f"{ins['wtp1t']:,.0f} KL", S['num']),       Paragraph('Main treatment plant', S['small'])],
        [Paragraph('WTP-2 Production', S['td']),     Paragraph(f"{ins['wtp2t']:,.0f} KL", S['num']),       Paragraph('Secondary plant', S['small'])],
        [Paragraph('WTP-3 Production', S['td']),     Paragraph(f"{ins['wtp3t']:,.0f} KL", S['num']),       Paragraph('Tertiary plant', S['small'])],
        [Paragraph('Total WTP Output', S['bold']),   Paragraph(f"{ins['wtp_total']:,.0f} KL",
             style('wtp_tot', fontName='Helvetica-Bold', fontSize=8.5, textColor=TEAL,
                   leading=12, alignment=TA_RIGHT)),                                                         Paragraph(f'Combined {month} production', S['small'])],
        [Paragraph('New Well Pump', S['td']),        Paragraph(f"{ins['new_well_total']:,.0f} KL", S['num']), Paragraph(f"{ins['new_well_active']} active days", S['small'])],
        [Paragraph('Laundromat Water', S['td']),     Paragraph(f"{ins['laundry_w_total']:.0f} KL", S['num']), Paragraph(f"{ins['laundry_active']} active days", S['small'])],
        [Paragraph('AC Outlet Consumption', S['td']),Paragraph(f"{ins['ac_total']:,.0f} L", S['num']),      Paragraph(f"Peak: {ins['ac_peak_day']} ({ins['ac_peak_val']:.0f} L)", S['small'])],
        [Paragraph('RO Cans Dispensed', S['td']),   Paragraph(f"{ins['ro_total']:,.0f} cans", S['num']),   Paragraph(f"Avg {ins['ro_avg']:.0f} cans/day", S['small'])],
        [Paragraph('TDS — RO 2000 (avg)', S['td']), Paragraph(f"{ins['tds_avg']:.1f} ppm", S['num']),     Paragraph(f"Range {ins['tds_min']}–{ins['tds_max']} ppm  |  threshold 28 ppm", S['small'])],
    ]
    w_tbl = Table(w_data, colWidths=[TW*0.38, TW*0.28, TW*0.34])
    w_tbl.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,0),  HexColor('#F3F4F6')),
        ('BACKGROUND',    (0,2),(-1,2),  HexColor('#FAFAFA')),
        ('BACKGROUND',    (0,4),(-1,4),  HexColor('#EFF6FF')),
        ('BACKGROUND',    (0,6),(-1,6),  HexColor('#FAFAFA')),
        ('BACKGROUND',    (0,8),(-1,8),  HexColor('#FAFAFA')),
        ('GRID',          (0,0),(-1,-1), 0.4, BORDER),
        ('TOPPADDING',    (0,0),(-1,-1), 5),
        ('BOTTOMPADDING', (0,0),(-1,-1), 5),
        ('LEFTPADDING',   (0,0),(-1,-1), 8),
        ('RIGHTPADDING',  (0,0),(-1,-1), 8),
        ('VALIGN',        (0,0),(-1,-1), 'MIDDLE'),
    ]))
    tail.append(w_tbl)
    tail.append(Spacer(1, 8))

    # Buildings — sub-meter detail (optional, only if the sheets parsed)
    if ins.get('building_rows'):
        tail.append(HRFlowable(width=TW, thickness=0.5, color=BORDER))
        tail.append(Paragraph("Buildings — Sub-Meter Detail", S['sec']))
        b_data = [[Paragraph('Building', S['th']), Paragraph('Panels', S['th']),
                   Paragraph('Total (kWh)', S['th']), Paragraph('Notes', S['th'])]]
        for r in sorted(ins['building_rows'], key=lambda x: -x['total']):
            b_data.append([
                Paragraph(r['label'], S['td']),
                Paragraph(str(r['panels']), S['tdc']),
                Paragraph(f"{r['total']:,.0f}", S['num']),
                Paragraph('Sub-meter, not in EB/DG totals', S['small']),
            ])
        b_data.append([
            Paragraph('<b>Total Sub-Metered</b>', S['bold']), Paragraph('', S['tdc']),
            Paragraph(f"<b>{ins['buildings_total']:,.0f}</b>",
                      style('bld_tot', fontName='Helvetica-Bold', fontSize=8.5,
                            textColor=BLUE, leading=12, alignment=TA_RIGHT)),
            Paragraph('', S['small']),
        ])
        b_tbl = Table(b_data, colWidths=[TW*0.34, TW*0.14, TW*0.20, TW*0.32])
        b_tbl.setStyle(TableStyle([
            ('BACKGROUND',    (0,0),(-1,0),  HexColor('#F3F4F6')),
            ('BACKGROUND',    (0,-1),(-1,-1),HexColor('#EFF6FF')),
            ('ROWBACKGROUNDS',(0,1),(-1,-2), [white, HexColor('#FAFAFA')]),
            ('GRID',          (0,0),(-1,-1), 0.4, BORDER),
            ('TOPPADDING',    (0,0),(-1,-1), 5),
            ('BOTTOMPADDING', (0,0),(-1,-1), 5),
            ('LEFTPADDING',   (0,0),(-1,-1), 8),
            ('RIGHTPADDING',  (0,0),(-1,-1), 8),
            ('VALIGN',        (0,0),(-1,-1), 'MIDDLE'),
        ]))
        tail.append(b_tbl)
        tail.append(Paragraph("Supplementary building panel sub-meters — not included in campus EB/DG totals above.", S['small']))
        tail.append(Spacer(1, 8))

    # Digii Tickets (all-time cumulative — the source file is always the
    # full history, so this table reflects everything to date, not just
    # this reporting month)
    if ins.get('has_ticket_data'):
        tail.append(HRFlowable(width=TW, thickness=0.5, color=BORDER))
        tail.append(Paragraph("Digii Tickets", S['sec']))

        DEPT_COLORS = ['#1B4FD8', '#0891B2', '#059669', '#D97706', '#7C3AED', '#CA8A04', '#DB2777', '#DC2626']
        tk_dept = ins['tickets_by_dept']
        tk_items = list(zip(tk_dept['labels'], tk_dept['vals']))
        tk_total_check = sum(v for _, v in tk_items) or 1
        tk_data = [
            [Paragraph('Problem Type', S['th']), Paragraph('Tickets', S['th']),
             Paragraph('Share', S['th']), Paragraph('Trend Bar', S['th'])],
        ]
        for i, (name, val) in enumerate(tk_items):
            pct = val / tk_total_check * 100
            bar = '█' * int(pct / 3)
            color = DEPT_COLORS[i % len(DEPT_COLORS)]
            tk_data.append([
                Paragraph(name, S['td']),
                Paragraph(f"{int(val):,}", S['num']),
                Paragraph(f"{pct:.1f}%", S['tdc']),
                Paragraph(f'<font color="{color}">{bar}</font>', S['td']),
            ])
        tk_data.append([
            Paragraph('<b>Total</b>', S['bold']),
            Paragraph(f"<b>{ins['ticket_total']:,}</b>", style('tk_tot', fontName='Helvetica-Bold',
                      fontSize=8.5, textColor=INK, leading=12, alignment=TA_RIGHT)),
            Paragraph('100%', S['tdc']),
            Paragraph('', S['td']),
        ])
        tk_tbl = Table(tk_data, colWidths=[TW*0.28, TW*0.18, TW*0.16, TW*0.38])
        tk_tbl.setStyle(TableStyle([
            ('BACKGROUND',    (0,0), (-1,0),  HexColor('#F3F4F6')),
            ('BACKGROUND',    (0,-1),(-1,-1), HexColor('#F0FDF4')),
            ('ROWBACKGROUNDS',(0,1), (-1,-2), [white, HexColor('#FAFAFA')]),
            ('GRID',          (0,0), (-1,-1), 0.4, BORDER),
            ('TOPPADDING',    (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ('LEFTPADDING',   (0,0), (-1,-1), 8),
            ('RIGHTPADDING',  (0,0), (-1,-1), 8),
            ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
        ]))
        tail.append(tk_tbl)
        tat_str = fmt_min_hm(ins['ticket_avg_tat_min']) if ins['ticket_avg_tat_min'] is not None else 'N/A'
        tail.append(Paragraph(
            f"{ins['ticket_pending']} pending · avg resolution time {tat_str} · "
            f"{ins['ticket_recurring_total']} recurring issue(s) flagged (same person, same problem, "
            f"resolved and reopened within 3 days)",
            S['small']
        ))
        tail.append(Spacer(1, 8))

    # Insights
    tail.append(HRFlowable(width=TW, thickness=0.5, color=BORDER))
    tail.append(Paragraph("Insights & Action Items", S['sec']))
    tail.append(Spacer(1, 4))
    for severity, message in ins['flags']:
        tail.append(FlagRow(severity, message, TW))
        tail.append(Spacer(1, 3))
    tail.append(Spacer(1, 8))

    # Outlook (dynamic — generated from ins['outlook'])
    tail.append(HRFlowable(width=TW, thickness=0.5, color=BORDER))
    tail.append(Paragraph(f"{ins['next_month']} Outlook & Action Items", S['sec']))
    tail.append(Spacer(1, 4))
    for item in ins['outlook']:
        tail.append(OutlookRow(item, TW))
        tail.append(Spacer(1, 4))
    tail.append(Spacer(1, 8))

    # Footer
    tail.append(HRFlowable(width=TW, thickness=0.4, color=BORDER))
    tail.append(Spacer(1, 4))
    tail.append(Paragraph(
        f"GL Campus Intelligence · Manamai Campus · Auto-generated {now} · "
        f"Data source: GL_Dashboard_v4_July2026.html",
        S['foot']
    ))

    story.extend(tail)

    doc.build(story)
    return buf.getvalue()


# ── Gmail ──────────────────────────────────────────────────────────────────────

def get_gmail_service():
    """
    Same token.json as gmail_fetcher.authenticate() — never opens a browser
    on an unattended run. If the stored token can't be silently refreshed,
    raise gmail_fetcher.AuthNeedsHumanError immediately rather than blocking
    on run_local_server() forever, since this function is also on the path
    that sends crash/freshness alert emails: a hang here would silently
    swallow the one notification meant to tell a human the token died.
    """
    from gmail_fetcher import AuthNeedsHumanError

    creds = None
    if TOKEN_FILE.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
        except ValueError as e:
            log.warning("token.json is unreadable/incomplete (%s) — treating as missing", e)
            creds = None
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            raise AuthNeedsHumanError(
                "Gmail OAuth needs a human to re-authorize: token.json is missing, "
                "invalid, or has no usable refresh_token. Run gmail_fetcher.py's "
                "bootstrap_token() interactively on a machine with a browser, then "
                "redeploy the resulting token.json to this host."
            )
        TOKEN_FILE.write_text(creds.to_json())
    return build('gmail', 'v1', credentials=creds)


def send_dashboard(recipients: list | None, month: str):
    """Send the latest dashboard HTML file as an email attachment."""
    if not recipients:
        recipients = TO_EMAILS

    html_path = DASHBOARD_FILE
    if not html_path.exists():
        raise FileNotFoundError(f"Dashboard file not found: {html_path}")

    html_bytes = html_path.read_bytes()
    service    = get_gmail_service()
    now        = datetime.now().strftime('%d %b %Y, %I:%M %p')
    filename   = html_path.name

    msg = MIMEMultipart()
    msg['Subject'] = f'GL Campus Intelligence — {month} Dashboard Update'
    msg['From']    = FROM_EMAIL
    msg['To']      = ', '.join(recipients)

    body = MIMEText(
        f"Please find attached the latest GL Campus Intelligence Dashboard for {month}.\n\n"
        f"Last updated: {now}\n"
        f"Campus: Manamai\n\n"
        f"Open the attached HTML file in any browser to view the full interactive dashboard.\n"
        f"Use the 'Download as PDF' button in the sidebar to export a PDF at any time.\n\n"
        f"This dashboard includes:\n"
        f"  • Energy: EB, DG, Solar, Power cuts\n"
        f"  • Facility: AC, Tenant sub-meters, Work orders\n"
        f"  • Water: WTP tanks, RO, New well, Laundromat\n"
        f"  • Historical: Month-on-month comparison\n\n"
        f"— GL Campus Intelligence System",
        'plain'
    )
    msg.attach(body)

    attachment = MIMEBase('text', 'html')
    attachment.set_payload(html_bytes)
    encoders.encode_base64(attachment)
    attachment.add_header('Content-Disposition', f'attachment; filename="{filename}"')
    msg.attach(attachment)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    service.users().messages().send(userId='me', body={'raw': raw}).execute()
    log.info("Dashboard sent to: %s", ', '.join(recipients))


def send_report(pdf_bytes: bytes, recipients: list | None, month: str):
    """Legacy — kept for run_pipeline.py compatibility. Now delegates to send_dashboard."""
    send_dashboard(recipients, month)


# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    dry_run = '--dry-run' in sys.argv
    extra   = []
    if '--to' in sys.argv:
        idx = sys.argv.index('--to')
        if idx + 1 < len(sys.argv):
            extra.append(sys.argv[idx + 1])

    log.info("Loading dashboard data …")
    data = load_dashboard_data()
    log.info("Computing insights …")
    ins  = compute_insights(data)

    if dry_run:
        log.info("Dry run — email not sent. Dashboard file: %s", DASHBOARD_FILE)
        return

    recipients = TO_EMAILS + extra
    log.info("Sending dashboard to: %s", ', '.join(recipients))
    send_dashboard(recipients, ins['month'])
    log.info("Done.")


if __name__ == '__main__':
    main()
