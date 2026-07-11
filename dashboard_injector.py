"""
Phase 3 — Dashboard Injector
Takes parsed Excel data and injects it into the HTML dashboard,
replacing hardcoded JS arrays with fresh values from the Excel file.
"""

import os
import re
import logging
import shutil
from pathlib import Path
from datetime import datetime

log = logging.getLogger(__name__)

BASE_DIR       = Path(__file__).parent
DASHBOARD_FILE = BASE_DIR / "GL_Dashboard_v4_July2026.html"

MONTH_ABBR = {
    1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
    7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"
}


# ── Number formatting ─────────────────────────────────────────────────────────

def fmt(val) -> str:
    """
    Format a number: integers stay as int, floats keep up to 1 decimal.
    None means "no reading was taken" and must stay distinguishable from a
    genuine 0 reading — emit JS null rather than collapsing it to "0".
    """
    if val is None:
        return "null"
    try:
        f = float(val)
        return str(int(f)) if f == int(f) else f"{f:.1f}"
    except (TypeError, ValueError):
        return "0"


def fmt_array(values: list) -> str:
    """Format a list of numbers as a compact JS array string."""
    return "[" + ",".join(fmt(v) for v in values) + "]"


# ── JS block builders ─────────────────────────────────────────────────────────

def build_jun_days(days: list, month_abbr: str) -> str:
    labels = [f"'{d:02d}-{month_abbr}'" for d in days]
    return "const CUR_DAYS = [" + ",".join(labels) + "];"


def build_simple_arrays(data: dict) -> dict:
    """
    Return { var_name: new_js_line } for each simple array variable.
    These are month-agnostic (curEB, curDG, ...) so the dashboard never needs
    a rename when the month rolls over — only the *values* change each run.
    """
    return {
        "curEB":     f"const curEB     = {fmt_array(data['junEB'])};",
        "curDG1":    f"const curDG1    = {fmt_array(data['junDG1'])};",
        "curDG2":    f"const curDG2    = {fmt_array(data['junDG2'])};",
        "curDG3":    f"const curDG3    = {fmt_array(data['junDG3'])};",
        "curDG":     f"const curDG     = {fmt_array(data['junDG'])};",
        "curDiesel": f"const curDiesel = {fmt_array(data['junDiesel'])};",
        "curStock":  f"const curStock  = {fmt_array(data['junStock'])};",
    }


def build_rm_solar_daily(solar: dict) -> str:
    c = fmt_array(solar["canteen"])
    h = fmt_array(solar["hostel"])
    n = fmt_array(solar["newAcad"])
    a = fmt_array(solar["arr"])
    return (
        "const rmSolarDaily = {\n"
        f"  canteen: {c},\n"
        f"  hostel:  {h},\n"
        f"  newAcad: {n},\n"
        f"  arr:     {a}\n"
        "};"
    )


def build_main_solar_daily(solar: dict) -> str:
    p = fmt_array(solar["pgdm"])
    n = fmt_array(solar["newAcad"])
    r = fmt_array(solar["rhs"])
    b = fmt_array(solar["bramahputra"])
    a = fmt_array(solar["adminLhs"])
    return (
        "const mainSolarDaily = {\n"
        f"  pgdm:        {p},\n"
        f"  newAcad:     {n},\n"
        f"  rhs:         {r},\n"
        f"  bramahputra: {b},\n"
        f"  adminLhs:    {a}\n"
        "};"
    )


def build_power_cuts(cuts: list) -> str:
    lines = []
    for c in cuts:
        lines.append(
            f"  {{date:'{c['date']}',dur:{c['dur']},"
            f"events:'{c['events']}',dg:'{c['dg']}'}}"
        )
    inner = ",\n".join(lines)
    return f"const curPowerCuts = [\n{inner},\n];"


def build_power_cuts_by_month(by_month: dict) -> str:
    """
    Every month's individual outage events, keyed 'YYYY-MM', so Longest
    Outage / Avg Cut can be computed for ANY month with event-level history
    in the source sheet — not just the live month and the one hardcoded
    Feb-2026 backfill. Same event shape (date/start/end/dur/events/dg) as
    curPowerCuts/powerCutsFeb, so existing drill-down code needs no changes.
    """
    month_entries = []
    for month_key in sorted(by_month.keys()):
        events = by_month[month_key]
        event_json = ",".join(
            "{" +
            f"date:{fmt_str(e['date'])},start:{fmt_str(e['start'])},end:{fmt_str(e['end'])},"
            f"dur:{e['dur']},events:{fmt_str(e['events'])},dg:{fmt_str(e['dg'])}"
            + "}"
            for e in events
        )
        month_entries.append(f"{fmt_str(month_key)}:[{event_json}]")
    return "const powerCutsByMonth = {" + ",".join(month_entries) + "};"


def _js_escape_structural(s: str) -> str:
    """
    Escape every character replace_var's balanced-bracket scanner treats as
    structurally significant — '[', ']', '{', '}', ';' — using JS \\uXXXX
    escapes, so the ENCODED string literal never contains a raw one of these
    characters. Without this, free-text Excel fields (an assignee name, a
    department rename, a pasted remark) containing a stray bracket/brace can
    close the scanner's balanced match one token early and splice stale
    leftover JS into the next run's injected block — a real, reproduced
    corruption path, not a theoretical one. Also escapes '<' defensively
    (an inline </script> substring would end the containing <script> tag).
    """
    return (
        str(s)
        .replace("\\", "\\\\")
        .replace("'", "\\'")
        .replace("[", "\\u005B").replace("]", "\\u005D")
        .replace("{", "\\u007B").replace("}", "\\u007D")
        .replace(";", "\\u003B")
        .replace("<", "\\u003C")
        .replace("\n", "\\n").replace("\r", "")
    )


def fmt_str(s: str) -> str:
    return "'" + _js_escape_structural(s) + "'"


def fmt_str_array(values: list) -> str:
    return "[" + ",".join(fmt_str(v) for v in values) + "]"


def inject_tickets(html: str, data: dict) -> str:
    """
    Digii Tickets — all-time cumulative counts from the service-request
    export (not month-indexed like everything else; the source file is
    always the full history, so we just re-emit fresh totals each run).
    """
    by_dept = data.get("ticketsByDept") or {}
    by_level = data.get("ticketsByLevel") or {}
    dept_by_level = data.get("ticketsDeptByLevel") or {}

    dept_labels = sorted(by_dept.keys(), key=lambda d: -by_dept[d])
    dept_vals = [by_dept[d] for d in dept_labels]

    levels = ["Level 1", "Level 2", "Level 3"]
    level_vals = [by_level.get(lv, 0) for lv in levels]

    dept_json = ",".join(
        f"{{dept:{fmt_str(d)},l1:{dept_by_level.get(d, {}).get('Level 1', 0)},"
        f"l2:{dept_by_level.get(d, {}).get('Level 2', 0)},"
        f"l3:{dept_by_level.get(d, {}).get('Level 3', 0)}}}"
        for d in dept_labels
    )

    html = replace_var(html, "ticketsByDept",
        f"const ticketsByDept = {{labels:{fmt_str_array(dept_labels)},vals:{fmt_array(dept_vals)}}};")
    html = replace_var(html, "ticketsByLevel",
        f"const ticketsByLevel = {{labels:{fmt_str_array(levels)},vals:{fmt_array(level_vals)}}};")
    html = replace_var(html, "ticketsDeptByLevel", f"const ticketsDeptByLevel = [{dept_json}];")
    html = replace_var(html, "ticketTotal", f"const ticketTotal = {data.get('ticketTotal', 0)};")
    html = replace_var(html, "ticketPending", f"const ticketPending = {data.get('ticketPending', 0) or 0};")
    avg_tat = data.get("ticketAvgTatMin")
    html = replace_var(html, "ticketAvgTatMin", f"const ticketAvgTatMin = {avg_tat if avg_tat is not None else 'null'};")

    rows = data.get("ticketRows") or []
    row_json = ",".join(
        "{" +
        f"id:{fmt_str(r['id']) if r.get('id') is not None else 'null'},"
        f"dept:{fmt_str(r['dept'])},"
        f"status:{fmt_str(r['status'])},"
        f"date:{fmt_str(r['date']) if r.get('date') else 'null'},"
        f"tat:{r['tat'] if r.get('tat') is not None else 'null'},"
        f"level:{fmt_str(r['level'])},"
        f"assignee:{fmt_str(r['assignee']) if r.get('assignee') else 'null'}"
        + "}"
        for r in rows
    )
    html = replace_var(html, "ticketRows", f"const ticketRows = [{row_json}];")

    recurring = data.get("ticketRecurring") or []
    recurring_json = ",".join(
        "{" +
        f"name:{fmt_str(r['name'])},"
        f"dept:{fmt_str(r['dept'])},"
        f"resolvedDate:{fmt_str(r['resolvedDate'])},"
        f"reopenedDate:{fmt_str(r['reopenedDate'])},"
        f"gapDays:{r['gapDays']},"
        f"priorId:{fmt_str(r['priorId']) if r.get('priorId') is not None else 'null'},"
        f"newId:{fmt_str(r['newId']) if r.get('newId') is not None else 'null'}"
        + "}"
        for r in recurring
    )
    html = replace_var(html, "ticketRecurring", f"const ticketRecurring = [{recurring_json}];")

    recurring_by_dept = data.get("ticketRecurringByDept") or {}
    rbd_labels = sorted(recurring_by_dept.keys(), key=lambda d: -recurring_by_dept[d])
    rbd_vals = [recurring_by_dept[d] for d in rbd_labels]
    html = replace_var(html, "ticketRecurringByDept",
        f"const ticketRecurringByDept = {{labels:{fmt_str_array(rbd_labels)},vals:{fmt_array(rbd_vals)}}};")

    # This-month overview — same all-time sheet, bucketed by Request Date so
    # the dashboard can show "raised this month" alongside the all-time totals.
    this_month = data.get("ticketThisMonth") or {"total": 0, "pending": 0, "resolved": 0, "avgTatMin": None, "byDept": {}}
    tm_by_dept = this_month.get("byDept") or {}
    tm_labels = sorted(tm_by_dept.keys(), key=lambda d: -tm_by_dept[d])
    tm_vals = [tm_by_dept[d] for d in tm_labels]
    tm_avg_tat = this_month.get("avgTatMin")
    html = replace_var(html, "ticketThisMonth",
        f"const ticketThisMonth = {{total:{this_month.get('total', 0)},"
        f"pending:{this_month.get('pending', 0)},"
        f"resolved:{this_month.get('resolved', 0)},"
        f"avgTatMin:{tm_avg_tat if tm_avg_tat is not None else 'null'},"
        f"byDept:{{labels:{fmt_str_array(tm_labels)},vals:{fmt_array(tm_vals)}}}}};")
    html = replace_var(html, "ticketThisMonthKey",
        f"const ticketThisMonthKey = {fmt_str(data.get('ticketThisMonthKey', ''))};")

    # Monthly trend — total tickets raised per month across every month the
    # export has history for (currently Jul 2025 onward), sorted chronologically.
    ticket_months = data.get("ticketMonths") or {}
    month_keys_sorted = sorted(ticket_months.keys())
    trend_json = ",".join(
        f"{{month:{fmt_str(mk)},total:{ticket_months[mk].get('total', 0)},"
        f"pending:{ticket_months[mk].get('pending', 0)},"
        f"resolved:{ticket_months[mk].get('resolved', 0)}}}"
        for mk in month_keys_sorted
    )
    html = replace_var(html, "ticketMonthlyTrend", f"const ticketMonthlyTrend = [{trend_json}];")

    log.info("Tickets injected: %d total (%d pending) across %d departments, %d rows, %d recurring issues, "
              "%d this month",
              data.get("ticketTotal", 0), data.get("ticketPending", 0) or 0, len(dept_labels), len(rows), len(recurring),
              this_month.get("total", 0))
    return html


def build_wtp_daily(wtp: dict) -> str:
    w1 = fmt_array(wtp["wtp1"])
    w2 = fmt_array(wtp["wtp2"])
    w3 = fmt_array(wtp["wtp3"])
    return (
        "const wtpDaily = {\n"
        f"  wtp1:{w1},\n"
        f"  wtp2:{w2},\n"
        f"  wtp3:{w3}\n"
        "};"
    )


def inject_wtp(html: str, data: dict) -> str:
    """Inject all water-related arrays from WTP parse data."""
    wtp = data.get("wtpDaily", {})
    if wtp.get("wtp1"):
        html = replace_var(html, "wtpDaily", build_wtp_daily(wtp))

    for key in ["newWellJunKL", "laundryDaily", "canteenFlowDaily", "sathiyaDailyKL",
                "acOutletTotal", "stpFlow1", "stpFlow2", "roCansDaily"]:
        val = data.get(key)
        if val is not None:
            html = replace_var(html, key, f"const {key} = {fmt_array(val)};")

    for key in ["wellAcre2", "wellAcre6", "tdsRo1000", "tdsRo2000"]:
        val = data.get(key)
        if val is not None:
            arr = "[" + ",".join("null" if v is None else fmt(v) for v in val) + "]"
            html = replace_var(html, key, f"const {key} = {arr};")

    blocks = data.get("acOutletBlocks")
    if blocks:
        inner = ",".join(f"[{','.join(fmt(v) for v in b)}]" for b in blocks)
        html = replace_var(html, "acOutletBlocks", f"const acOutletBlocks = [{inner}];")

    return html


# ── Regex replacers ───────────────────────────────────────────────────────────

def replace_var(html: str, var_name: str, new_block: str) -> str:
    """
    Replace a JS variable declaration in the HTML.
    Handles flat arrays, nested arrays ([[...]]), objects, and arrays with null values.
    Matches: const varName = <value>;
    where <value> is a balanced [ ] or { } expression.
    """
    # Use a balanced-bracket approach via a simple scanner instead of regex
    # Allow any whitespace between var_name and '='
    import re as _re
    m = _re.search(rf'const\s+{_re.escape(var_name)}\s*=', html)
    if not m:
        log.warning("Variable not found in HTML: %s", var_name)
        return html
    idx = m.start()

    # Find the start of the value (skip whitespace after '=')
    val_start = m.end()
    while val_start < len(html) and html[val_start] in (' ', '\t', '\n'):
        val_start += 1

    opener = html[val_start]
    if opener not in ('[', '{'):
        # Scalar value (number/string/bool) — ends at the next semicolon.
        end = html.index(';', val_start) + 1
    else:
        closer = ']' if opener == '[' else '}'
        depth = 0
        pos = val_start
        while pos < len(html):
            if html[pos] == opener:
                depth += 1
            elif html[pos] == closer:
                depth -= 1
                if depth == 0:
                    break
            pos += 1

        # pos is now at the closing bracket; skip optional semicolon
        end = pos + 1
        if end < len(html) and html[end] == ';':
            end += 1

    new_html = html[:idx] + new_block + html[end:]
    log.info("Updated: %s", var_name)
    return new_html


def replace_scalar_var(html: str, var_name: str, new_value) -> str:
    """Replace a primitive `const varName = <number>;` declaration."""
    pattern = rf'const\s+{re.escape(var_name)}\s*=\s*[^;]+;'
    new_html, n = re.subn(pattern, f"const {var_name}={new_value};", html)
    if n == 0:
        log.warning("Variable not found in HTML: %s", var_name)
    else:
        log.info("Updated: %s", var_name)
    return new_html


def replace_jun_days(html: str, new_line: str) -> str:
    pattern = r"const CUR_DAYS\s*=\s*\[[\s\S]*?\];"
    new_html, n = re.subn(pattern, new_line, html)
    if n == 0:
        log.warning("CUR_DAYS not found in HTML")
    else:
        log.info("Updated: CUR_DAYS")
    return new_html


# ── Monthly totals updater ────────────────────────────────────────────────────

def _extract_flat_array(html: str, var_name: str) -> list:
    """Extract a flat numeric JS array from the HTML as a Python list."""
    m = re.search(rf'const\s+{re.escape(var_name)}\s*=\s*\[([^\]]*)\]', html)
    if not m:
        return []
    parts = [p.strip() for p in m.group(1).split(',') if p.strip()]
    result = []
    for p in parts:
        try:
            result.append(float(p))
        except ValueError:
            result.append(0.0)
    return result


def _extract_nested_totals(html: str, var_name: str) -> list:
    """Extract the 'totals' sub-array from a nested object like wtpMonthly or rmSolarMonthly."""
    m = re.search(rf'const\s+{re.escape(var_name)}\s*=\s*\{{[^}}]*totals:\[([^\]]*)\]', html)
    if not m:
        return []
    parts = [p.strip() for p in m.group(1).split(',') if p.strip()]
    result = []
    for p in parts:
        try:
            result.append(float(p))
        except ValueError:
            result.append(0.0)
    return result


def _set_index(lst: list, idx: int, val) -> list:
    """Return a new list with lst[idx] replaced by val."""
    out = list(lst)
    while len(out) <= idx:
        out.append(0)
    out[idx] = val
    return out


def _extract_hist_year_block(html: str, year: int) -> tuple:
    """Return (start, end, body) spanning 'YEAR:{...}' inside HIST_DATA, or None."""
    m = re.search(rf'(?<![\d.]){year}:\{{', html)
    if not m:
        return None
    depth = 0
    pos = m.end() - 1  # position of the opening '{'
    start_brace = pos
    while pos < len(html):
        if html[pos] == '{':
            depth += 1
        elif html[pos] == '}':
            depth -= 1
            if depth == 0:
                break
        pos += 1
    return m.start(), pos + 1, html[start_brace:pos + 1]


def _extract_hist_field(body: str, field: str) -> list:
    m = re.search(rf'{field}:\[([^\]]*)\]', body)
    if not m:
        return []
    out = []
    for p in [x.strip() for x in m.group(1).split(',')]:
        out.append(None if p == 'null' else float(p))
    return out


def _fmt_hist_array(values: list, decimals: int = None) -> str:
    parts = []
    for v in values:
        if v is None:
            parts.append('null')
        elif decimals is not None:
            parts.append(f"{v:.{decimals}f}" if v != int(v) or decimals else str(int(v)))
        else:
            parts.append(fmt(v))
    return "[" + ",".join(parts) + "]"


def update_hist_data(html: str, yr: int, mon_idx: int, eb, dg, diesel, days) -> str:
    """
    Patch HIST_DATA[yr] in place for the given month index so the Historical
    tab stays current without manual edits. Leaves other years untouched.
    """
    span = _extract_hist_year_block(html, yr)
    if not span:
        log.warning("HIST_DATA[%d] not found in HTML — skipping historical sync", yr)
        return html
    start, end, body = span

    eb_arr     = _set_index(_extract_hist_field(body, "eb"),     mon_idx, eb)
    dg_arr     = _set_index(_extract_hist_field(body, "dg"),     mon_idx, dg)
    diesel_arr = _set_index(_extract_hist_field(body, "diesel"), mon_idx, diesel)
    days_arr   = _set_index(_extract_hist_field(body, "days"),   mon_idx, days)

    note_m = re.search(r'note:\[([^\]]*)\]', body)
    notes = [x.strip() for x in note_m.group(1).split(',')] if note_m else []
    while len(notes) <= mon_idx:
        notes.append("''")
    notes[mon_idx] = "''"

    new_body = (
        "{eb:"     + _fmt_hist_array(eb_arr) +
        ",dg:"     + _fmt_hist_array(dg_arr) +
        ",diesel:" + _fmt_hist_array(diesel_arr, decimals=1) +
        ",days:"   + _fmt_hist_array(days_arr) +
        ",note:["  + ",".join(notes) + "]}"
    )
    html = html[:start] + f"{yr}:{new_body}" + html[end:]
    log.info("HIST_DATA[%d][%d] updated (eb=%s dg=%s diesel=%s days=%s)", yr, mon_idx, eb, dg, diesel, days)
    return html


def update_monthly_totals(html: str, data: dict) -> str:
    """
    Update all monthly-total arrays in the HTML for the current month.
    Sums the freshly-injected daily arrays and patches the correct index
    in each monthly array (index = month_number - 1).
    """
    mon_idx = data["month"] - 1   # 0-based; June=5

    # ── ebMonthly ─────────────────────────────────────────────────────────────
    eb_total = round(sum(data.get("junEB", [])))
    if eb_total:
        arr = _set_index(_extract_flat_array(html, "ebMonthly"), mon_idx, eb_total)
        html = replace_var(html, "ebMonthly", f"const ebMonthly={fmt_array(arr)};")
        log.info("ebMonthly[%d] = %s", mon_idx, eb_total)

        # ── ebAvgDaily (average EB per active day, same index as ebMonthly) ────
        active_days = sum(1 for v in data.get("junEB", []) if v)
        eb_avg = round(eb_total / active_days) if active_days else 0
        avg_arr = _set_index(_extract_flat_array(html, "ebAvgDaily"), mon_idx, eb_avg)
        html = replace_var(html, "ebAvgDaily", f"const ebAvgDaily={fmt_array(avg_arr)};")
        log.info("ebAvgDaily[%d] = %s", mon_idx, eb_avg)

    # ── dgMonthly ─────────────────────────────────────────────────────────────
    dg_total = round(sum(data.get("junDG", [])))
    if dg_total:
        arr = _set_index(_extract_flat_array(html, "dgMonthly"), mon_idx, dg_total)
        html = replace_var(html, "dgMonthly", f"const dgMonthly={fmt_array(arr)};")
        log.info("dgMonthly[%d] = %s", mon_idx, dg_total)

    # ── dieselMonthly ─────────────────────────────────────────────────────────
    diesel_total = round(sum(data.get("junDiesel", [])), 1)
    if diesel_total:
        arr = _set_index(_extract_flat_array(html, "dieselMonthly"), mon_idx, diesel_total)
        html = replace_var(html, "dieselMonthly", f"const dieselMonthly={fmt_array(arr)};")
        log.info("dieselMonthly[%d] = %s", mon_idx, diesel_total)

    # ── cutMonthly ────────────────────────────────────────────────────────────
    cut_total = sum(c.get("dur", 0) for c in data.get("powerCuts", []))
    if cut_total:
        arr = _set_index(_extract_flat_array(html, "cutMonthly"), mon_idx, cut_total)
        html = replace_var(html, "cutMonthly", f"const cutMonthly={fmt_array(arr)};")
        log.info("cutMonthly[%d] = %s", mon_idx, cut_total)

    # ── wtpMonthly ────────────────────────────────────────────────────────────
    wtp = data.get("wtpDaily", {})
    if wtp.get("wtp1"):
        wtp_keys = ["wtp1", "wtp2", "wtp3"]
        # Read all four sub-arrays from current HTML
        m = re.search(r'const\s+wtpMonthly\s*=\s*\{([^}]*)\}', html)
        if m:
            body = m.group(1)
            sub = {}
            for k in wtp_keys + ["totals"]:
                sm = re.search(rf'{k}:\[([^\]]*)\]', body)
                sub[k] = [float(x.strip()) for x in sm.group(1).split(',') if x.strip()] if sm else []
            for k in wtp_keys:
                t = round(sum(wtp.get(k, [])), 1)
                sub[k] = _set_index(sub[k], mon_idx, t)
            # Recompute totals as sum of the three
            sub["totals"] = [round(sub["wtp1"][i] + sub["wtp2"][i] + sub["wtp3"][i], 1)
                             for i in range(len(sub["wtp1"]))]
            new_block = (
                f"const wtpMonthly={{wtp1:{fmt_array(sub['wtp1'])},"
                f"wtp2:{fmt_array(sub['wtp2'])},"
                f"wtp3:{fmt_array(sub['wtp3'])},"
                f"totals:{fmt_array(sub['totals'])}}};")
            html = replace_var(html, "wtpMonthly", new_block)
            log.info("wtpMonthly[%d] updated", mon_idx)

    # ── laundryMonthly ────────────────────────────────────────────────────────
    laundry_total = round(sum(data.get("laundryDaily", [])), 1)
    if laundry_total:
        arr = _set_index(_extract_flat_array(html, "laundryMonthly"), mon_idx, laundry_total)
        html = replace_var(html, "laundryMonthly", f"const laundryMonthly={fmt_array(arr)};")
        log.info("laundryMonthly[%d] = %s", mon_idx, laundry_total)

    # ── newWellMonthly ────────────────────────────────────────────────────────
    well_total = round(sum(data.get("newWellJunKL", [])), 1)
    arr = _set_index(_extract_flat_array(html, "newWellMonthly"), mon_idx, well_total)
    html = replace_var(html, "newWellMonthly", f"const newWellMonthly={fmt_array(arr)};")
    log.info("newWellMonthly[%d] = %s", mon_idx, well_total)

    # ── sathiyaMonthly ────────────────────────────────────────────────────────
    if "sathiyaDailyKL" in data:
        sathiya_total = round(sum(data.get("sathiyaDailyKL", [])), 1)
        arr = _set_index(_extract_flat_array(html, "sathiyaMonthly"), mon_idx, sathiya_total)
        html = replace_var(html, "sathiyaMonthly", f"const sathiyaMonthly={fmt_array(arr)};")
        log.info("sathiyaMonthly[%d] = %s", mon_idx, sathiya_total)

    # ── rmSolarMonthly ────────────────────────────────────────────────────────
    rm = data.get("rmSolarDaily", {})
    if rm.get("canteen"):
        rm_keys = ["canteen", "hostel", "newAcad", "arr"]
        m = re.search(r'const\s+rmSolarMonthly\s*=\s*\{([^}]*)\}', html)
        if m:
            body = m.group(1)
            sub = {}
            for k in rm_keys + ["totals"]:
                sm = re.search(rf'{k}:\[([^\]]*)\]', body)
                sub[k] = [float(x.strip()) for x in sm.group(1).split(',') if x.strip()] if sm else []
            for k in rm_keys:
                t = round(sum(rm.get(k, [])), 1)
                sub[k] = _set_index(sub[k], mon_idx, t)
            sub["totals"] = [round(sub["canteen"][i] + sub["hostel"][i] + sub["newAcad"][i] + sub["arr"][i], 1)
                             for i in range(len(sub["canteen"]))]
            new_block = (
                f"const rmSolarMonthly={{canteen:{fmt_array(sub['canteen'])},"
                f"hostel:{fmt_array(sub['hostel'])},"
                f"newAcad:{fmt_array(sub['newAcad'])},"
                f"arr:{fmt_array(sub['arr'])},"
                f"totals:{fmt_array(sub['totals'])}}};")
            html = replace_var(html, "rmSolarMonthly", new_block)
            log.info("rmSolarMonthly[%d] updated", mon_idx)

    # ── HIST_DATA (year-over-year historical comparison) ─────────────────────
    eb_total_h = round(sum(data.get("junEB", [])))
    if eb_total_h:
        active_days = sum(1 for v in data.get("junEB", []) if v)
        html = update_hist_data(html, yr=data["year"], mon_idx=mon_idx,
                                 eb=eb_total_h,
                                 dg=round(sum(data.get("junDG", []))),
                                 diesel=round(sum(data.get("junDiesel", [])), 1),
                                 days=active_days)

    # ── mainSolarMonthly ──────────────────────────────────────────────────────
    ms = data.get("mainSolarDaily", {})
    if ms.get("pgdm"):
        ms_keys = ["pgdm", "newAcad", "rhs", "bramahputra", "adminLhs"]
        m = re.search(r'const\s+mainSolarMonthly\s*=\s*\{([^}]*)\}', html)
        if m:
            body = m.group(1)
            sub = {}
            for k in ms_keys + ["totals"]:
                sm = re.search(rf'{k}:\[([^\]]*)\]', body)
                sub[k] = [float(x.strip()) for x in sm.group(1).split(',') if x.strip()] if sm else []
            for k in ms_keys:
                t = round(sum(ms.get(k, [])), 1)
                sub[k] = _set_index(sub[k], mon_idx, t)
            sub["totals"] = [round(sum(sub[k][i] for k in ms_keys), 1)
                             for i in range(len(sub["pgdm"]))]
            new_block = (
                f"const mainSolarMonthly={{pgdm:{fmt_array(sub['pgdm'])},"
                f"newAcad:{fmt_array(sub['newAcad'])},"
                f"rhs:{fmt_array(sub['rhs'])},"
                f"bramahputra:{fmt_array(sub['bramahputra'])},"
                f"adminLhs:{fmt_array(sub['adminLhs'])},"
                f"totals:{fmt_array(sub['totals'])}}};")
            html = replace_var(html, "mainSolarMonthly", new_block)
            log.info("mainSolarMonthly[%d] updated", mon_idx)

    return html


def _js_str(s: str) -> str:
    """
    Escape a Python string for embedding in a single-quoted JS string literal.
    Delegates to _js_escape_structural so brackets/braces/semicolons from
    free-text Excel fields can't break replace_var's balanced-bracket scanner
    (see that function's docstring) — this and fmt_str must stay in sync.
    """
    return _js_escape_structural(s)


def inject_facility(html: str, data: dict) -> str:
    """Inject Manpower, Street Lights, and AMC sections — previously hardcoded
    in the dashboard HTML, now sourced from the 'Man Power', 'Street light
    Data', and 'AMC' sheets of the electrical report."""
    depts = data.get("mpDepts")
    if depts:
        inner = ",".join(f"{{name:'{_js_str(d['name'])}',n:{int(d['n'])}}}" for d in depts)
        html = replace_var(html, "mpDepts", f"const mpDepts=[{inner}];")

    shifts = data.get("mpShifts")
    if shifts:
        labels = ",".join(f"'{_js_str(l)}'" for l in shifts["labels"])
        vals = ",".join(str(int(v)) for v in shifts["vals"])
        html = replace_var(html, "mpShifts", f"const mpShifts={{labels:[{labels}],vals:[{vals}]}};")

    if "mpTotalStrength" in data:
        html = replace_var(html, "MP_STATS",
            f"const MP_STATS={{total:{int(data['mpTotalStrength'])},"
            f"active:{int(data['mpActive'])},wc:{int(data['mpWc'])},leave:{int(data['mpLeave'])}}};")

    lights = data.get("streetLights")
    if lights:
        inner = ",".join(
            f"{{type:'{_js_str(l['type'])}',total:{l['total']},ok:{l['ok']},fail:{l['fail']}}}"
            for l in lights
        )
        html = replace_var(html, "streetLights", f"const streetLights=[{inner}];")

    amc = data.get("amcContracts")
    if amc:
        inner = ",".join(
            f"{{system:'{_js_str(a['system'])}',vendor:'{_js_str(a['vendor'])}',"
            f"last:'{_js_str(a['last'])}',next:'{_js_str(a['next'])}',status:'{_js_str(a['status'])}'}}"
            for a in amc
        )
        html = replace_var(html, "amcContracts", f"const amcContracts=[{inner}];")

    if depts or shifts or lights or amc:
        log.info("Facility sections updated: manpower=%s streetlights=%s amc=%s",
                  bool(depts), bool(lights), bool(amc))
    return html


def _extract_existing_building_keys(html: str) -> set:
    """
    Return the set of building keys already present in the HTML's BUILDINGS
    object, without fully parsing it — just enough to know what NOT to lose.
    """
    m = re.search(r'const\s+BUILDINGS\s*=\s*\{', html)
    if not m:
        return set()
    start = m.end() - 1
    depth, pos = 0, start
    while pos < len(html):
        if html[pos] == '{': depth += 1
        elif html[pos] == '}':
            depth -= 1
            if depth == 0: break
        pos += 1
    body = html[start:pos + 1]
    return set(re.findall(r"(\w+):\{days:\[", body))


def inject_buildings(html: str, data: dict) -> str:
    """
    Inject the per-building sub-meter panel data (Buildings section).
    This REPLACES the whole BUILDINGS object each run, so if this run's
    source file is missing a sheet that a previous, more complete run had
    (e.g. an older-layout electrical report got picked up), that building
    would silently vanish from the dashboard. Guard against that: only
    proceed if this run's building set is not a strict regression — i.e. it
    covers at least the keys already on disk, or there simply weren't any
    keys on disk yet.
    """
    buildings = data.get("buildings")
    if not buildings:
        return html

    existing_keys = _extract_existing_building_keys(html)
    missing = existing_keys - buildings.keys()
    if missing:
        log.warning(
            "BUILDINGS: this run's source file is missing sheets for %s (present on disk from a "
            "prior run) — keeping those untouched instead of deleting them.",
            ", ".join(sorted(missing)),
        )
        # Can't recover the old per-building data from the new `data` dict (it
        # was never parsed this run), so we leave those keys out of the
        # rewritten object only if we truly have no way to preserve them.
        # Safer default: refuse to shrink the building set at all.
        return html

    parts = []
    for key, b in buildings.items():
        days = ",".join(str(d) for d in b["days"])
        panels = ",".join(
            f"{{name:'{_js_str(p['name'])}',daily:{fmt_array(p['daily'])}}}"
            for p in b["panels"]
        )
        months = b.get("months", {})
        month_keys_sorted = sorted(months.keys())
        monthly_totals = ",".join(
            f"'{mk}':{fmt(months[mk]['total'])}" for mk in month_keys_sorted
        )
        parts.append(
            f"{key}:{{days:[{days}],panels:[{panels}],monthlyTotals:{{{monthly_totals}}}}}"
        )

    new_block = "const BUILDINGS={" + ",".join(parts) + "};"
    html = replace_var(html, "BUILDINGS", new_block)
    log.info("BUILDINGS updated: %s", ", ".join(buildings.keys()))
    return html


def inject_chair_counts(html: str, data: dict) -> str:
    """Inject the classroom/hall chair inventory (Facilities > Chair Count)."""
    chair_counts = data.get("chairCounts")
    if not chair_counts:
        return html

    parts = []
    for key, c in chair_counts.items():
        halls = ",".join(
            f"{{name:'{_js_str(h['name'])}',wch:{h['wch']},fch:{h['fch']},"
            f"padch:{h['padch']},total:{h['total']}}}"
            for h in c["halls"]
        )
        parts.append(f"{key}:{{label:'{_js_str(c['label'])}',asOf:'{_js_str(c['asOf'])}',halls:[{halls}]}}")

    new_block = "const CHAIR_COUNTS={" + ",".join(parts) + "};"
    html = replace_var(html, "CHAIR_COUNTS", new_block)
    log.info("CHAIR_COUNTS updated: %s", ", ".join(chair_counts.keys()))
    return html


# ── Main injector ─────────────────────────────────────────────────────────────

def inject_report_archive(html: str) -> str:
    """
    Inject the Past Reports list (archive/report_archive.json), embedding each
    PDF's bytes as base64 directly in the HTML. This is required because the
    dashboard is shared as a single HTML file (email attachment) — a relative
    'archive/xxx.pdf' link only resolves on the machine that has that folder
    on disk, so anyone else opening the emailed HTML gets a broken link.
    Embedding trades ~33% size overhead per PDF for the file being genuinely
    self-contained and working identically on any device.
    """
    import json
    import base64
    archive_dir = BASE_DIR / "archive"
    index_file = archive_dir / "report_archive.json"
    if not index_file.exists():
        return html
    entries = json.loads(index_file.read_text(encoding="utf-8"))
    parts = []
    for e in entries:
        pdf_data_js = "null"
        if e.get("file"):
            pdf_path = archive_dir / e["file"]
            if pdf_path.exists():
                b64 = base64.b64encode(pdf_path.read_bytes()).decode("ascii")
                pdf_data_js = f"'{b64}'"
            else:
                log.warning("Archived PDF missing on disk, skipping embed: %s", e["file"])
        parts.append(
            "{month:%d,year:%d,label:'%s',abbr:'%s',file:%s,pdfData:%s,placeholder:%s}" % (
                e["month"], e["year"], e["label"], e["abbr"],
                f"'{e['file']}'" if e.get("file") else "null",
                pdf_data_js,
                "true" if e.get("placeholder") else "false",
            )
        )
    new_block = "const REPORT_ARCHIVE=[" + ",".join(parts) + "];"
    html = replace_var(html, "REPORT_ARCHIVE", new_block)
    log.info("REPORT_ARCHIVE updated: %d entries (PDFs embedded as base64)", len(entries))
    return html


def inject(data: dict, dashboard_path: str = None) -> str:
    """
    Inject parsed data into the dashboard HTML.
    Returns the path of the updated file.
    """
    path    = Path(dashboard_path) if dashboard_path else DASHBOARD_FILE
    mon_num = data["month"]
    yr      = data["year"]
    abbr    = MONTH_ABBR[mon_num]   # e.g. "Jun"

    # Backup original before modifying
    backup = path.with_suffix(f".backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
    shutil.copy2(path, backup)
    log.info("Backup saved: %s", backup.name)

    html = path.read_text(encoding="utf-8")
    original_len = len(html)

    # 1 — Day labels
    html = replace_jun_days(html, build_jun_days(data["days"], abbr))

    # 2 — Simple daily arrays
    for var_name, new_line in build_simple_arrays(data).items():
        html = replace_var(html, var_name, new_line)

    # 3 — RM Solar daily
    html = replace_var(html, "rmSolarDaily", build_rm_solar_daily(data["rmSolarDaily"]))

    # 3b — Main Solar daily
    if data.get("mainSolarDaily", {}).get("pgdm"):
        html = replace_var(html, "mainSolarDaily", build_main_solar_daily(data["mainSolarDaily"]))

    # 4 — Power cuts
    html = replace_var(html, "curPowerCuts", build_power_cuts(data["powerCuts"]))
    if data.get("powerCutsByMonth"):
        html = replace_var(html, "powerCutsByMonth", build_power_cuts_by_month(data["powerCutsByMonth"]))

    # 5 — Digii Tickets (all-time cumulative, re-emitted fresh each run)
    if data.get("ticketsByDept"):
        html = inject_tickets(html, data)

    # 7 — WTP / water data (injected only when wtp data is present in payload)
    if "wtpDaily" in data:
        html = inject_wtp(html, data)

    # 8 — Electrical sub-meters (EV, tenants, laundromat elec)
    for key in ["evDailyKWh", "nescafeDailyKWh", "cvbDailyKWh", "teaDailyKWh",
                "yummyDailyKWh", "laundryElecKWh"]:
        val = data.get(key)
        if val is not None:
            html = replace_var(html, key, f"const {key} = {fmt_array(val)};")

    # 9 — Monthly totals (summed from the daily arrays just injected above)
    html = update_monthly_totals(html, data)

    # 10 — Current year (drives page title, historical tab labels, etc.)
    html = replace_scalar_var(html, "CUR_YEAR", yr)

    # 11 — Facility sections (Manpower, Street Lights, AMC)
    html = inject_facility(html, data)

    # 12 — Buildings (per-building sub-meter panels)
    html = inject_buildings(html, data)

    # 13 — Chair Count (classroom/hall seating inventory)
    html = inject_chair_counts(html, data)

    # 14 — Past Reports archive list
    html = inject_report_archive(html)

    # Write to a temp file in the same directory, then atomically swap it in.
    # A direct path.write_text() truncates the live file before writing the
    # new content — if the process is killed mid-write (Task Scheduler
    # timeout, host OOM-kill, forced reboot), the live dashboard is left
    # truncated/corrupt and stays that way until a human notices and manually
    # restores a .backup_*.html. os.replace() is a single atomic rename on
    # both Windows and POSIX, so the live file is always either the old
    # complete version or the new complete version — never a partial write.
    tmp_path = path.with_suffix(path.suffix + f".tmp{datetime.now().strftime('%Y%m%d_%H%M%S%f')}")
    tmp_path.write_text(html, encoding="utf-8")
    os.replace(str(tmp_path), str(path))

    log.info("Dashboard updated: %s  (%d chars -> %d chars)", path.name, original_len, len(html))
    return str(path)


if __name__ == "__main__":
    import sys
    sys.path.insert(0, str(BASE_DIR))
    from excel_parser import parse_electrical_file

    logging.basicConfig(level=logging.INFO, format="%(levelname)s  %(message)s")

    elec_file = BASE_DIR / "downloads" / "Electrical, AC, Carpentry, STP & Plumbing Daily Report June 2026.xlsx"

    log.info("=== Phase 3 — Dashboard Injector ===")
    data = parse_electrical_file(str(elec_file))
    out  = inject(data)
    log.info("Done. Open: %s", out)
