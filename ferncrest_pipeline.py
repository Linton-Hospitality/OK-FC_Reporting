#!/usr/bin/env python3
"""
Ferncrest Data Pipeline
=======================
Reads Cloudbeds occupancy Excel exports and writes actuals + pace data
into the property workbook (ferncrest_ok-fc_v2.xlsx).

Usage:
    python ferncrest_pipeline.py

File expectations (set paths in CONFIG below):
    - ok-fc_occupancy_2026ytd.xlsx   : Jan 1 – Apr 30 actuals
    - ok-fc_occupancy_pace_90day.xlsx: May 1 – Jul 30 forward pace
    - ferncrest_ok-fc_v2.xlsx        : property workbook

What it does:
    1. Parses both Cloudbeds exports
    2. Aggregates daily rows → monthly totals
    3. Calculates ADR, occupancy %, RevPAR per month
    4. Writes actuals into Actuals_Input tab (Jan–Apr)
    5. Writes pace data into Pace_Weekly tab (May–Jul)
    6. Saves workbook
    7. Prints a summary to console (Slack/report step comes next)
"""

import openpyxl
from openpyxl import load_workbook
from collections import defaultdict
from datetime import datetime, date
import os
import sys
from dotenv import load_dotenv

load_dotenv("/Users/stephcleung/Library/CloudStorage/GoogleDrive-stephanie@lintonhospitality.com/Shared drives/Ferncrest/1_Locations/02_OK-FC/2_Finance/AssetManagement/.env", override=True)

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG — update these paths to match your folder structure
# ─────────────────────────────────────────────────────────────────────────────
CONFIG = {
    "ytd_file":       "/Users/stephcleung/Library/CloudStorage/GoogleDrive-stephanie@lintonhospitality.com/Shared drives/Ferncrest/1_Locations/02_OK-FC/2_Finance/AssetManagement/ok-fc_occupancy_2026ytd.xlsx",
    "pace_file":      "/Users/stephcleung/Library/CloudStorage/GoogleDrive-stephanie@lintonhospitality.com/Shared drives/Ferncrest/1_Locations/02_OK-FC/2_Finance/AssetManagement/ok-fc_occupancy_pace_150day.xlsx",
    "workbook":       "/Users/stephcleung/Library/CloudStorage/GoogleDrive-stephanie@lintonhospitality.com/Shared drives/Ferncrest/1_Locations/02_OK-FC/2_Finance/AssetManagement/ferncrest_ok-fc_v2.xlsx",
    "property_code":  "OK-FC",
    "total_units":    12,
    "fiscal_year":    2026,
}

# Days in each month for 2026 (non-leap year)
DAYS_IN_MONTH = {
    1: 31, 2: 28, 3: 31, 4: 30, 5: 31, 6: 30,
    7: 31, 8: 31, 9: 30, 10: 31, 11: 30, 12: 31
}

MONTH_NAMES = {
    1:"Jan", 2:"Feb", 3:"Mar", 4:"Apr", 5:"May", 6:"Jun",
    7:"Jul", 8:"Aug", 9:"Sep", 10:"Oct", 11:"Nov", 12:"Dec"
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1: Parse a Cloudbeds occupancy export
# ─────────────────────────────────────────────────────────────────────────────
def parse_cloudbeds_export(filepath):
    """
    Reads a Cloudbeds occupancy xlsx export.
    Returns a dict keyed by month number (int):
        {
            1: {
                "nights_sold": 56,
                "revenue": 13503.44,
                "available_nights": 372,   # units × days in month
                "occupancy_pct": 0.151,
                "adr": 241.13,
                "revpar": 36.30,
                "daily_rows": [...]         # raw daily data for pace tab
            },
            ...
        }
    """
    print(f"\n📂 Reading: {filepath}")

    if not os.path.exists(filepath):
        print(f"  ❌ File not found: {filepath}")
        sys.exit(1)

    wb = load_workbook(filepath, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    # Validate headers
    headers = rows[0]
    expected = ("Date", "Rooms Occupied Year 2026",
                "ACCOMMODATIONS BOOKED 2026", "Revenue 2026")
    if headers != expected:
        print(f"  ⚠️  Unexpected headers: {headers}")
        print(f"       Expected: {expected}")
        print(f"       Column mapping may be off — check before trusting output.")

    # Aggregate by month
    monthly = defaultdict(lambda: {
        "nights_sold": 0,
        "revenue": 0.0,
        "daily_rows": []
    })

    skipped = 0
    for row in rows[1:]:
        date_str, occ_pct, booked, revenue = row

        # Skip blank or zero rows (no-bookings days still matter for
        # available nights calc, but zero revenue/bookings rows are fine)
        if date_str is None:
            skipped += 1
            continue

        # Parse MM/DD date format from Cloudbeds
        try:
            month = int(str(date_str).split("/")[0])
        except (ValueError, IndexError):
            skipped += 1
            continue

        nights = int(booked) if booked else 0
        rev    = float(revenue) if revenue else 0.0

        monthly[month]["nights_sold"] += nights
        monthly[month]["revenue"]     += rev
        monthly[month]["daily_rows"].append({
            "date":        date_str,
            "occ_pct":     float(occ_pct) if occ_pct else 0.0,
            "nights_sold": nights,
            "revenue":     rev,
        })

    if skipped:
        print(f"  ℹ️  Skipped {skipped} blank rows")

    # Calculate derived metrics per month
    result = {}
    units = CONFIG["total_units"]
    for month, data in sorted(monthly.items()):
        days        = DAYS_IN_MONTH.get(month, 30)
        avail       = units * days
        nights_sold = data["nights_sold"]
        revenue     = data["revenue"]
        occ         = nights_sold / avail if avail > 0 else 0
        adr         = revenue / nights_sold if nights_sold > 0 else 0
        revpar      = revenue / avail if avail > 0 else 0

        result[month] = {
            "nights_sold":      nights_sold,
            "available_nights": avail,
            "revenue":          round(revenue, 2),
            "occupancy_pct":    round(occ, 4),
            "adr":              round(adr, 2),
            "revpar":           round(revpar, 2),
            "daily_rows":       data["daily_rows"],
        }

        print(f"  {MONTH_NAMES[month]:>3}: "
              f"nights={nights_sold:>3} / {avail}  "
              f"occ={occ:>5.1%}  "
              f"ADR=${adr:>7.2f}  "
              f"rev=${revenue:>9,.2f}")

    return result


# ─────────────────────────────────────────────────────────────────────────────
# STEP 2: Write actuals into Actuals_Input tab
# ─────────────────────────────────────────────────────────────────────────────
def write_actuals(ws_act, monthly_data):
    """
    Finds the correct rows in Actuals_Input and writes monthly values.
    Columns D–O = Jan–Dec (columns 4–15).
    """
    print("\n📝 Writing actuals into Actuals_Input tab...")

    # Row map: label → row number in Actuals_Input
    # We scan the sheet to find them dynamically — safer than hardcoding
    row_map = {}
    for row in ws_act.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                label = cell.value.strip()
                if label in [
                    "Room Revenue (excl. tax)",
                    "Nights Available",
                    "Nights Occupied",
                    "ADR (Actual)",
                ]:
                    row_map[label] = cell.row

    if not row_map:
        print("  ⚠️  Could not find target rows in Actuals_Input — check tab structure")
        return

    print(f"  Found rows: {row_map}")

    for month, data in monthly_data.items():
        col = 3 + month  # Jan=col4, Feb=col5, etc.

        if "Room Revenue (excl. tax)" in row_map:
            ws_act.cell(row_map["Room Revenue (excl. tax)"], col,
                        data["revenue"])

        if "Nights Available" in row_map:
            ws_act.cell(row_map["Nights Available"], col,
                        data["available_nights"])

        if "Nights Occupied" in row_map:
            ws_act.cell(row_map["Nights Occupied"], col,
                        data["nights_sold"])

        if "ADR (Actual)" in row_map:
            ws_act.cell(row_map["ADR (Actual)"], col,
                        data["adr"])

        print(f"  ✅ {MONTH_NAMES[month]}: "
              f"rev=${data['revenue']:,.2f}  "
              f"nights={data['nights_sold']}  "
              f"ADR=${data['adr']:.2f}")


# ─────────────────────────────────────────────────────────────────────────────
# STEP 3: Write pace data into Pace_Weekly tab
# ─────────────────────────────────────────────────────────────────────────────
def write_pace(ws_pace, monthly_data, as_of_date=None):
    """
    Appends a new weekly snapshot block to the Pace_Weekly rolling log.
    Finds the last data row and appends below it.
    """
    print("\n📊 Writing pace data into Pace_Weekly tab...")

    if as_of_date is None:
        as_of_date = date.today().strftime("%Y-%m-%d")

    # Budget occupancy assumptions (from Config/Budget tab)
    budget_occ = {
        5: 0.30, 6: 0.30, 7: 0.40, 8: 0.55,
        9: 0.20, 10: 0.50, 11: 0.20, 12: 0.10
    }

    # Find last used row in Pace_Weekly
    last_row = 1
    for row in ws_pace.iter_rows():
        for cell in row:
            if cell.value is not None:
                last_row = max(last_row, cell.row)

    # Add a section header for this week's snapshot
    insert_row = last_row + 2
    ws_pace.merge_cells(
        start_row=insert_row, start_column=1,
        end_row=insert_row, end_column=14
    )
    header_cell = ws_pace.cell(insert_row, 1,
        f"▶  WEEK OF {as_of_date}  (script auto-populated)")
    header_cell.font = openpyxl.styles.Font(
        name="Arial", bold=True, color="FFFCF9", size=10)
    header_cell.fill = openpyxl.styles.PatternFill(
        "solid", start_color="FF1F3320", end_color="FF1F3320")
    header_cell.alignment = openpyxl.styles.Alignment(
        horizontal="left", vertical="center", indent=1)

    data_row = insert_row + 1
    units = CONFIG["total_units"]

    for month, data in sorted(monthly_data.items()):
        days        = DAYS_IN_MONTH.get(month, 30)
        avail       = units * days
        nights_sold = data["nights_sold"]
        bud_occ     = budget_occ.get(month, 0.25)
        month_label = f"{MONTH_NAMES[month]} {CONFIG['fiscal_year']}"

        # Weeks until arrival (approx)
        today = date.today()
        # First of the arrival month
        arr_date = date(CONFIG["fiscal_year"], month, 1)
        weeks_out = max(0, (arr_date - today).days // 7)

        ws_pace.cell(data_row, 1,  as_of_date)
        ws_pace.cell(data_row, 2,  month_label)
        ws_pace.cell(data_row, 3,  weeks_out)
        ws_pace.cell(data_row, 4,  avail)
        ws_pace.cell(data_row, 5,  nights_sold)
        # Col 6: occ % paced — formula
        ws_pace.cell(data_row, 6,
            f"=IFERROR(E{data_row}/D{data_row},0)")
        ws_pace.cell(data_row, 6).number_format = "0.0%"
        # Col 7: LY same point — leave blank for now (no LY data yet)
        ws_pace.cell(data_row, 7,  0)
        # Col 8: LY occ %
        ws_pace.cell(data_row, 8,
            f"=IFERROR(G{data_row}/D{data_row},0)")
        ws_pace.cell(data_row, 8).number_format = "0.0%"
        # Col 9: Budget occ %
        ws_pace.cell(data_row, 9,  bud_occ)
        ws_pace.cell(data_row, 9).number_format = "0.0%"
        # Col 10: Pace vs LY
        ws_pace.cell(data_row, 10,
            f"=IFERROR(F{data_row}-H{data_row},0)")
        ws_pace.cell(data_row, 10).number_format = "0.0%"
        # Col 11: Pace vs Budget
        ws_pace.cell(data_row, 11,
            f"=IFERROR(F{data_row}-I{data_row},0)")
        ws_pace.cell(data_row, 11).number_format = "0.0%"
        # Col 12: ADR on books
        ws_pace.cell(data_row, 12, data["adr"])
        ws_pace.cell(data_row, 12).number_format = '$#,##0.00'
        # Col 13: Pickup since last week
        ws_pace.cell(data_row, 13,
            f'=IFERROR(E{data_row}-VLOOKUP(B{data_row},'
            f'B$5:E{data_row-1},4,0),"— first entry")')
        # Col 14: Flag
        ws_pace.cell(data_row, 14,
            f'=IF(E{data_row}=0,"— no bookings yet",'
            f'IF(K{data_row}<-0.05,"⚠️ Behind budget — review rate",'
            f'IF(K{data_row}>0.05,"✅ Ahead — hold rate or push ADR",'
            f'IF(J{data_row}<-0.03,"↓ Soft vs LY — monitor",'
            f'"◼ On track"))))')

        occ_pct = nights_sold / avail if avail > 0 else 0
        vs_bud  = occ_pct - bud_occ
        flag    = ("⚠️ Behind" if vs_bud < -0.05
                   else "✅ Ahead" if vs_bud > 0.05
                   else "◼ On track")

        print(f"  {month_label}: {nights_sold} nights / {avail} avail  "
              f"({occ_pct:.1%} paced vs {bud_occ:.1%} budget)  {flag}")

        data_row += 1


# ─────────────────────────────────────────────────────────────────────────────
# STEP 4: Update Report_Data bridge tab
# ─────────────────────────────────────────────────────────────────────────────
def update_report_data(ws_rd, ytd_data, pace_data):
    """
    Updates key-value pairs in Report_Data tab so the HTML report
    script can read a single structured output without touching
    other tabs.
    """
    print("\n🔗 Updating Report_Data bridge tab...")

    # Build update dict: key → new value
    # Aggregate YTD totals
    ytd_revenue     = sum(d["revenue"] for d in ytd_data.values())
    ytd_nights      = sum(d["nights_sold"] for d in ytd_data.values())
    ytd_avail       = sum(d["available_nights"] for d in ytd_data.values())
    ytd_occ         = ytd_nights / ytd_avail if ytd_avail else 0
    ytd_adr         = ytd_revenue / ytd_nights if ytd_nights else 0
    ytd_revpar      = ytd_revenue / ytd_avail if ytd_avail else 0

    # Pace: May, Jun, Jul on books
    may = pace_data.get(5, {})
    jun = pace_data.get(6, {})
    jul = pace_data.get(7, {})

    updates = {
        "actual_net_revenue_ytd":   round(ytd_revenue, 2),
        "actual_nights_sold_ytd":   ytd_nights,
        "actual_occ_rate_ytd":      round(ytd_occ, 4),
        "actual_adr_ytd":           round(ytd_adr, 2),
        "actual_revpar_ytd":        round(ytd_revpar, 2),
        "pace_may_on_books":        may.get("nights_sold", 0),
        "pace_may_occ_pct":         may.get("occupancy_pct", 0),
        "pace_may_adr":             may.get("adr", 0),
        "pace_jun_on_books":        jun.get("nights_sold", 0),
        "pace_jun_occ_pct":         jun.get("occupancy_pct", 0),
        "pace_jun_adr":             jun.get("adr", 0),
        "pace_jul_on_books":        jul.get("nights_sold", 0),
        "pace_jul_occ_pct":         jul.get("occupancy_pct", 0),
        "pace_jul_adr":             jul.get("adr", 0),
        "last_updated":             date.today().strftime("%Y-%m-%d"),
    }

    # Scan Report_Data for matching keys and update values
    updated = []
    for row in ws_rd.iter_rows():
        key_cell = row[0]  # Column A
        val_cell = row[1]  # Column B
        if key_cell.value in updates:
            val_cell.value = updates[key_cell.value]
            updated.append(key_cell.value)

    # Add any missing keys at the bottom
    missing = [k for k in updates if k not in updated]
    if missing:
        last = ws_rd.max_row + 1
        for key in missing:
            ws_rd.cell(last, 1, key)
            ws_rd.cell(last, 2, updates[key])
            last += 1
        print(f"  ℹ️  Added {len(missing)} new keys: {missing}")

    print(f"  ✅ Updated {len(updated)} Report_Data fields")
    print(f"\n  📈 YTD Summary:")
    print(f"     Revenue:    ${ytd_revenue:>10,.2f}")
    print(f"     Occ Rate:   {ytd_occ:>9.1%}")
    print(f"     ADR:        ${ytd_adr:>10.2f}")
    print(f"     RevPAR:     ${ytd_revpar:>10.2f}")
    print(f"\n  📅 90-Day Pace (nights on books):")
    print(f"     May: {may.get('nights_sold',0):>3} nights  "
          f"({may.get('occupancy_pct',0):.1%} paced)  "
          f"ADR ${may.get('adr',0):.2f}")
    print(f"     Jun: {jun.get('nights_sold',0):>3} nights  "
          f"({jun.get('occupancy_pct',0):.1%} paced)  "
          f"ADR ${jun.get('adr',0):.2f}")
    print(f"     Jul: {jul.get('nights_sold',0):>3} nights  "
          f"({jul.get('occupancy_pct',0):.1%} paced)  "
          f"ADR ${jul.get('adr',0):.2f}")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────# STEP 5: Call Claude API for AI recommendations
# ─────────────────────────────────────────────────────────────────────────────
def get_ai_recommendations(ytd_data, pace_data, as_of_date):
    """
    Calls Claude API with current pace and actuals context.
    Returns 4 recommendation HTML blocks ready to inject into the dashboard.
    Falls back to placeholder text if API call fails.
    """
    import urllib.request
    import json
    import os

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        print("  ⚠️  ANTHROPIC_API_KEY not set — using placeholder recommendations")
        return None

    # Build context from real data
    ytd_revenue     = sum(d["revenue"] for d in ytd_data.values())
    ytd_nights      = sum(d["nights_sold"] for d in ytd_data.values())
    ytd_avail       = sum(d["available_nights"] for d in ytd_data.values())
    ytd_occ         = ytd_nights / ytd_avail if ytd_avail else 0
    ytd_adr         = ytd_revenue / ytd_nights if ytd_nights else 0

    budget_occ = {
        5:0.30, 6:0.30, 7:0.40, 8:0.55, 9:0.20, 10:0.50, 11:0.20, 12:0.10
    }

    pace_lines = []
    for month, data in sorted(pace_data.items()):
        bud = budget_occ.get(month, 0.25)
        occ = data["occupancy_pct"]
        vs  = occ - bud
        pace_lines.append(
            f"  {MONTH_NAMES[month]}: {data['nights_sold']} nights on books "
            f"({occ:.1%} paced vs {bud:.0%} budget, {vs:+.1%} variance)"
            + (f", ADR on books ${data['adr']:.0f}" if data['adr'] > 0 else "")
        )

    prompt = f"""You are an expert hospitality asset manager providing weekly pricing and demand recommendations for a glamping property.

PROPERTY: Ferncrest Flint Creek, Colcord, Oklahoma (NE Oklahoma, serves Tulsa/OKC/NWA/DFW drive market)
UNITS: 12 glamping dome tents (mix of creek-side premium and commons units)
AS-OF DATE: {as_of_date}
BUDGET ADR: $233

YTD ACTUALS (Jan–Apr 2026):
  Revenue: ${ytd_revenue:,.0f}
  Occupancy: {ytd_occ:.1%}
  ADR: ${ytd_adr:.0f}
  Nights sold: {ytd_nights} of {ytd_avail} available

FORWARD PACE (nights currently on books):
{chr(10).join(pace_lines)}

KEY CONTEXT:
  - 40% of bookings historically arrive within 7 days of check-in (last-minute dominant market)
  - Early bookers are least price-sensitive — ADR on books skews high vs final blended rate
  - Creek-side units average $241 ADR vs Commons $193 — rate ladder matters
  - July has only 10 nights but at $395 ADR — strong early-booker signal
  - Oklahoma schools release ~May 22, expect summer bookings to accelerate

Generate exactly 4 pricing and demand recommendations for this week. Each should be specific, actionable, and tied to the actual data above. Consider upcoming holidays, events, and demand patterns for NE Oklahoma.

Respond ONLY with a JSON array of exactly 4 objects, no markdown, no preamble:
[
  {{
    "priority": "urgent|opportunity|monitor",
    "priority_label": "short label (e.g. 'Urgent — 22 days out')",
    "title": "recommendation title (max 8 words)",
    "body": "2-3 sentence explanation tied to specific data points",
    "action": "one concrete action to take this week"
  }}
]"""

    try:
        payload = json.dumps({
            "model": "claude-sonnet-4-20250514",
            "max_tokens": 1000,
            "messages": [{"role": "user", "content": prompt}]
        }).encode()

        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01"
            }
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            result = json.loads(resp.read())
            recs_json = result["content"][0]["text"].strip()
            # Strip markdown fences if present
            recs_json = recs_json.replace("```json","").replace("```","").strip()
            recs = json.loads(recs_json)
            print(f"  ✅ Got {len(recs)} AI recommendations from Claude API")
            return recs

    except Exception as e:
        print(f"  ⚠️  Claude API call failed: {e} — using placeholder recommendations")
        return None


def build_ai_rec_html(recs):
    """Converts recommendation dicts to HTML blocks for injection."""
    if not recs:
        return None  # Use hardcoded placeholder in template

    html_blocks = []
    priority_labels = {
        "urgent":      ("urgent",      "⚡"),
        "opportunity": ("opportunity", "↑"),
        "monitor":     ("monitor",     "◼"),
    }

    for rec in recs:
        p = rec.get("priority", "monitor")
        css_class, icon = priority_labels.get(p, ("monitor", "◼"))
        label = rec.get("priority_label", p.title())
        title = rec.get("title", "")
        body  = rec.get("body", "")
        action= rec.get("action", "")

        html_blocks.append(f'''      <div class="ai-rec">
        <div class="ai-rec-priority {css_class}">{icon} {label}</div>
        <div class="ai-rec-title">{title}</div>
        <div class="ai-rec-body">{body}</div>
        <div class="ai-rec-action">→ Action: {action}</div>
      </div>''')

    return "\n".join(html_blocks)


# ─────────────────────────────────────────────────────────────────────────────
# STEP 6: Inject data into HTML dashboard
# ─────────────────────────────────────────────────────────────────────────────
def generate_html_report(ytd_data, pace_data, as_of_date,
                         template_path="ferncrest_internal_dashboard.html",
                         output_path="ferncrest_internal_dashboard.html"):
    """
    Reads the HTML template, replaces all INJECT placeholders with
    real data, writes the output file.
    """
    import re

    print(f"\n📄 Generating HTML report...")

    if not os.path.exists(template_path):
        print(f"  ⚠️  Template not found: {template_path}")
        return

    with open(template_path, "r") as f:
        html = f.read()

    # ── Compute values ──
    ytd_revenue = sum(d["revenue"] for d in ytd_data.values())
    ytd_nights  = sum(d["nights_sold"] for d in ytd_data.values())
    ytd_avail   = sum(d["available_nights"] for d in ytd_data.values())
    ytd_occ     = ytd_nights / ytd_avail if ytd_avail else 0
    ytd_adr     = ytd_revenue / ytd_nights if ytd_nights else 0
    ytd_revpar  = ytd_revenue / ytd_avail if ytd_avail else 0

    # Budget assumptions (Jan–Apr)
    budget_occ_ytd = {1:0.20, 2:0.20, 3:0.20, 4:0.20}
    budget_adr     = 233
    units          = CONFIG["total_units"]
    days_map       = DAYS_IN_MONTH

    budget_rev_ytd = sum(
        budget_adr * units * budget_occ_ytd.get(m, 0.20) * days_map[m]
        for m in ytd_data.keys()
    )
    budget_occ_avg_ytd = sum(budget_occ_ytd.get(m, 0.20)
                             for m in ytd_data.keys()) / len(ytd_data)
    budget_revpar_ytd  = budget_adr * budget_occ_avg_ytd

    rev_vs_budget    = ytd_revenue - budget_rev_ytd
    occ_vs_budget    = ytd_occ - budget_occ_avg_ytd
    adr_vs_budget    = ytd_adr - budget_adr
    revpar_vs_budget = ytd_revpar - budget_revpar_ytd

    def fmt_dollar(v, show_sign=False):
        sign = "+" if v >= 0 else "−"
        amt  = abs(v)
        if amt >= 1000:
            return f"{'+'if show_sign and v>=0 else sign if v<0 else ''}${amt/1000:.1f}K"
        return f"{'+'if show_sign and v>=0 else sign if v<0 else ''}${amt:.0f}"

    def fmt_pct(v, show_sign=False):
        sign = "+" if v >= 0 else "−"
        return f"{'+' if show_sign and v>=0 else sign if v<0 else ''}{abs(v)*100:.1f} pts"

    # ── Last closed month ──
    last_month = max(ytd_data.keys())
    data_through = f"{MONTH_NAMES[last_month]} 30, {CONFIG['fiscal_year']}" \
                   if last_month in [4,6,9,11] \
                   else f"{MONTH_NAMES[last_month]} {days_map[last_month]}, {CONFIG['fiscal_year']}"

    # ── Pace JS data block ──
    budget_occ_fwd = {
        5:0.30, 6:0.30, 7:0.40, 8:0.55, 9:0.20, 10:0.50, 11:0.20, 12:0.10
    }
    today = date.today()
    pace_js_rows = []
    for month in sorted(pace_data.keys()):
        data    = pace_data[month]
        avail   = data["available_nights"]
        on_books= data["nights_sold"]
        bud_occ = budget_occ_fwd.get(month, 0.25)
        adr     = data["adr"]
        arr_date= date(CONFIG["fiscal_year"], month, 1)
        weeks_out = max(0, (arr_date - today).days // 7)
        fy = CONFIG["fiscal_year"]
        mo_name = MONTH_NAMES[month]
        pace_js_rows.append(
            f"  {{month:'{mo_name} {fy}', "
            f"avail:{avail}, onBooks:{on_books}, "
            f"budgetOcc:{bud_occ}, adr:{adr:.2f}, weeksOut:{weeks_out}}},"
        )

    pace_js = "const paceData=[\n" + "\n".join(pace_js_rows) + "\n];"

    # ── AI recommendations ──
    print("  🤖 Fetching AI recommendations...")
    recs     = get_ai_recommendations(ytd_data, pace_data, as_of_date)
    recs_html= build_ai_rec_html(recs)

    # ── Simple placeholder replacer ──
    def inject(html, key, value):
        pattern = f"<!-- INJECT: {key} -->.*?<!-- /INJECT -->"
        replacement = f"<!-- INJECT: {key} -->{value}<!-- /INJECT -->"
        return re.sub(pattern, replacement, html, flags=re.DOTALL)

    # Inject scalar values
    html = inject(html, "generated_date", today.strftime("%B %-d, %Y"))
    html = inject(html, "data_through",   data_through)
    html = inject(html, "week_of",        today.strftime("%B %-d, %Y"))
    html = inject(html, "pace_as_of",     today.strftime("%B %-d, %Y"))

    # KPI values
    html = inject(html, "ytd_revenue",         f"${ytd_revenue/1000:.1f}K")
    html = inject(html, "ytd_rev_vs_budget",   fmt_dollar(rev_vs_budget))
    html = inject(html, "ytd_rev_vs_forecast", fmt_dollar(rev_vs_budget))
    html = inject(html, "ytd_occ",             f"{ytd_occ:.1%}")
    html = inject(html, "ytd_occ_vs_budget",   fmt_pct(occ_vs_budget))
    html = inject(html, "ytd_adr",             f"${ytd_adr:.0f}")
    html = inject(html, "ytd_adr_vs_budget",   fmt_dollar(adr_vs_budget, show_sign=True))
    html = inject(html, "ytd_revpar",          f"${ytd_revpar:.0f}")
    html = inject(html, "ytd_revpar_vs_budget",fmt_dollar(revpar_vs_budget))

    # Pace JS data block
    html = re.sub(
        r"// INJECT: pace_data_js\nconst paceData=\[.*?\];",
        f"// INJECT: pace_data_js\n{pace_js}",
        html, flags=re.DOTALL
    )

    # AI recommendations (only if API returned something)
    if recs_html:
        html = re.sub(
            r"<!-- INJECT: ai_recommendations -->.*?<!-- /INJECT -->",
            f"<!-- INJECT: ai_recommendations -->\n{recs_html}\n      <!-- /INJECT -->",
            html, flags=re.DOTALL
        )

    # Fix KPI variance CSS classes based on actual direction
    # Revenue behind → rust, ADR ahead → sage
    for key, val in [
        ("ytd_rev_vs_budget",    rev_vs_budget >= 0),
        ("ytd_rev_vs_forecast",  rev_vs_budget >= 0),
        ("ytd_occ_vs_budget",    occ_vs_budget >= 0),
        ("ytd_adr_vs_budget",    adr_vs_budget >= 0),
        ("ytd_revpar_vs_budget", revpar_vs_budget >= 0),
    ]:
        css_class = "ahead" if val else "behind"
        html = html.replace(
            f'<!-- INJECT: {key} -->',
            f'<!-- INJECT: {key} -->'
        )

    with open(output_path, "w") as f:
        f.write(html)

    print(f"  ✅ HTML report saved: {output_path}")
    print(f"     Revenue: ${ytd_revenue:,.0f} | Occ: {ytd_occ:.1%} | ADR: ${ytd_adr:.0f} | RevPAR: ${ytd_revpar:.0f}")


# ─────────────────────────────────────────────────────────────────────────────
# STEP 7: Post Slack digest
# ─────────────────────────────────────────────────────────────────────────────
def post_slack_digest(ytd_data, pace_data, as_of_date,
                      prev_pace_data=None):
    """
    Posts a weekly digest to Slack via webhook.
    Focuses on trend direction, pick-up velocity, ADR movement,
    and specific actionable flags — not raw pace vs budget.

    prev_pace_data: pace snapshot from last week (same format as pace_data).
                    Used to calculate week-over-week pick-up.
                    If None, pick-up delta is not shown.
    """
    import urllib.request
    import json
    import os
    from datetime import date, timedelta

    webhook_url = os.environ.get("SLACK_WEBHOOK_URL", "")
    if not webhook_url:
        print("\n📣 Slack webhook not configured — skipping post")
        print("   Set SLACK_WEBHOOK_URL environment variable to enable")
        return

    print("\n📣 Posting Slack digest...")

    today = date.today()

    # ── YTD numbers ──
    ytd_revenue = sum(d["revenue"] for d in ytd_data.values())
    ytd_nights  = sum(d["nights_sold"] for d in ytd_data.values())
    ytd_avail   = sum(d["available_nights"] for d in ytd_data.values())
    ytd_occ     = ytd_nights / ytd_avail if ytd_avail else 0
    ytd_adr     = ytd_revenue / ytd_nights if ytd_nights else 0

    # ── Pick-up this week (total nights booked across all future months) ──
    total_on_books_now  = sum(d["nights_sold"] for d in pace_data.values())
    total_on_books_prev = sum(d["nights_sold"] for d in prev_pace_data.values())                           if prev_pace_data else None
    weekly_pickup = total_on_books_now - total_on_books_prev                     if total_on_books_prev is not None else None

    # ── Biggest mover this week ──
    biggest_mover = None
    biggest_pickup = 0
    if prev_pace_data:
        for month, data in pace_data.items():
            prev = prev_pace_data.get(month, {}).get("nights_sold", 0)
            pickup = data["nights_sold"] - prev
            if pickup > biggest_pickup:
                biggest_pickup = pickup
                biggest_mover  = (MONTH_NAMES[month], pickup)

    # ── ADR on new bookings (weighted avg across future months) ──
    adr_months = [(d["adr"], d["nights_sold"])
                  for d in pace_data.values()
                  if d["adr"] > 0 and d["nights_sold"] > 0]
    blended_adr = (sum(a * n for a, n in adr_months) / sum(n for _, n in adr_months)
                   if adr_months else 0)

    # ── Weekends needing attention (next 21 days, < 30% booked) ──
    # Approximate from monthly pace — flag months within 3 weeks
    urgent_months = []
    for month, data in sorted(pace_data.items()):
        arr_date  = date(CONFIG["fiscal_year"], month, 1)
        weeks_out = max(0, (arr_date - today).days // 7)
        occ       = data["occupancy_pct"]
        if weeks_out <= 3 and occ < 0.30 and data["nights_sold"] > 0:
            urgent_months.append((MONTH_NAMES[month], occ, weeks_out))

    # ── One AI action item (pull from recommendations if available) ──
    # We pass the top recommendation as a simple string
    # This gets populated by get_ai_recommendations separately
    # For now derive a simple heuristic action
    fwd_months = sorted(pace_data.keys())
    action_month = None
    for m in fwd_months[:3]:
        d = pace_data[m]
        arr  = date(CONFIG["fiscal_year"], m, 1)
        wks  = max(0, (arr - today).days // 7)
        if wks <= 6 and d["nights_sold"] > 0:
            action_month = (MONTH_NAMES[m], d["nights_sold"],
                            d["occupancy_pct"], d["adr"], wks)
            break

    if action_month:
        mo, nights, occ, adr, wks = action_month
        if occ < 0.20:
            action_text = (f"*This week:* {mo} is {wks} weeks out with "
                           f"{nights} nights booked ({occ:.1%}). "
                           f"Push group outreach and packages — volume gap, not a rate issue.")
        elif adr > 280:
            action_text = (f"*This week:* {mo} early bookers paying ${adr:.0f} ADR. "
                           f"Hold rate — do not discount to chase volume.")
        else:
            action_text = (f"*This week:* {mo} pacing at {occ:.1%} with "
                           f"${adr:.0f} ADR on books. Monitor — no action needed yet.")
    else:
        action_text = "*This week:* No near-term months require immediate action."

    # ── Build Slack blocks ──
    prop_code = CONFIG["property_code"]
    prop_name = "Flint Creek" if prop_code == "OK-FC" else prop_code

    # Pick-up line
    if weekly_pickup is not None:
        pickup_str = (f"*+{weekly_pickup} nights* picked up this week"
                      if weekly_pickup >= 0
                      else f"*{weekly_pickup} nights* net change this week")
        if biggest_mover:
            pickup_str += f" · biggest mover: *{biggest_mover[0]}* (+{biggest_mover[1]})"
    else:
        pickup_str = f"*{total_on_books_now} nights* total on books across all future months"

    # ADR line
    adr_str = (f"*ADR on books:* ${blended_adr:.0f} (blended across forward months)"
               if blended_adr > 0 else "*ADR on books:* No forward bookings yet")

    # Urgent flag
    if urgent_months:
        mo, occ, wks = urgent_months[0]
        urgent_str = f"*Watch:* {mo} is {wks} weeks out at {occ:.1%} booked — needs attention"
    else:
        urgent_str = "*No urgent near-term flags*"

    blocks = [
        {
            "type": "header",
            "text": {
                "type": "plain_text",
                "text": f"{prop_name} — Monday Digest  {as_of_date}"
            }
        },
        {
            "type": "section",
            "fields": [
                {"type":"mrkdwn","text":f"*YTD Revenue*\n${ytd_revenue:,.0f}"},
                {"type":"mrkdwn","text":f"*YTD Occ*\n{ytd_occ:.1%}"},
                {"type":"mrkdwn","text":f"*YTD ADR*\n${ytd_adr:.0f}"},
                {"type":"mrkdwn","text":f"*Nights Sold YTD*\n{ytd_nights}"},
            ]
        },
        {"type": "divider"},
        {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": f":chart_with_upwards_trend: *This Week*\n{pickup_str}\n{adr_str}\n{urgent_str}"
            }
        },
        {"type": "divider"},
        {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": f":dart: {action_text}"
            }
        },
        {
            "type": "context",
            "elements": [
                {
                    "type": "mrkdwn",
                    "text": (f"Ferncrest Pipeline · {as_of_date} · "
                             f"Open dashboard for full pace detail + AI recommendations")
                }
            ]
        }
    ]

    try:
        payload = json.dumps({"blocks": blocks}).encode()
        req = urllib.request.Request(
            webhook_url,
            data=payload,
            headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            print(f"  ✅ Slack digest posted (status {resp.status})")
    except Exception as e:
        print(f"  ⚠️  Slack post failed: {e}")

def main():
    print("=" * 60)
    print("  FERNCREST DATA PIPELINE")
    print(f"  Property: {CONFIG['property_code']}")
    print(f"  Run date: {date.today().strftime('%Y-%m-%d')}")
    print("=" * 60)

    # Parse both Cloudbeds exports
    ytd_data  = parse_cloudbeds_export(CONFIG["ytd_file"])
    pace_data = parse_cloudbeds_export(CONFIG["pace_file"])

    # Load the property workbook
    wb_path = CONFIG["workbook"]
    if not os.path.exists(wb_path):
        print(f"\n❌ Workbook not found: {wb_path}")
        print("   Place ferncrest_ok-fc_v2.xlsx in the same folder as this script.")
        sys.exit(1)

    print(f"\n📖 Loading workbook: {wb_path}")
    wb = load_workbook(wb_path)
    print(f"   Sheets: {wb.sheetnames}")

    # Write to tabs
    write_actuals(wb["Actuals_Input"], ytd_data)
    write_pace(wb["Pace_Weekly"], pace_data)
    update_report_data(wb["Report_Data"], ytd_data, pace_data)

    # Save workbook
    wb.save(wb_path)
    print(f"\n💾 Workbook saved: {wb_path}")

    # Generate HTML report
    generate_html_report(
        ytd_data, pace_data, as_of_date=date.today().strftime("%Y-%m-%d"),
        template_path="ferncrest_internal_dashboard.html",
        output_path="ferncrest_internal_dashboard.html"
    )

    # Post Slack digest
    post_slack_digest(ytd_data, pace_data,
                      as_of_date=date.today().strftime("%Y-%m-%d"))

    print("\n✅ Pipeline complete.")
    print("=" * 60)

    return {
        "ytd": ytd_data,
        "pace": pace_data,
    }


if __name__ == "__main__":
    main()


# ─────────────────────────────────────────────────────────────────────────────
