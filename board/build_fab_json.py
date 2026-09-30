#!/usr/bin/env python3
"""
MIW FABRICATION BOARD - data builder
Writes production_data_fab.json (+ index_fab.html if fab_template.html is present).

Standalone on purpose: it does NOT touch build_production_json.py, so the live
production board can never break because of this report. The readers and the
man-hours filters are copied from that builder so the numbers tie out.

What it produces
  MONTH grid : rows P1 / P3 / P4 / TOTAL, columns = each Mon-Sat week of the month
  WEEK grid  : rows P1 / P3 / P4 / TOTAL, columns = Mon..Sat of the current week
  Every cell : fit tons, man hours, MH per ton

Sources (the same Daily Reports drop the live board reads)
  <SOURCE>/QC Final Report/*      -> fit tons
  <SOURCE>/QuickBooks Time/*.csv  -> man hours

Usage:  python build_fab_json.py <SOURCE_FOLDER> <OUT_DIR>
Env:    TODAY=YYYY-MM-DD    override the date for testing
        HOURS_MODE=all|bays  all (default) = every dept in the plant, rows add up
                             bays          = fitting areas only, matches the live
                                             board's per-plant MH/ton tiles
"""
import json, datetime, sys, os, re, io, csv, glob
import openpyxl

SOURCE  = sys.argv[1] if len(sys.argv) > 1 else "."
OUT_DIR = sys.argv[2] if len(sys.argv) > 2 else "."
TODAY_OVERRIDE = os.environ.get("TODAY")
HOURS_MODE = os.environ.get("HOURS_MODE", "all").lower()
# THE CANONICAL FEED. Same file the production board and the WEEK ## SCHEDULES workbook
# read, so all three tie out. Falls back to the QC Final Report .xls folder if unset.
FEED_FIT = os.environ.get("FEED_FIT", "")

# ---- goals: per-week fit tons. Keep in step with build_production_json.py ----
PLANT_WEEK_GOAL = {"P1": 600.0, "P4": 100.0}
WORK_DAYS_PER_WEEK = 7              # Mon-Sun on the grid (Sunday added 2026-09-29)
GOAL_DAYS_PER_WEEK = 5              # Mon-Fri drives pace, Saturday is bonus
MHTON_GREEN, MHTON_AMBER = 20.0, 30.0

# Plant 3 removed 2026-08-19 at the plant manager's request: they do not want to
# see it on this board. Add ("P3","Plant 3") back here to restore it.
PLANTS = [("P1", "Plant 1"), ("P4", "Plant 4")]
FIT_LOC_TO_PLANT = {"East Bay": "P1", "West Bay": "P1", "Paint Bay": "P1",
                    "Plant 3": "P3", "Plant 4": "P4"}
NONWORK_KEYS = ("lunch", "vacation", "time off", "no call", "no show", "tardy",
                "leave early", "holiday", "bereavement", "jury", "sick", "pto")


def tile_plant(loc):        # per-area blocks: P1 = bays, P4 = Fab only
    loc = loc or ""
    if loc in ("P1 - East Bay", "P1 - West Bay"): return "P1"
    if loc.startswith("P3"): return "P3"
    if loc == "P4 - Fab": return "P4"
    return None


def overall_plant(loc):     # every department in the plant
    loc = loc or ""
    if loc.startswith("P1"): return "P1"
    if loc.startswith("P3"): return "P3"
    if loc.startswith("P4"): return "P4"
    return None


HOURS_PLANT = overall_plant if HOURS_MODE == "all" else tile_plant


# ---------- helpers (same as the live builder) ----------
def monday_of(d): return d - datetime.timedelta(days=d.weekday())


def numf(v):
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).replace(",", "").replace("#", "").replace("$", "").strip())
    except Exception: return 0.0


def file_date(fn):
    m = re.match(r"(\d{8})", os.path.basename(fn))
    if not m: return None
    s = m.group(1)
    try: return datetime.date(int(s[:4]), int(s[4:6]), int(s[6:8]))
    except Exception: return None


def loadx(p):
    try: return openpyxl.load_workbook(io.BytesIO(open(p, "rb").read()), data_only=True)
    except Exception: return None       # mid-write / cloud-only / corrupt: skip


def find_header(ws):
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "Job #": return r
    return None


def is_jobnum(a):
    return isinstance(a, (int, float)) or (isinstance(a, str) and str(a).strip().isdigit())


def _row_with(ws, label):
    for r in range(1, min(ws.max_row, 20) + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and label.lower() in v.lower(): return r
    return None


def _row_value(ws, r):
    for c in range(2, ws.max_column + 1):
        v = ws.cell(r, c).value
        if v not in (None, ""): return v
    return None


def _parse_date(v):
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date): return v
    if isinstance(v, str):
        for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"):
            try: return datetime.datetime.strptime(v.strip(), fmt).date()
            except Exception: pass
    return None


def qc_workdate(ws, fd):
    r = _row_with(ws, "Test Date")
    if r:
        v = _row_value(ws, r)
        d = _parse_date(v)
        if d: return d
        if isinstance(v, str) and "current date" in v.lower():
            m = re.search(r"-\s*(\d+)", v)
            return fd - datetime.timedelta(days=int(m.group(1)) if m else 0)
    return fd - datetime.timedelta(days=1)


def atomic_write(path, text):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f: f.write(text)
    os.replace(tmp, path)


# ---------- readers ----------
def read_fit(root):
    rows = []
    for fn in glob.glob(os.path.join(root, "QC Final Report", "*")):
        base = os.path.basename(fn)
        if base.startswith("~") or base.startswith("."): continue
        fd = file_date(fn)
        if not fd: continue
        wb = loadx(fn)
        if wb is None: continue
        ws = wb.active
        h = find_header(ws)
        if not h: continue
        wd = qc_workdate(ws, fd)
        for r in range(h + 1, ws.max_row + 1):
            a = ws.cell(r, 1).value
            if not is_jobnum(a): continue
            p = FIT_LOC_TO_PLANT.get(ws.cell(r, 4).value)
            if not p: continue
            rows.append({"dt": wd, "p": p, "t": numf(ws.cell(r, 8).value) / 2000.0,
                         "j": int(str(a).strip())})
    return rows


def read_fit_feed(path):
    """Fit rows from feed_fit.csv, the canonical feed.

    Plant is ALREADY RESOLVED there (C-Beams and SHOP fold into P1 and report as West Bay,
    Shop 3 is Plant 3), so do not re-derive it from Test Location: the old map had no
    C-Beams entry and silently dropped it. P3 is filtered out here because this board has
    excluded Plant 3 since 2026-08-19 at the plant manager's request."""
    rows = []
    if not path or not os.path.exists(path): return rows
    keep = {k for k, _ in PLANTS}
    try: rdr = csv.DictReader(open(path, encoding="utf-8-sig"))
    except Exception: return rows
    for r in rdr:
        p = (r.get("Plant") or "").strip()
        if p not in keep: continue
        try: dt = datetime.date.fromisoformat((r.get("Report Date") or "")[:10])
        except Exception: continue
        j = None
        m = re.search(r"(\d{3,5})", str(r.get("Job #") or ""))
        if m: j = int(m.group(1))
        rows.append({"dt": dt, "p": p, "t": numf(r.get("Weight")) / 2000.0, "j": j})
    return rows


def read_hours(root):
    rows = []
    for fn in glob.glob(os.path.join(root, "QuickBooks Time", "*.csv")):
        try:
            raw = open(fn, encoding="utf-8", errors="replace").read().replace("\x00", "")
        except Exception:
            continue
        try:
            for row in csv.DictReader(io.StringIO(raw)):
                try: d = datetime.date.fromisoformat(row["local_date"])
                except Exception: continue
                j = None
                m = re.search(r"(\d{3,5})", str(row.get("jobcode_2") or ""))
                if m: j = int(m.group(1))
                if j is None:
                    code = ((row.get("jobcode_1") or "") + " " + (row.get("jobcode_2") or "")).lower()
                    if any(k in code for k in NONWORK_KEYS): continue
                p = HOURS_PLANT(row.get("group"))
                if not p: continue
                rows.append({"dt": d, "p": p, "h": numf(row.get("hours"))})
        except Exception:
            continue        # one corrupt CSV must never crash the build
    return rows


# ---------- aggregation ----------
def mhton(h, t): return round(h / t, 1) if t > 0 else None


def build_grid(fit_rows, hr_rows, buckets, today):
    bidx = {}
    for i, (_, ds) in enumerate(buckets):
        for d in ds: bidx[d] = i
    nb = len(buckets)
    tons = {p: [0.0] * nb for p, _ in PLANTS}
    hrs = {p: [0.0] * nb for p, _ in PLANTS}
    for r in fit_rows:
        i = bidx.get(r["dt"])
        if i is not None and r["p"] in tons: tons[r["p"]][i] += r["t"]
    for r in hr_rows:
        i = bidx.get(r["dt"])
        if i is not None and r["p"] in hrs: hrs[r["p"]][i] += r["h"]

    rows_out = []
    for key, name in PLANTS:
        cells = [{"tons": round(tons[key][i], 1), "hours": round(hrs[key][i]),
                  "mhton": mhton(hrs[key][i], tons[key][i])} for i in range(nb)]
        tt, th = sum(tons[key]), sum(hrs[key])
        rows_out.append({"key": key, "name": name, "cells": cells,
                         "total": {"tons": round(tt, 1), "hours": round(th),
                                   "mhton": mhton(th, tt)}})
    tot_cells = []
    for i in range(nb):
        t = sum(tons[p][i] for p, _ in PLANTS)
        h = sum(hrs[p][i] for p, _ in PLANTS)
        tot_cells.append({"tons": round(t, 1), "hours": round(h), "mhton": mhton(h, t)})
    T = sum(sum(tons[p]) for p, _ in PLANTS)
    H = sum(sum(hrs[p]) for p, _ in PLANTS)
    return {"cols": [b[0] for b in buckets], "rows": rows_out,
            "total_row": {"key": "ALL", "name": "Total", "cells": tot_cells,
                          "total": {"tons": round(T, 1), "hours": round(H),
                                    "mhton": mhton(H, T)}}}


def workdays_between(a, b):
    n, d = 0, a
    while d <= b:
        if d.weekday() < 5: n += 1      # Mon-Fri; Saturday is bonus
        d += datetime.timedelta(days=1)
    return n


def main():
    today = datetime.date.fromisoformat(TODAY_OVERRIDE) if TODAY_OVERRIDE else datetime.date.today()
    if FEED_FIT and os.path.exists(FEED_FIT):
        fit_rows = read_fit_feed(FEED_FIT); print("fit source: feed_fit.csv")
    else:
        fit_rows = read_fit(SOURCE); print("fit source: QC Final Report folder (feed not found)")
    hr_rows = read_hours(SOURCE)   # hours are not in PowerFab, QuickBooks only

    # ---- month: one bucket per Mon-Sat week touching the month ----
    m_start = today.replace(day=1)
    m_end = (m_start.replace(day=28) + datetime.timedelta(days=4)).replace(day=1) - datetime.timedelta(days=1)
    wk = monday_of(m_start)
    mbuckets = []
    while wk <= m_end:
        days = {wk + datetime.timedelta(days=i) for i in range(WORK_DAYS_PER_WEEK)}
        days = {d for d in days if m_start <= d <= m_end}
        if days:
            md = min(days)
            mbuckets.append((f"Wk {md.month}/{md.day}", days))
        wk += datetime.timedelta(days=7)
    month = build_grid(fit_rows, hr_rows, mbuckets, today)

    # ---- week: Mon..Sat of the current week ----
    ws_start = monday_of(today)
    wdays = [ws_start + datetime.timedelta(days=i) for i in range(WORK_DAYS_PER_WEEK)]
    wbuckets = [(d.strftime("%a"), {d}) for d in wdays]
    week = build_grid(fit_rows, hr_rows, wbuckets, today)

    month_idx = next((i for i, (_, ds) in enumerate(mbuckets) if today in ds), len(mbuckets) - 1)
    wdelta = (today - ws_start).days
    week_idx = wdelta if 0 <= wdelta < WORK_DAYS_PER_WEEK else WORK_DAYS_PER_WEEK - 1

    seen = {r["dt"] for r in fit_rows} | {r["dt"] for r in hr_rows}
    last_day = max([d for d in seen if d <= today], default=today)

    wk_goal = sum(PLANT_WEEK_GOAL.values())
    m_days = workdays_between(m_start, m_end)
    m_goal = round(wk_goal * m_days / GOAL_DAYS_PER_WEEK)
    m_elapsed = max(1, workdays_between(m_start, min(last_day, m_end)))
    w_elapsed = max(1, workdays_between(ws_start, min(last_day, wdays[-1])))
    mt = month["total_row"]["total"]
    wt = week["total_row"]["total"]

    data = {
        "generated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "month_label": today.strftime("%B %Y"),
        "week_label": (f"Week {ws_start.isocalendar()[1]} · {ws_start.strftime('%b')} {ws_start.day}–{wdays[-1].day}"
                       if ws_start.month == wdays[-1].month else
                       f"Week {ws_start.isocalendar()[1]} · {ws_start.strftime('%b')} {ws_start.day} – {wdays[-1].strftime('%b')} {wdays[-1].day}"),
        "as_of": f"{last_day.month}/{last_day.day}",
        "hours_mode": HOURS_MODE,
        "month": month, "week": week,
        "month_today_index": month_idx, "week_today_index": week_idx,
        "goals": {
            "week_tons": wk_goal, "month_tons": m_goal,
            "week_pct": round(wt["tons"] / wk_goal * 100) if wk_goal else 0,
            "month_pct": round(mt["tons"] / m_goal * 100) if m_goal else 0,
            "month_proj": round(mt["tons"] / m_elapsed * m_days),
            "week_proj": round(wt["tons"] / w_elapsed * GOAL_DAYS_PER_WEEK),
            "month_expected": round(m_goal * m_elapsed / m_days, 1) if m_days else 0,
            "week_expected": round(wk_goal * w_elapsed / GOAL_DAYS_PER_WEEK, 1),
            "elapsed_days": m_elapsed, "total_days": m_days,
            "week_elapsed": w_elapsed,
        },
        "thresholds": {"green": MHTON_GREEN, "amber": MHTON_AMBER},
    }

    os.makedirs(OUT_DIR, exist_ok=True)
    atomic_write(os.path.join(OUT_DIR, "production_data_fab.json"), json.dumps(data, indent=2))

    tpl = os.path.join(OUT_DIR, "fab_template.html")
    if os.path.exists(tpl):
        atomic_write(os.path.join(OUT_DIR, "index_fab.html"),
                     open(tpl, encoding="utf-8").read()
                     .replace("__DATA_FILE__", "production_data_fab.json")
                     .replace("__FALLBACK_JSON__", json.dumps(data)))

    print(f"FAB  as of {data['as_of']}  hours_mode={HOURS_MODE}")
    print(f"MONTH {data['month_label']}: {mt['tons']} t / {m_goal} goal | {mt['hours']} hrs | {mt['mhton']} MH per ton")
    for r in month["rows"]:
        print("   ", r["key"], [c["tons"] for c in r["cells"]], "->", r["total"])
    print(f"WEEK  {data['week_label']}: {wt['tons']} t | {wt['hours']} hrs | {wt['mhton']} MH per ton")
    for r in week["rows"]:
        print("   ", r["key"], [c["tons"] for c in r["cells"]], "->", r["total"])


if __name__ == "__main__":
    main()
