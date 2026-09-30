"""MIW QC board builder (sample, 2026-09-29). A breakdown of the Production Board's QC numbers:
by area (East Bay, West Bay, Plant 4), by job, by day, and each job's fabbed vs released tons.
No forecasting. Same source as the Production Board: feed_fit.csv (QC-Final passed), tons = lb / 2000.

  argv: FEED_FIT FEED_RELEASED CM_JSON TEMPLATE OUT
  FEED_RELEASED  SCHEDULES\\Daily Reports\\PowerFab API\\feed_released.csv  (released BOM tons per job)
  CM_JSON        only for job names (Control Module job list); optional
"""
import csv, json, os, sys, datetime as dt
from collections import defaultdict

FEED_FIT, FEED_REL, CM_JSON, TEMPLATE, OUT = sys.argv[1:6]
NOW = dt.datetime.fromisoformat(os.environ["NOW"]) if os.environ.get("NOW") else dt.datetime.now()
TODAY = NOW.date()
AREAS = [("EB", "East Bay", "P1"), ("WB", "West Bay", "P1"), ("P4", "Plant 4", "P4")]
LOC_AREA = {"East Bay": "EB", "West Bay": "WB", "C-Beams": "WB", "SHOP": "WB", "Paint Bay": "EB", "Plant 4": "P4"}
WEEKS_BACK, MONTHS_BACK, TOP_JOBS = 5, 2, 8

rows = []                                             # (date, job, area, tons, pieces)
with open(FEED_FIT, newline="", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        a = LOC_AREA.get(r["Test Location"])          # Plant 3 is collected, not shown (same as the board)
        if a: rows.append((dt.date.fromisoformat(r["Report Date"][:10]), r["Job #"], a, float(r["Weight"] or 0) / 2000, int(float(r["Total"] or 0))))
released = {}
with open(FEED_REL, newline="", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        released[r["Job #"]] = float(r["Released_Tons"] or 0)
names = {}
if os.path.exists(CM_JSON):
    for j in json.load(open(CM_JSON, encoding="utf-8")).get("jobs", []): names[j["job"]] = j.get("project", "")
lifetime = defaultdict(float)
for (_, j, _, t, _) in rows: lifetime[j] += t          # through today; P3 excluded, like every board number

def r1(v): return round(v, 1)
def period(scope, ps, pe, label):
    if scope == "week":
        buckets = [(ps + dt.timedelta(days=i), ps + dt.timedelta(days=i)) for i in range(7)]
        blabels = [(b[0].strftime("%a"), b[0].strftime("%-m/%-d")) for b in buckets]
    else:
        buckets, x = [], ps - dt.timedelta(days=ps.weekday())
        while x <= pe:
            buckets.append((max(x, ps), min(x + dt.timedelta(days=6), pe))); x += dt.timedelta(days=7)
        blabels = [("Wk " + b[0].strftime("%-m/%-d"), "") for b in buckets]
    ti = next((i for i, b in enumerate(buckets) if b[0] <= TODAY <= b[1]), -1)
    P = [r for r in rows if ps <= r[0] <= pe]
    area = {}
    for k, n, pl in AREAS:
        A = [r for r in P if r[2] == k]
        daily = [r1(sum(r[3] for r in A if b[0] <= r[0] <= b[1])) for b in buckets]
        daily_pc = [sum(r[4] for r in A if b[0] <= r[0] <= b[1]) for b in buckets]
        byjob = defaultdict(float)
        for r in A: byjob[r[1]] += r[3]
        top = max(byjob.items(), key=lambda x: x[1]) if byjob else None
        area[k] = {"name": n, "plant": pl, "tons": r1(sum(r[3] for r in A)), "pieces": sum(r[4] for r in A),
                   "daily": daily, "daily_pc": daily_pc, "jobs": len(byjob),
                   "top_job": top[0] if top else None, "top_tons": r1(top[1]) if top else None}
    jobs = defaultdict(lambda: {"EB": 0.0, "WB": 0.0, "P4": 0.0, "pc": 0})
    for r in P: jobs[r[1]][r[2]] += r[3]; jobs[r[1]]["pc"] += r[4]
    jl = []
    for j, v in sorted(jobs.items(), key=lambda x: -(x[1]["EB"] + x[1]["WB"] + x[1]["P4"]))[:TOP_JOBS]:
        tot = v["EB"] + v["WB"] + v["P4"]
        rel = released.get(j) or None
        done = lifetime[j] if closed_through(pe) is None else sum(r[3] for r in rows if r[1] == j and r[0] <= pe)
        jl.append({"job": j, "name": names.get(j, ""), "EB": r1(v["EB"]), "WB": r1(v["WB"]), "P4": r1(v["P4"]),
                   "tons": r1(tot), "pieces": v["pc"], "fabbed": r1(done), "released": r1(rel) if rel else None,
                   "left": r1(max(0.0, rel - done)) if rel else None, "pct": min(100, round(done / rel * 100)) if rel else None})
    tot = sum(a["tons"] for a in area.values())
    return {"scope": scope, "label": label, "period_label": scope, "today_label": min(TODAY, pe).strftime("%-m/%-d"),
            "days": [{"label": a, "date": b} for a, b in blabels], "today_index": ti,
            "total": r1(tot), "pieces": sum(a["pieces"] for a in area.values()),
            "p1": r1(area["EB"]["tons"] + area["WB"]["tons"]), "area": area, "jobs": jl,
            "generated_at": NOW.strftime("%Y-%m-%d %H:%M")}
def closed_through(pe): return None if pe >= TODAY else pe

weeks, months = [], []
m0 = TODAY - dt.timedelta(days=TODAY.weekday())
for i in range(WEEKS_BACK - 1, -1, -1):
    ps = m0 - dt.timedelta(days=7 * i); pe = ps + dt.timedelta(days=6)
    span = f"{ps.strftime('%b %-d')}–{pe.day}" if ps.month == pe.month else f"{ps.strftime('%b %-d')} – {pe.strftime('%b %-d')}"
    label = f"Week {ps.isocalendar()[1]} · {span}"
    weeks.append({"key": ps.isoformat(), "label": label, "data": period("week", ps, pe, label)})
for i in range(MONTHS_BACK - 1, -1, -1):
    y, m = TODAY.year, TODAY.month - i
    if m <= 0: y, m = y - 1, m + 12
    ps = dt.date(y, m, 1); pe = dt.date(y + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1)
    months.append({"key": ps.isoformat(), "label": ps.strftime("%B %Y"), "data": period("month", ps, pe, ps.strftime("%B %Y"))})
H = {"generated_at": NOW.strftime("%Y-%m-%d %H:%M"), "weeks": weeks, "months": months,
     "current_week": m0.isoformat(), "current_month": TODAY.replace(day=1).isoformat()}
open(OUT, "w", encoding="utf-8").write(open(TEMPLATE, encoding="utf-8").read().replace("__QC_JSON__", json.dumps(H)))
print("wrote", OUT)
