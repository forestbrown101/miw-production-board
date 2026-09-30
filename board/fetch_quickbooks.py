# Pull QuickBooks Time hours -> CSV that matches the manual timesheet report (all 22 columns).
#   Token: C:\MIW\qbt_token.txt
#   Live (writes today's file to the QuickBooks Time folder):  python fetch_quickbooks.py
#   Validate a day (writes test_<date>.csv here, not to the board): python fetch_quickbooks.py --day 2026-06-29
#   Quick totals:                                                   python fetch_quickbooks.py --check 2026-06-29
import urllib.request, urllib.parse, json, datetime, os, sys
from collections import defaultdict

here = os.path.dirname(os.path.abspath(__file__))

# --- where the timesheet CSVs live -------------------------------------------
# This script now runs from the Server folder, so it can no longer assume the
# QuickBooks Time folder is its own parent. run_board.bat passes it explicitly.
#   python fetch_quickbooks.py --qbdir "<...>\Daily Reports\QuickBooks Time"
# With no --qbdir it falls back to the old behaviour (parent folder), so a copy
# sitting back in the Schedule Board folder still works unchanged.
QBDIR = (sys.argv[sys.argv.index("--qbdir") + 1] if "--qbdir" in sys.argv
         else os.path.join(os.path.dirname(here), "QuickBooks Time"))
TOKEN = None
for p in [r"C:\MIW\qbt_token.txt", os.path.join(here, "qbt_token.txt")]:
    if os.path.exists(p):
        TOKEN = open(p).read().strip(); break
if not TOKEN:
    print(r"ERROR: put your token in C:\MIW\qbt_token.txt"); sys.exit(1)

BASE = "https://rest.tsheets.com/api/v1"
def api(path, params):
    url = BASE + path + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"Authorization": "Bearer " + TOKEN})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.load(r)
def fetch_all(path, key, extra=None):
    out = {}; page = 1
    while True:
        params = {"per_page": 200, "page": page}
        if extra: params.update(extra)
        d = api(path, params)
        out.update(d.get("results", {}).get(key, {}))
        if not d.get("more"): break
        page += 1
    return out

COLS = ["username","payroll_id","fname","lname","number","group","local_date","local_day",
        "local_start_time","local_end_time","tz","hours","jobcode_1","jobcode_2","class",
        "service item","tasks","location","notes","approved_status","has_flags","flag_types"]

def localtime(s):
    return s[:19].replace("T", " ") if s else ""

def get_rows(start, end):
    groups = fetch_all("/groups", "groups")
    users  = fetch_all("/users", "users")
    cfields = fetch_all("/customfields", "customfields")
    cf_id = {c.get("name"): cid for cid, c in cfields.items()}   # name -> id
    gname = {gid: g.get("name", "") for gid, g in groups.items()}
    ts = {}; jobcodes = {}; page = 1
    while True:
        d = api("/timesheets", {"start_date": start, "end_date": end,
                                "supplemental_data": "yes", "per_page": 200, "page": page})
        ts.update(d.get("results", {}).get("timesheets", {}))
        jobcodes.update(d.get("supplemental_data", {}).get("jobcodes", {}))
        if not d.get("more"): break
        page += 1
    rows = []
    for t in ts.values():
        u = users.get(str(t.get("user_id")), {})
        leaf = jobcodes.get(str(t.get("jobcode_id")), {})
        pid = leaf.get("parent_id")
        parent = jobcodes.get(str(pid)) if pid else None
        if parent:
            jc1, jc2 = parent.get("name",""), leaf.get("name","")
        else:
            jc1, jc2 = leaf.get("name",""), ""
        date = t.get("date","")
        day = datetime.date.fromisoformat(date).strftime("%a") if date else ""
        cf = t.get("customfields", {}) or {}
        appr_to = u.get("approved_to","")
        approved = "approved" if (appr_to and appr_to not in ("","0000-00-00") and date and date <= appr_to) else "unapproved"
        rows.append({
            "username": u.get("username",""), "payroll_id": u.get("payroll_id",""),
            "fname": u.get("first_name",""), "lname": u.get("last_name",""),
            "number": u.get("employee_number",""), "group": gname.get(str(u.get("group_id")), ""),
            "local_date": date, "local_day": day,
            "local_start_time": localtime(t.get("start","")), "local_end_time": localtime(t.get("end","")),
            "tz": ("%.1f" % float(t.get("tz"))) if (t.get("tz") is not None and t.get("start")) else "",
            "hours": "%.2f" % ((t.get("duration") or 0)/3600.0),
            "jobcode_1": jc1, "jobcode_2": jc2,
            "class": cf.get(cf_id.get("Class",""), ""),
            "service item": cf.get(cf_id.get("Service Item",""), ""),
            "tasks": cf.get(cf_id.get("Tasks",""), ""),
            "location": t.get("location",""), "notes": t.get("notes",""),
            "approved_status": approved, "has_flags": "", "flag_types": "",
        })
    return rows

def q(v):  # TSheets-style: quote if it contains space, comma, quote or newline
    v = "" if v is None else str(v)
    if v and any(c in v for c in ' ,"\n'):
        return '"' + v.replace('"', '""') + '"'
    return v
def write_csv(rows, path):
    with open(path, "w", newline="", encoding="utf-8") as f:
        f.write('"' + '","'.join(COLS) + '"\n')
        for r in rows:
            f.write(",".join(q(r.get(c, "")) for c in COLS) + "\n")

if "--check" in sys.argv:
    day = sys.argv[sys.argv.index("--check")+1]
    rows = get_rows(day, day); byg = defaultdict(float)
    for r in rows: byg[r["group"] or "(blank)"] += float(r["hours"])
    print("CHECK", day, "| rows:", len(rows))
    for k, v in sorted(byg.items(), key=lambda x:-x[1]): print("   %-24s %.1f" % (k, round(v,1)))
    print("   TOTAL", round(sum(byg.values()),1))
elif "--backfill" in sys.argv:
    day = sys.argv[sys.argv.index("--backfill")+1]
    rows = get_rows(day, day)
    out = os.path.join(QBDIR, "auto_timesheet_%s.csv" % day)
    write_csv(rows, out)
    print("backfilled", out, "| rows:", len(rows))
elif "--day" in sys.argv:
    day = sys.argv[sys.argv.index("--day")+1]
    rows = get_rows(day, day)
    out = os.path.join(here, "test_%s.csv" % day)
    write_csv(rows, out)
    print("wrote", out, "| rows:", len(rows))
else:
    qbdir = QBDIR
    for delta in (1, 0):  # finish yesterday, then today (self-heals a day the server missed)
        day = (datetime.date.today() - datetime.timedelta(days=delta)).isoformat()
        rows = get_rows(day, day)
        write_csv(rows, os.path.join(qbdir, "auto_timesheet_%s.csv" % day))
        print("wrote auto_timesheet_%s.csv | rows: %d | hours: %.1f" % (day, len(rows), sum(float(r["hours"]) for r in rows)))
