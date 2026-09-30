"""
pull_cut_by_saw.py
Cut tons + pieces by SAW (-> plant), straight from the API. Replaces the operator-filtered
saw report EXPORTS that keep coming back blank.

Cut has no plant or machine dimension in PowerFab; the only link to a saw is who signed off
the Cut station (Station Completed By). At MIW that value is mostly the saw-login account
(Saw1, Saw2, Saw3, Saw5, Saw5Night), plus a few people. We map operator -> saw -> plant,
drop the angle line and unmapped names.

GOTCHA: with Station Completed By in the GroupBy the weights come back in KILOGRAMS even
though other productivity calls return pounds. Read the response UOM and convert.

    py pull_cut_by_saw.py --out DATA --start 2025-01-01     (backfill; end defaults today)
    py pull_cut_by_saw.py --out DATA --days 14              (nightly trailing window)
"""
from __future__ import annotations
import argparse, csv, os, sys, datetime
from powerfab import Config, PowerFab, PowerFabError, cmd

LB_PER_KG = 2.20462
OP2SAW = {"Saw1":"Saw 1","Saw2":"Saw 2","Saw3":"Saw 3","Saw4":"Saw 4",
          "Saw5":"Saw 5","Saw5Night":"Saw 5","Lkelly":"Saw 5",
          # legacy person logins seen in 2025, kept so history splits too
          "Dburroughs":"Saw 1","Tsmiling":"Saw 2","Mnuckols":"Saw 3","KPearson":"Saw 5",
          # Forest 2026-09-29: Afilyaw signs off PLATE cut. Not a saw, but it gets its own row
          # on the board. Plant left blank until confirmed, so it counts in the cut total but
          # not in the P1/P4 split.
          "Afilyaw":"Plate"}
SAW2PLANT = {"Saw 1":"P1","Saw 2":"P1","Saw 3":"P1","Saw 4":"P1","Saw 5":"P4","Plate":""}
# ONE-DAY EXCEPTIONS: (login, date) -> saw, for someone filling in on a machine. Applies to that
# date only, never becomes a standing rule. Once the date ages out of the 14 day window the
# history file keeps the corrected row.
OP2SAW_ONE_DAY = {
    ("Dmccrea", "2026-09-28"): "Saw 1",   # Forest 2026-09-29: Dave Mccrea filled in on Saw 1, exception
}
# 2026-09-29: logins that signed off Cut but map to no saw. Still dropped here (the feed carries
# them as "(unattributed)"), but now written to cut_unmapped_logins.csv so the name is visible.
# Job 2319 put 38 t on 9/21 and 32 t on 9/28 through a login that is not in OP2SAW.
UNMAPPED = []


def _ln(e):
    t = getattr(e, "tag", e)
    return t.split("}")[-1] if isinstance(t, str) else ""


def parse(block):
    # UNITS (confirmed 2026-09-04): ProductionShippingProductivity_Get returns weights in
    # KILOGRAMS but ECHOES <WeightUOM>Imperial</WeightUOM> (the response lies; see the data
    # skill Rule 2). The old uom=="kg" test therefore never fired and cut came out 2.2x low.
    # Proven against fit tons (inspection, correct): cut/fit ran ~0.45 until this was forced,
    # and cut*2.20462 lands right on fit. So always convert kg->lb for this command.
    conv = LB_PER_KG
    rows = []
    for res in [e for e in block.iter() if _ln(e) == "Result"]:
        date = ""
        for c in res:
            if _ln(c) == "StartDate": date = (c.text or "")[:10]; break
        for g in [c for c in res if _ln(c) == "Group"]:
            fields = {f.get("FieldName"): (f.text or "") for f in g.findall("{*}Field")}
            if fields.get("Station") != "Cut": continue
            login = fields.get("Station Completed By", "")
            saw = OP2SAW_ONE_DAY.get((login, date)) or OP2SAW.get(login)
            w = q = 0.0
            for c in g:
                if _ln(c) == "Weight": w = float(c.text or 0)
                if _ln(c) == "Quantity": q = int(float(c.text or 0))
            lb = w * conv
            if not saw:                              # AngleLine + unmapped people
                UNMAPPED.append({"Date": date, "Login": fields.get("Station Completed By", ""),
                                 "Job": fields.get("Job #", ""), "Tons": round(lb/2000.0, 4), "Pieces": q})
                continue
            rows.append({"Date": date, "Saw": saw, "Plant": SAW2PLANT[saw],
                         "Job": fields.get("Job #", ""), "Pounds": round(lb, 2),
                         "Tons": round(lb/2000.0, 4), "Pieces": q})
    return rows


def months(s, e):
    d = s.replace(day=1)
    while d <= e:
        nxt = (d.replace(day=28)+datetime.timedelta(days=4)).replace(day=1)
        yield d.isoformat(), min(nxt-datetime.timedelta(days=1), e).isoformat()
        d = nxt


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--start"); ap.add_argument("--end", default=datetime.date.today().isoformat())
    ap.add_argument("--days", type=int, default=14)
    ap.add_argument("--name", default="cut_by_saw.csv")
    ap.add_argument("--config", default="powerfab.ini")
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    pf = PowerFab(Config.load(a.config))
    end = datetime.date.fromisoformat(a.end)
    start = datetime.date.fromisoformat(a.start) if a.start else end - datetime.timedelta(days=a.days)

    rows = []
    for m0, m1 in months(start, end):
        r = pf.run(cmd("ProductionShippingProductivity_Get",
                       GetOptions=dict(StatusType="productivity-production-station", DateSummary="day",
                                       StartDate=m0, EndDate=m1,
                                       GroupBy=["Job #", "Station", "Station Completed By"])))
        blocks = r.commands("ProductionShippingProductivity_Get")
        got = parse(blocks[0]) if blocks else []
        rows += got
        print(f"[{m0} .. {m1}]  cut rows {len(got)}", flush=True)

    path = os.path.join(a.out, a.name)
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=["Date","Saw","Plant","Job","Pounds","Tons","Pieces"])
        w.writeheader(); w.writerows(rows)
    print(f"\nDONE {a.name} {len(rows):,} rows  [{start} to {end}]")
    upath = os.path.join(a.out, "cut_unmapped_logins.csv")
    with open(upath, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=["Date","Login","Job","Tons","Pieces"])
        w.writeheader(); w.writerows(sorted(UNMAPPED, key=lambda x: (x["Date"], -x["Tons"])))
    by = {}
    for u in UNMAPPED: by[u["Login"]] = by.get(u["Login"], 0.0) + u["Tons"]
    for k, v in sorted(by.items(), key=lambda x: -x[1]):
        print(f"UNMAPPED login {k or '(blank)'}: {v:,.1f} t in window -> add to OP2SAW if it is a saw")
    return 0


if __name__ == "__main__":
    try: sys.exit(main())
    except PowerFabError as ex: print(f"PowerFab error: {ex}", file=sys.stderr); sys.exit(2)
