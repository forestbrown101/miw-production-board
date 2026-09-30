"""
pull_paint_by_plant.py
Paint tons + pieces by PLANT, straight from the API. Same call and same trick as
pull_cut_by_saw.py: the Paint station has no plant dimension in PowerFab, so the only
link to a plant is who signed off the station (Station Completed By).

Plant 2 paint shop started 2026-09-15. The only person logging paint there is Jvaughn,
so Jvaughn -> P2 and every other login -> P1 (the original paint bay). Add a name to
OP2PLANT when a second Plant 2 painter starts logging.

UNITS: with Station Completed By in the GroupBy the weights come back in KILOGRAMS while
the response echoes Imperial (data skill Rule 2). Always convert.

    py pull_paint_by_plant.py --out DATA --days 14                 (5-minute trailing window)
    py pull_paint_by_plant.py --out DATA --start 2026-09-01        (backfill; end defaults today)
"""
from __future__ import annotations
import argparse, csv, os, sys, datetime
from powerfab import Config, PowerFab, PowerFabError, cmd

LB_PER_KG = 2.20462
# login (case-insensitive) -> plant. Anyone not listed paints at Plant 1.
OP2PLANT = {"jvaughn": "P2"}
DEFAULT_PLANT = "P1"


def _ln(e):
    t = getattr(e, "tag", e)
    return t.split("}")[-1] if isinstance(t, str) else ""


def plant_of(user: str) -> str:
    return OP2PLANT.get((user or "").strip().lower(), DEFAULT_PLANT)


def parse(block):
    conv = LB_PER_KG          # response lies about units; force kg -> lb (Rule 2)
    rows = []
    for res in [e for e in block.iter() if _ln(e) == "Result"]:
        date = ""
        for c in res:
            if _ln(c) == "StartDate": date = (c.text or "")[:10]; break
        for g in [c for c in res if _ln(c) == "Group"]:
            fields = {f.get("FieldName"): (f.text or "") for f in g.findall("{*}Field")}
            if fields.get("Station") != "Paint": continue
            user = fields.get("Station Completed By", "")
            w = q = 0.0
            for c in g:
                if _ln(c) == "Weight": w = float(c.text or 0)
                if _ln(c) == "Quantity": q = int(float(c.text or 0))
            lb = w * conv
            rows.append({"Date": date, "Plant": plant_of(user), "User": user,
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
    ap.add_argument("--name", default="paint_by_plant.csv")
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
        p2 = sum(x["Tons"] for x in got if x["Plant"] == "P2")
        print(f"[{m0} .. {m1}]  paint rows {len(got)}  (P2 {p2:,.1f} t)", flush=True)

    path = os.path.join(a.out, a.name)
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=["Date","Plant","User","Job","Pounds","Tons","Pieces"])
        w.writeheader(); w.writerows(rows)
    print(f"\nDONE {a.name} {len(rows):,} rows  [{start} to {end}]")
    return 0


if __name__ == "__main__":
    try: sys.exit(main())
    except PowerFabError as ex: print(f"PowerFab error: {ex}", file=sys.stderr); sys.exit(2)
