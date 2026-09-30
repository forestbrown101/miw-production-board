"""
backfill_history.py
One-time historical pull. Loops month by month so no single response is huge or times out.
Writes two permanent history files the board can read for any past week or month:

  qc_history.csv       fit tons + pieces BY PLANT (inspection module, kg->lb handled by parse)
  station_history.csv  tons by job x station x day (Cut = cut total, QC-FINAL = fit total)

Cut has no plant split before the saw exports (~Jul 2026); station_history carries the cut
TOTAL for those older periods, which is what the historical view shows.

    py backfill_history.py --out DATA --start 2025-01-01            (end defaults to today)
"""
from __future__ import annotations
import argparse, os, sys, datetime
from powerfab import Config, PowerFab, PowerFabError, add_tons, cmd, write_csv
import pull_qc_by_plant as qcp


def months(s: datetime.date, e: datetime.date):
    d = s.replace(day=1)
    while d <= e:
        nxt = (d.replace(day=28) + datetime.timedelta(days=4)).replace(day=1)
        yield d.isoformat(), min(nxt - datetime.timedelta(days=1), e).isoformat()
        d = nxt


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--start", default="2025-01-01")
    ap.add_argument("--end", default=datetime.date.today().isoformat())
    ap.add_argument("--config", default="powerfab.ini")
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    pf = PowerFab(Config.load(a.config))
    s = datetime.date.fromisoformat(a.start); e = datetime.date.fromisoformat(a.end)

    qc_rows, st_rows = [], []
    for m0, m1 in months(s, e):
        print(f"[{m0} .. {m1}]", flush=True)
        # station: cut/fit totals by job x station x day
        r = pf.run(cmd("ProductionShippingProductivity_Get",
                       GetOptions=dict(StatusType="productivity-production-station",
                                       DateSummary="day", StartDate=m0, EndDate=m1,
                                       GroupBy=["Job #", "Station"])))
        srows = r.summary_rows("ProductionShippingProductivity_Get")
        # UNITS: ProductionShippingProductivity_Get returns KILOGRAMS but echoes
        # WeightUOM=Imperial. The response lies. pull_daily.py was fixed for this on
        # 2026-09-04 (force_uom="Metric") but THIS FILE WAS NOT, so every backfill kept
        # writing kg into station_history.csv. That is why a full rebuild on 2026-09-08
        # made cut read 14,829 t against 31,480 t of fit, a ratio of 0.47 = 1/2.20462.
        # Do not "read the echoed unit" here. It is wrong for this one command.
        add_tons(srows, "Metric")
        st_rows += srows
        # inspection: fit by plant (WeightEach -> lb inside parse)
        r2 = pf.run(cmd("GetInspectionTestRecords", TestDateMin=m0, TestDateMax=m1, IncludeDetails=0))
        blocks = r2.commands("GetInspectionTestRecords")
        qrows = qcp.parse(blocks[0]) if blocks else []
        qc_rows += qrows
        print(f"   station {len(srows):>4}   inspection {len(qrows):>5}", flush=True)

    nq = write_csv(qc_rows, os.path.join(a.out, "qc_history.csv"))
    ns = write_csv(st_rows, os.path.join(a.out, "station_history.csv"))
    print(f"\nDONE  qc_history.csv {nq:,} rows   station_history.csv {ns:,} rows   [{a.start} to {a.end}]")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except PowerFabError as ex:
        print(f"PowerFab error: {ex}", file=sys.stderr); sys.exit(2)
