"""
join_hours.py
Join PowerFab tons to QuickBooks Time hours, by job and by date.

MIW does NOT track labor hours in PowerFab. The API's Hours field comes back at or near
zero and that is correct behavior, not a bug. Hours live in the QuickBooks Time daily
exports under Daily Reports\\QuickBooks Time.

    python join_hours.py --pf ..\\..\\daily_by_job_and_station.csv \\
                         --qb "C:\\...\\Daily Reports\\QuickBooks Time" \\
                         --start 2026-08-14 --end 2026-08-28 \\
                         --out ..\\..\\tons_and_hours.csv

The join key is job number plus date. PowerFab reports the job as "2402"; QuickBooks
writes it as "Job #2402" in jobcode_2, so the number is extracted with a regex. That also
catches the malformed variants documented in the Weekly HR Report audit, "Job # 2361" with
a space, "Job #2314-REV", "Job #2428-", all of which fall out of the workbook's exact-match
SUMIFS entirely. Here they land on the right job.

Conventions follow the Fabrication Board, so the numbers tie out:
  * fit tons  = the QC-FINAL station
  * cut tons  = the Cut station (P3 Cut is reported separately)
  * man hours = every P1* and P4* department, job-coded only
  * Plant 3 is reported but excluded from the headline MH per ton, because P3 charges
    hours and finals almost no tonnage, which makes its ratio meaningless
  * non-job-coded rows (lunch, PTO, tardy, no call no show) never carry a job number and
    are excluded, which is the usual explanation when hours "do not match"
"""

from __future__ import annotations

import argparse
import collections
import csv
import glob
import os
import re
import sys

JOB_RE = re.compile(r"#\s*(\d+)")
PLANT_RE = re.compile(r"\s*(P\d)\b")

# PowerFab station -> the column it feeds
STATIONS = {
    "Cut": "Cut_t",
    "P3 Cut": "P3Cut_t",
    "QC-Fit": "Fit_t",
    "QC-FINAL": "Final_t",
    "Paint": "Paint_t",
}


def read_tons(path: str) -> dict:
    out: dict = collections.defaultdict(lambda: collections.defaultdict(float))
    with open(path, encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            job, date = r.get("Job #", ""), r.get("StartDate", "")
            if not job or not date:
                continue
            col = STATIONS.get(r.get("Station", ""))
            if col:
                out[(job, date)][col] += float(r.get("Tons") or 0)
    return out


def read_hours(folder: str, start: str, end: str) -> tuple[dict, dict]:
    hours: dict = collections.defaultdict(lambda: collections.defaultdict(float))
    skipped = {"non_job_hours": 0.0, "files": 0, "rows": 0}
    pattern = os.path.join(folder, "auto_timesheet_*.csv")
    for f in sorted(glob.glob(pattern)):
        m = re.search(r"(\d{4}-\d{2}-\d{2})", os.path.basename(f))
        if not m or not (start <= m.group(1) <= end):
            continue
        skipped["files"] += 1
        with open(f, encoding="utf-8-sig", errors="replace") as fh:
            for r in csv.DictReader(fh):
                skipped["rows"] += 1
                h = float(r.get("hours") or 0)
                jm = JOB_RE.search(r.get("jobcode_2") or "")
                if not jm:
                    skipped["non_job_hours"] += h
                    continue
                job, date = jm.group(1), (r.get("local_date") or "").strip()
                pm = PLANT_RE.match(r.get("group") or "")
                plant = pm.group(1) if pm else "Other"
                hours[(job, date)][f"Hours_{plant}"] += h
                hours[(job, date)]["Hours_Total"] += h
    return hours, skipped


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--pf", required=True, help="daily_by_job_and_station.csv from the API pull")
    ap.add_argument("--qb", required=True, help="QuickBooks Time folder")
    ap.add_argument("--start", required=True)
    ap.add_argument("--end", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()

    tons = read_tons(a.pf)
    hours, info = read_hours(a.qb, a.start, a.end)

    keys = sorted(set(tons) | set(hours))
    plant_cols = sorted({c for v in hours.values() for c in v if c != "Hours_Total"})
    cols = (["Date", "Job"] + list(dict.fromkeys(STATIONS.values()))
            + plant_cols + ["Hours_Total", "MH_per_ton_fit"])

    rows = []
    for job, date in keys:
        t, h = tons.get((job, date), {}), hours.get((job, date), {})
        row = {"Date": date, "Job": job}
        for c in dict.fromkeys(STATIONS.values()):
            row[c] = round(t.get(c, 0.0), 4)
        for c in plant_cols:
            row[c] = round(h.get(c, 0.0), 2)
        row["Hours_Total"] = round(h.get("Hours_Total", 0.0), 2)
        # Board convention: P1 and P4 only over QC-FINAL tons.
        billable = h.get("Hours_P1", 0.0) + h.get("Hours_P4", 0.0)
        row["MH_per_ton_fit"] = round(billable / row["Final_t"], 2) if row["Final_t"] > 0.05 else ""
        rows.append(row)

    with open(a.out, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)

    tons_only = sorted({k[0] for k in tons} - {k[0] for k in hours})
    hours_only = sorted({k[0] for k in hours} - {k[0] for k in tons})
    print(f"{info['files']} QuickBooks files, {info['rows']:,} timesheet rows")
    print(f"{len(rows):,} joined rows -> {a.out}")
    print(f"non-job-coded hours excluded: {info['non_job_hours']:,.1f}")
    if tons_only:
        print(f"WARNING jobs with tons but NO hours: {tons_only}")
    if hours_only:
        print(f"note: jobs with hours but no scanned tons: {hours_only}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
