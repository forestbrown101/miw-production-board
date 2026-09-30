"""
merge_history.py
Upsert the fresh trailing pull into the permanent history files, by date. Dates present in
the new pull replace those dates in history; older dates are kept; new dates are appended.
Idempotent: a missed run or a double run cannot corrupt history.

    py merge_history.py --qc DATA/qc_by_plant.csv --qc-hist DATA/qc_history.csv
                        --station DATA/daily_by_job_and_station.csv --station-hist DATA/station_history.csv
"""
from __future__ import annotations
import argparse, csv, os, sys


def _read(path):
    if not os.path.exists(path):
        return [], []
    with open(path, newline="", encoding="utf-8") as fh:
        r = csv.DictReader(fh)
        return list(r), (r.fieldnames or [])


def _upsert(new_path, hist_path, date_col):
    new, nf = _read(new_path)
    hist, hf = _read(hist_path)
    if not new:
        print(f"  {os.path.basename(new_path)}: no new rows, history untouched")
        return
    fields = hf or nf
    fresh_dates = {(r.get(date_col) or "")[:10] for r in new}
    kept = [r for r in hist if (r.get(date_col) or "")[:10] not in fresh_dates]
    merged = kept + new
    merged.sort(key=lambda r: (r.get(date_col) or ""))
    tmp = hist_path + ".tmp"
    with open(tmp, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=fields, extrasaction="ignore")
        w.writeheader(); w.writerows(merged)
    os.replace(tmp, hist_path)
    print(f"  {os.path.basename(hist_path)}: {len(kept):,} kept + {len(new):,} refreshed "
          f"({len(fresh_dates)} days) = {len(merged):,} rows")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--qc"); ap.add_argument("--qc-hist")
    ap.add_argument("--station"); ap.add_argument("--station-hist")
    ap.add_argument("--cut"); ap.add_argument("--cut-hist")
    ap.add_argument("--paint"); ap.add_argument("--paint-hist")
    a = ap.parse_args()
    print("merging trailing pull into history...")
    if a.qc and a.qc_hist: _upsert(a.qc, a.qc_hist, "Date")
    if a.station and a.station_hist: _upsert(a.station, a.station_hist, "StartDate")
    if a.cut and a.cut_hist: _upsert(a.cut, a.cut_hist, "Date")
    if a.paint and a.paint_hist: _upsert(a.paint, a.paint_hist, "Date")
    return 0


if __name__ == "__main__":
    sys.exit(main())
