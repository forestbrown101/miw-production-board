"""
pull_qc_by_plant.py
Tons and pieces by PLANT, from the inspection module. The API replacement for the
"DAILY SUMMARY BY JOB - ALL PLANTS" QC Final Report.

Why the inspection module and not production control: plant at MIW is captured as the
inspection Test Location, values West Bay, East Bay, Plant 4, Plant 3, Paint Bay. Those
map onto the QuickBooks Time groups P1 - West Bay, P1 - East Bay, P4 - Fab. The
ProductionShippingProductivity_Get command has 46 GroupBy values and Test Location is not
one of them, so it can never split P1 from P4. This can.

VERIFIED against 20260828 DAILY SUMMARY BY JOB - ALL PLANTS.xls, 2026-09-02:
    108 pieces and 243,100 lb, matching to the piece and to the pound, across all eight
    job-and-plant combinations.

The three rules that verification established, none of them obvious:

  1. Use the NON-condensed GetInspectionTestRecords. The condensed variant omits
     WeightEach on every record.
  2. Weight is WeightEach x the RECORD Quantity. There is also a Quantity on the
     InspectionTestSubject and it is NOT the one to multiply by; using it gives 595,384 lb
     against a true 243,100.
  3. WeightEach comes back in KILOGRAMS, tagged UOM="kg". This command takes no
     InputDisplayUnits, so there is no way to ask for pounds. Read the UOM and convert.

    python pull_qc_by_plant.py --out DATA --days 14
"""

from __future__ import annotations

import argparse
import collections
import os
import sys
import xml.etree.ElementTree as ET

from powerfab import (Config, PowerFab, PowerFabError, cmd, date_window, write_csv)

LB_PER_KG = 2.20462


def _ln(e) -> str:
    return e.tag.split("}", 1)[-1]


def _kid(e, tag) -> str:
    for c in e:
        if _ln(c) == tag:
            return (c.text or "").strip()
    return ""


def _subject(e, tag):
    for c in e:
        if _ln(c) == "InspectionTestSubject":
            for g in c:
                if _ln(g) == tag:
                    return g
    return None


def parse(block) -> list[dict]:
    """Flatten one GetInspectionTestRecords response block to tidy rows."""
    # InspectionTestID is on every record, the human title is only in the test
    # definitions that come back alongside them.
    titles: dict[str, str] = {}
    for it in block:
        if _ln(it) in ("InspectionTest", "LatestVersionInspectionTest"):
            d = {_ln(c): (c.text or "").strip() for c in it}
            tid = d.get("InspectionTestID")
            for k in ("Title", "TestTitle", "Name", "Description"):
                if tid and d.get(k):
                    titles.setdefault(tid, d[k])
                    break

    rows = []
    for e in block:
        if _ln(e) != "InspectionTestRecord":
            continue
        qty = int(_kid(e, "Quantity") or 0)

        w = _subject(e, "WeightEach")
        lb = 0.0
        if w is not None:
            v = float(w.text or 0)
            lb = v * (LB_PER_KG if (w.get("UOM") or "").lower() == "kg" else 1.0)

        tid = _kid(e, "InspectionTestID")
        rows.append({
            "Date": _kid(e, "TestDateTime")[:10],
            "Test": titles.get(tid, f"TestID {tid}"),
            "Plant": _kid(e, "TestLocation"),
            "Job": _sub_text(e, "JobNumber"),
            "MainMark": _sub_text(e, "MainMark"),
            "Sequence": _sub_text(e, "Sequence"),
            "Lot": _sub_text(e, "LotNumber"),
            "WorkPackage": _sub_text(e, "WorkPackageNumber"),
            "Finish": _sub_text(e, "Finish"),
            "Quantity": qty,
            "Pounds": round(lb * qty, 2),
            "Tons": round(lb * qty / 2000.0, 4),
            "Failed": int(_kid(e, "TestFailed") or 0),
            "Warning": int(_kid(e, "TestWarning") or 0),
            "Inspector": _kid(e, "InspectorName"),
            "TestHours": float(_kid(e, "TestHours") or 0),
        })
    return rows


def _sub_text(e, tag) -> str:
    g = _subject(e, tag)
    return (g.text or "").strip() if g is not None else ""


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=".")
    ap.add_argument("--days", type=int, default=14)
    ap.add_argument("--config", default="powerfab.ini")
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)
    start, end = date_window(args.days)
    pf = PowerFab(Config.load(args.config))

    # IncludeDetails=0 drops the per-field answer blocks, which are most of the payload
    # and which we do not need: TestFailed already carries pass or fail. One day with
    # details on was 2.6 MB.
    resp = pf.run(cmd("GetInspectionTestRecords",
                      TestDateMin=start, TestDateMax=end, IncludeDetails=0))

    blocks = resp.commands("GetInspectionTestRecords")
    rows = parse(blocks[0]) if blocks else []

    path = os.path.join(args.out, "qc_by_plant.csv")
    print(f"{'qc_by_plant.csv':36} {write_csv(rows, path):>7,} rows   "
          f"[{start} to {end}]")
    if not rows:
        print("  WARNING: no inspection records returned.", file=sys.stderr)
        return 1

    passed = [r for r in rows if not r["Failed"]]
    by_test = collections.Counter()
    for r in passed:
        by_test[r["Test"]] += r["Tons"]
    for t, tons in by_test.most_common():
        print(f"    {t:38} {tons:9,.1f} t")

    print("  plant split, passed only:")
    by_plant = collections.defaultdict(lambda: [0, 0.0])
    for r in passed:
        by_plant[r["Plant"]][0] += r["Quantity"]
        by_plant[r["Plant"]][1] += r["Tons"]
    for p, (q, t) in sorted(by_plant.items(), key=lambda x: -x[1][1]):
        print(f"    {p or '(blank)':14} {q:6,} pcs   {t:9,.1f} t")

    fails = [r for r in rows if r["Failed"]]
    if fails:
        print(f"  {len(fails)} failed inspection(s) in the window")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except PowerFabError as e:
        print(f"PowerFab error: {e}", file=sys.stderr)
        sys.exit(2)
