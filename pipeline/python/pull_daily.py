"""
pull_daily.py
The nightly MIW pull. Run it on a schedule, it drops CSVs where the boards already look.

This is what replaces the PowerFab Automated Events. One run, one round trip for most of
it, and if anything is wrong you get an error instead of a blank report.

    python pull_daily.py --out "C:\\MIW\\data"
"""

import argparse
import os
import sys

from powerfab import (IMPERIAL, Config, PowerFab, PowerFabError, add_tons, cmd,
                      date_window, write_csv)

# The server caps a GetInventorySummary response and flags LimitHasMore. Observed cap was
# 5,000 rows, which silently truncated the first pull. Page until it says there is no more.
INVENTORY_PAGE = 5000
INVENTORY_MAX_PAGES = 40


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=".", help="folder to write CSVs into")
    ap.add_argument("--days", type=int, default=14, help="trailing window, days")
    ap.add_argument("--config", default="powerfab.ini")
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)
    start, end = date_window(args.days)
    pf = PowerFab(Config.load(args.config))

    v = pf.version()
    print(f"PowerFab API {v.get('MajorVersion')} {v.get('MinorVersion')} "
          f"(interface {v.get('InterfaceVersionNumber')})")

    # --- one round trip for the three date-scoped feeds ----------------------
    resp = pf.run(
        # [0] daily tons per job and per station, for the boards.
        #     GroupBy repeats, so both dimensions come back in one command.
        cmd("ProductionShippingProductivity_Get",
            GetOptions=dict(
                StatusType="productivity-production-station",
                DateSummary="day",
                StartDate=start,
                EndDate=end,
                GroupBy=["Job #", "Station"],
            )),

        # [1] material received by job and day, for the Material Flow dashboard
        cmd("PurchaseOrderSummary_Get",
            GetOptions=dict(
                SummaryType="receiving-history",
                DateSummary="day",
                StartDate=start,
                EndDate=end,
                GroupBy=["Job #"],
            )),

        # [2] the exact location spellings PowerFab is using
        cmd("GetInventoryLocations"),
    )

    # UNITS: ProductionShippingProductivity_Get returns KILOGRAMS but echoes WeightUOM=Imperial
    # (the response lies). Confirmed 2026-09-04 against fit tons: station cut ran ~0.45x of fit
    # (impossible, cut precedes final) until this was forced to metric. force_uom overrides the
    # lying echo so Tons convert kg->lb. Purchasing + inventory keep reading their own echoed unit.
    n = pf_write(resp, "ProductionShippingProductivity_Get", args.out,
                 "daily_by_job_and_station.csv", kind="summary", force_uom="Metric")
    pf_write(resp, "PurchaseOrderSummary_Get", args.out,
             "received_by_job.csv", kind="summary")
    pf_write(resp, "GetInventoryLocations", args.out,
             "inventory_locations.csv", kind="flat")

    # --- inventory, paged, and in POUNDS ------------------------------------
    # Commands that accept InputDisplayUnits and are not given it default to METRIC.
    # The first pull came back in kg and mm without saying so.
    rows, offset = [], 0
    for page in range(INVENTORY_MAX_PAGES):
        r = pf.run(cmd("GetInventorySummary",
                       _repeat={"SummaryField": ["Location", "SecondaryLocation",
                                                 "Shape", "Grade", "Dimensions"]},
                       InputDisplayUnits=IMPERIAL,
                       Offset=offset, Limit=INVENTORY_PAGE))
        batch = r.inventory_rows()
        rows += batch
        if not r.has_more() or not batch:
            break
        offset += len(batch)
        print(f"  inventory page {page + 1}: {len(rows):,} rows so far")
    else:
        print(f"  WARNING: stopped at {INVENTORY_MAX_PAGES} pages, there may be more",
              file=sys.stderr)

    # We SEND an explicit Imperial InputDisplayUnits block, and PowerFab ignored it:
    # the response still came back tagged kg. So convert from what the response says,
    # never from what we asked for.
    inv_uom = rows[0].get("InStockWeight_UOM") if rows else None
    add_tons(rows, inv_uom)
    path = os.path.join(args.out, "inventory_by_location.csv")
    print(f"{'inventory_by_location.csv':36} {write_csv(rows, path):>7,} rows"
          f"   [weight: {inv_uom or '?'}]")

    return 0


def pf_write(resp, command, out_dir, fname, kind, force_uom=None):
    path = os.path.join(out_dir, fname)
    rows = (resp.summary_rows(command) if kind == "summary" else resp.rows(command))
    # Read the units the response reports rather than assuming. These commands disagree
    # with each other: the purchase order summary has been seen returning metric in the same
    # session. ProductionShippingProductivity_Get ECHOES Imperial but actually returns kg, so
    # its caller passes force_uom to override the lying echo.
    uom = force_uom or resp.units(command).get("WeightUOM")
    add_tons(rows, uom)
    n = write_csv(rows, path)
    print(f"{fname:36} {n:>7,} rows" + (f"   [weight: {uom}]" if uom else ""))
    # A zero row pull is the exact failure mode the .mrt reports had. Say so loudly.
    if n == 0:
        print(f"  WARNING: {command} returned no rows. Check the window and the options "
              f"before trusting any board built on this.", file=sys.stderr)
    return n


if __name__ == "__main__":
    try:
        sys.exit(main())
    except PowerFabError as e:
        print(f"PowerFab error: {e}", file=sys.stderr)
        sys.exit(2)
