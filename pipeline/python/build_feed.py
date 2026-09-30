"""
build_feed.py
ONE canonical production feed, straight off the PowerFab API, consumed by BOTH the
Production Board and the WEEK ## SCHEDULES workbook.

Why this exists: the board and the workbook each had their own reader with their own
rules, so they drifted. Every mapping rule now lives HERE, once. Both consumers read
what this writes.

Design rules:
  1. NEVER filter at the source. P3 is in the feed. The board drops it at DISPLAY.
  2. An unmapped Test Location is a LOUD FAILURE, never a silent drop. That is how
     C-Beams (4,552.7 tons over 15 months) sat outside the board unnoticed.
  3. Cut TOTALS come from the station feed. Saw attribution comes from cut_by_saw,
     which is LOSSY BY DESIGN: it maps operator -> saw and drops anyone unmapped,
     which is 26% of all cut tonnage. The unattributed remainder is CARRIED as its
     own row, never dropped and never spread across the saws.
  4. Cut must be >= fit over any long window, because a piece is cut before it is
     finaled. If it is not, either the API returned kilograms and nobody converted,
     or something is being dropped. Refuse to publish either way.

Usage:
    py build_feed.py --data DATA_DIR [--out DIR] [--assume-kg] [--quiet]
"""
from __future__ import annotations
import argparse, csv, os, sys, json, collections

LB_PER_KG = 2.20462
FINAL_TESTS = ("VT - Final Visual Weld Inspection", "Final Visual Inspection")

# --- THE CANONICAL MAP. Every plant question in the company resolves here. ---
# Decided with Forest 2026-09-08:
#   C-Beams no longer exists (2025-01 to 2026-03, 4,552.7 t). Folds into P1,
#     reports as West Bay.
#   Shop 3 is Plant 3 under its old name. Shop 3 ends 2026-02, Plant 3 begins 2026-02.
#   SHOP is 68.9 t across 2025, when P1 was the only plant running. -> P1.
LOCATION = {
    "East Bay":  ("P1", "East Bay"),
    "West Bay":  ("P1", "West Bay"),
    "Paint Bay": ("P1", "Paint Bay"),
    "C-Beams":   ("P1", "West Bay"),
    "SHOP":      ("P1", "West Bay"),
    "Shop 3":    ("P3", "Plant 3"),
    "Plant 3":   ("P3", "Plant 3"),
    "Plant 4":   ("P4", "Plant 4"),
}

UNATTRIBUTED = "(unattributed)"


def rd(path):
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def num(v):
    try: return float(str(v).replace(",", "").strip())
    except Exception: return 0.0


REJECTS = []
def jobnum(v, where, tons):
    """PowerFab carries junk job codes. Keep them OUT of the feed and record them,
    because Power Query types Job # as a number and one bad row errors the load."""
    j = str(v or "").strip()
    if j.isdigit():
        return j
    if j:
        REJECTS.append((where, j, round(tons, 4)))
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--out", default=None)
    ap.add_argument("--assume-kg", action="store_true",
                    help="station history predates the kg->lb fix; convert on read. "
                         "TEMPORARY, re-run the backfill with the fixed pull instead.")
    ap.add_argument("--quiet", action="store_true")
    ap.add_argument("--allow-shrink", action="store_true",
                    help="permit a feed much smaller than the last published one. "
                         "ONLY after a deliberate history rebuild.")
    a = ap.parse_args()
    out = a.out or a.data
    os.makedirs(out, exist_ok=True)
    def say(*m):
        if not a.quiet: print(*m)

    conv = LB_PER_KG if a.assume_kg else 1.0

    # ---------------- FIT, from the inspection module ----------------
    qc = rd(os.path.join(a.data, "qc_history.csv"))
    if not qc:
        print("FATAL: qc_history.csv not found or empty in", a.data); return 2

    unmapped = collections.Counter()
    fit = collections.defaultdict(lambda: [0.0, 0.0])
    fit_tons_total = 0.0
    for r in qc:
        if r.get("Test") not in FINAL_TESTS:  continue
        if num(r.get("Failed")) != 0:         continue   # keep the passing record only
        loc = (r.get("Plant") or "").strip()
        if loc not in LOCATION:
            unmapped[loc] += 1
            continue
        plant, shown = LOCATION[loc]
        job = jobnum(r.get("Job"), "fit", num(r.get("Tons")))
        if job is None: continue
        k = (r["Date"][:10], job, shown, plant)
        fit[k][0] += num(r.get("Quantity"))
        fit[k][1] += num(r.get("Pounds"))
        fit_tons_total += num(r.get("Tons"))

    if unmapped:
        print("\n*** UNMAPPED TEST LOCATIONS, NOTHING WAS PUBLISHED ***")
        for k, v in unmapped.most_common():
            print(f"      {k!r}  {v} rows")
        print("    Add each one to LOCATION in build_feed.py and re-run.")
        print("    Do not delete this check. Silent drops are how C-Beams vanished.\n")
        return 2

    # ---------------- CUT ----------------
    # TOTAL per date+job comes from the station feed, which is complete.
    cut_total = collections.defaultdict(lambda: [0.0, 0.0])   # (date, job) -> [pieces, lb]
    cut_tons_total = 0.0
    paint = collections.defaultdict(lambda: [0.0, 0.0])
    paint_tons_total = 0.0
    for r in rd(os.path.join(a.data, "station_history.csv")):
        st = (r.get("Station") or "").strip()
        if st not in ("Cut", "Paint"): continue
        job = jobnum(r.get("Job #"), st.lower(), num(r.get("Tons")))
        if job is None: continue          # blank Job # is the grand-total row
        t = num(r.get("Tons")) * conv
        k = (r["StartDate"][:10], job)
        tgt = cut_total if st == "Cut" else paint
        tgt[k][0] += num(r.get("Quantity"))
        tgt[k][1] += t * 2000.0
        if st == "Cut": cut_tons_total += t
        else:           paint_tons_total += t

    # SAW ATTRIBUTION, lossy on purpose. Operator -> saw, unmapped operators dropped.
    # Roughly a quarter of cut tonnage has no saw. That remainder is carried, not lost.
    saw_rows = rd(os.path.join(a.data, "cut_by_saw_history.csv"))
    attributed = collections.defaultdict(lambda: [0.0, 0.0])  # (date, job, saw, plant)
    attr_by_key = collections.defaultdict(float)              # (date, job) -> lb
    for r in saw_rows:
        job = jobnum(r.get("Job"), "cut-saw", num(r.get("Tons")))
        if job is None: continue
        d = r["Date"][:10]
        lb = num(r.get("Tons")) * conv * 2000.0
        attributed[(d, job, r.get("Saw") or "", r.get("Plant") or "")][0] += num(r.get("Pieces"))
        attributed[(d, job, r.get("Saw") or "", r.get("Plant") or "")][1] += lb
        attr_by_key[(d, job)] += lb

    # Build the cut feed: saw rows where we have them, plus the remainder per date+job.
    cut = {}
    for k, (pcs, lb) in attributed.items():
        cut[k] = [pcs, lb]
    unattr_lb = 0.0
    for (d, job), (pcs, lb) in cut_total.items():
        rem = lb - attr_by_key.get((d, job), 0.0)
        if rem > 1.0:                     # more than a pound, worth carrying
            cut[(d, job, UNATTRIBUTED, "")] = [0.0, rem]
            unattr_lb += rem
    # If a saw attributes MORE than the station feed for the same date, the STATION
    # feed is wrong for that date, almost always an unconverted kg day. Say so loudly.
    over = collections.Counter()
    for (d, job), lb in attr_by_key.items():
        gap = lb - cut_total.get((d, job), [0.0, 0.0])[1]
        if gap > 2.0: over[d] += gap
    if over:
        say("\n*** STATION FEED LOOKS WRONG ON THESE DATES ***")
        for d, lb in sorted(over.items()):
            stt = sum(v[1] for k, v in cut_total.items() if k[0] == d) / 2000.0
            sat = sum(v for k, v in attr_by_key.items() if k[0] == d) / 2000.0
            say(f"    {d}  station {stt:,.1f} t but saws attribute {sat:,.1f} t"
                f"  (ratio {sat/stt if stt else 0:.2f})")
        say("    A ratio near 2.20 means station_history holds unconverted KILOGRAMS")
        say("    for that date. Re-pull just those days:")
        say("      py backfill_history.py --out DATA --start <date> --end <date>")
        say("      py merge_history.py --station DATA\\daily_by_job_and_station.csv --station-hist DATA\\station_history.csv")
        say("    The feed uses the larger figure, so it is already right. History is not.\n")

    attr_pct = 100.0 * (1 - unattr_lb / (cut_tons_total * 2000.0)) if cut_tons_total else 0.0
    say(f"cut totals from station_history; saw attribution {attr_pct:.1f}% "
        f"({unattr_lb/2000.0:,.1f} t carried as {UNATTRIBUTED})")

    # ---------------- SANITY GATE ----------------
    ratio = cut_tons_total / fit_tons_total if fit_tons_total else 0.0
    say(f"\nsanity gate: cut {cut_tons_total:,.1f} t / fit {fit_tons_total:,.1f} t = {ratio:.3f}")
    if ratio < 0.90:
        print("\n*** SANITY GATE FAILED, NOTHING WAS PUBLISHED ***")
        print(f"    Cut is {ratio:.2f}x fit. Cutting precedes final inspection, so cut")
        print( "    can never be below fit. Two known causes:")
        print(f"      {ratio:.2f} near 0.45  -> the API returned KILOGRAMS and nobody converted")
        print(f"      {ratio:.2f} near 0.72  -> cut was read from cut_by_saw, which drops")
        print( "                          unmapped operators. Use station_history for totals.")
        print( "    Fix the cause. Do not relax this gate.\n")
        return 2

    # ---------------- REGRESSION GATE ----------------
    # The sanity gate above compares cut to fit, so it cannot see a TRUNCATION:
    # if history is destroyed both sides shrink together and the ratio still looks
    # fine. On 2026-09-08 backfill_history.py was run with a one-day range, which
    # is a full-file rewrite, and 20 months of history became one day. The ratio
    # read 1.176 and the feed published happily. This catches that.
    state_path = os.path.join(out, "feed_state.json")
    prev = None
    if os.path.exists(state_path):
        try: prev = json.load(open(state_path, encoding="utf-8"))
        except Exception: prev = None
    fit_dates = sorted({k[0] for k in fit})
    state = {"fit_rows": len(fit), "fit_tons": round(fit_tons_total, 1),
             "cut_tons": round(cut_tons_total, 1),
             "first_date": fit_dates[0] if fit_dates else "",
             "last_date": fit_dates[-1] if fit_dates else ""}
    if prev and not a.allow_shrink:
        why = []
        if prev.get("fit_rows", 0) and state["fit_rows"] < 0.60 * prev["fit_rows"]:
            why.append(f"rows {prev['fit_rows']:,} -> {state['fit_rows']:,}")
        if prev.get("fit_tons", 0) and state["fit_tons"] < 0.60 * prev["fit_tons"]:
            why.append(f"fit tons {prev['fit_tons']:,.0f} -> {state['fit_tons']:,.0f}")
        if prev.get("first_date") and state["first_date"] > prev["first_date"]:
            why.append(f"history now starts {state['first_date']}, was {prev['first_date']}")
        if why:
            print("\n*** FEED SHRANK, NOTHING WAS PUBLISHED ***")
            for w in why: print("      " + w)
            print("    History looks truncated. The usual cause is backfill_history.py run")
            print("    with a narrow --start/--end: it rewrites the WHOLE file, it does not")
            print("    patch a date range. Rebuild with code\\restore_history.bat.")
            print("    If you MEANT to shrink it, re-run with --allow-shrink.\n")
            return 2

    # ---------------- WRITE ----------------
    fit_path = os.path.join(out, "feed_fit.csv")
    with open(fit_path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Report Date","Name","Job #","Test Location","Plant","Total","Weight",
                    "Failures","Resolved","Outstanding","Failure Rate"])
        for (d, job, shown, plant), (pcs, lb) in sorted(fit.items()):
            w.writerow([d, "PowerFab API", job, shown, plant,
                        int(round(pcs)), round(lb, 2), 0, 0, 0, 0])

    cut_path = os.path.join(out, "feed_cut.csv")
    with open(cut_path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Date","Job #","Saw","Plant","Comp","Comp Wt"])
        for (d, job, saw, plant), (pcs, lb) in sorted(cut.items()):
            w.writerow([d, job, saw, plant, int(round(pcs)), round(lb, 2)])

    paint_path = os.path.join(out, "feed_paint.csv")
    with open(paint_path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Date","Job #","Comp","Comp Wt"])
        for (d, job), (pcs, lb) in sorted(paint.items()):
            w.writerow([d, job, int(round(pcs)), round(lb, 2)])

    say(f"\nfeed_fit.csv    {len(fit):6d} rows  {fit_tons_total:10,.1f} t   -> {fit_path}")
    say(f"feed_cut.csv    {len(cut):6d} rows  {cut_tons_total:10,.1f} t   -> {cut_path}")
    say(f"feed_paint.csv  {len(paint):6d} rows  {paint_tons_total:10,.1f} t   -> {paint_path}")

    if REJECTS:
        rej = os.path.join(out, "feed_rejects.csv")
        with open(rej, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f); w.writerow(["Feed","Job #","Tons"]); w.writerows(REJECTS)
        bad = collections.Counter()
        for where, j, t in REJECTS: bad[j] += t
        say("\nNON-NUMERIC JOB CODES held out of the feed (see feed_rejects.csv):")
        for j, t in bad.most_common(): say(f"   {j!r}  {t:.2f} t")

    by_plant = collections.Counter()
    for (d, job, shown, plant), (pcs, lb) in fit.items():
        by_plant[plant] += lb / 2000.0
    say("\nfit tons by plant, all history:")
    for p in ("P1", "P3", "P4"): say(f"   {p}  {by_plant[p]:10,.1f}")
    try:
        json.dump(state, open(state_path, "w", encoding="utf-8"), indent=2)
    except Exception as e:
        say("could not write feed_state.json:", e)

    say("\nP3 IS IN THE FEED ON PURPOSE. The board drops it at display, the workbook keeps it.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
