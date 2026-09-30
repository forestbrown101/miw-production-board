# MIW Production Board — Code Overview

## What it does
Every 5 minutes, a scheduled job reads the raw PowerFab + QuickBooks Time drop files,
computes tons / hours / MH-ton by plant, and rebuilds two web dashboards (weekly + monthly)
that display on the shop TVs.

## Data flow
PowerFab  -> Daily Reports\Cut Daily\*.xls        (cut tons, every 5 min)
PowerFab  -> Daily Reports\QC Final Report\*.xls   (fit tons, every 5 min)
QuickBooks Time API -> Daily Reports\QuickBooks Time\auto_timesheet_*.csv  (hours, every 5 min)
        |
        v
run_board.bat  (Task Scheduler, every 5 min)
    1. fetch_quickbooks.py   -> writes today's hours CSV
    2. build_production_json.py -> reads all three folders, writes:
           production_data.json      (weekly)
           production_data_month.json (monthly)
           index.html / index_month.html (self-contained boards)
        |
        v
serve_board.bat serves the folder on :8080  ->  Skykit TVs / browsers

## Files
- build_production_json.py  Engine. Reads Cut Daily / QC Final Report / QuickBooks Time,
                            maps locations to plants, computes metrics, writes JSON + HTML.
- fetch_quickbooks.py       Calls the QuickBooks Time REST API, writes the hours CSV.
- board_template.html       Board layout/design (data injected each build).
- run_board.bat             Orchestrator (fetch -> build). Self-locating paths.
- serve_board.bat           Local web server on port 8080.
- MIW_Board_Task.xml        Windows Task Scheduler job (every 5 min).
- index.html / index_month.html   Generated boards (weekly / monthly).
- production_data*.json      Generated data the boards read.

## Key settings (top of build_production_json.py)
- WEEK_GOAL / CUT_WEEK_GOAL   weekly tons goals (monthly auto-scales)
- FIT_LOC_TO_PLANT            QC test-location -> plant (incl. Paint Bay -> P1)
- tile_plant() / overall_plant()  hours mapping (per-area tiles vs all-department overall)

## Secrets
- C:\MIW\qbt_token.txt  QuickBooks Time API access token (local only, not in OneDrive).

## Cutover notes
- Hours are now auto-pulled; the manual morning export is retired.
- Only one source per date in QuickBooks Time (auto files own current dates; older manual
  files kept for earlier history) to avoid double-counting.
