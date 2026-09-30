# MIW Production Board

The live Production Board for Merchant Iron Works, as it runs on the OPERATIONS server every 5 minutes.
This `main` branch is a copy of what is deployed on the Server site as of 2026-09-29 (Week 40), including
the 9/29 fixes: Sunday column, Plate row, unmapped-login logging, and the saw table job tooltip.

The department boards (QC first) are on the `department-boards` branch and are not live.

## Layout

| Folder | What it is | Where it lives on the Server site |
|---|---|---|
| `pipeline/` | PowerFab API pull: station tons, QC by plant, cut by saw, paint by plant, merged into history, then the one canonical feed (`feed_fit`, `feed_cut`, `feed_paint`) | `Server\PowerFab API Pull\code\` |
| `board/` | Builds the board from the feeds plus QuickBooks Time hours: `explorer.html` (Weekly/Monthly with history) and `index_fab.html` (Fabrication Board) | `Server\Production Board\` |
| `command-center/` | The MIW Command Center launcher and its tile list | `SCHEDULES\MIW Command Center\App\` |

## How it runs

1. `pipeline/run_powerfab_pull.bat` (Task Scheduler, from 05:00 every 5 minutes) runs `pull_daily.py`,
   `pull_qc_by_plant.py`, `pull_cut_by_saw.py`, `pull_paint_by_plant.py`, `merge_history.py`, `build_feed.py`,
   `fetch_quickbooks.py` and `join_hours.py`. The feeds are written to `SCHEDULES\Daily Reports\PowerFab API\`.
2. `board/run_board.bat` runs `build_history_board.py` (which imports `build_production_json.py`),
   `build_fab_json.py` and `publish_backlog_link.py`, then copies `explorer.html` and `index_fab.html` to
   `SCHEDULES\MIW Command Center\App\`.
3. The Command Center opens those files in Edge app mode. Each page carries its own data and reloads
   itself every 5 minutes.

## Rules that are not obvious from the code

- `ProductionShippingProductivity_Get` says Imperial and sends kilograms. Every pull converts kg to lb.
- Weight in the feeds is already extended. Never multiply weight by quantity. Tons = lb / 2000.
- Fit = QC-Final passed (`VT - Final Visual Weld Inspection`). Plant comes only from the inspection Test Location.
- Cut is split to saws by the `Station Completed By` login (`OP2SAW` in `pull_cut_by_saw.py`).
- Plant tiles use fabrication hours only; the Overall MH/ton box uses every P1/P4 department. Plant 3 is collected, not shown.
- Week columns run Monday to Sunday; pace runs Monday to Friday.

## Not in this repo

- `powerfab.ini` (copy `pipeline/python/powerfab.ini.template` and fill in on the server) and the QuickBooks
  Time token (`C:\MIW\qbt_token.txt`). Never commit either.
- `Tekla.PowerFab.API.dll` and the built `MIWFab.exe`. Build with `dotnet build -c Release` in `pipeline/MIWFab`.
- All data: history CSVs, feeds, generated HTML and JSON.
