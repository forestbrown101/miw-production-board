# Department boards (in progress, not live)

A further breakdown of the Production Board, one page per department. The same kind of data as the main
board, split further (by bay, by job, by day). No forecasting, no plans. Each page reuses the Production
Board's own head, CSS and refresh engine, so it looks and behaves like the main board.

## QC board (`qc/`)

- Totals: tons fit, pieces, Plant 1 tons.
- East Bay (Texas A&M maroon), West Bay (lavender) and Plant 4 side by side: tons, pieces, share of total,
  busiest job, jobs through QC.
- Jobs: QC tons per job, each bar split East / West / P4.
- Daily table: East Bay, West Bay, Plant 1 total, Plant 4, Total, Monday to Sunday.
- Job progress: QC-Final tons to date against released (PowerFab BOM) tons, and tons left.
- "Production Board" back button; Weekly/Monthly and the period carry both ways in the URL hash
  (`v=week|month&p=<period key>`, same format as `explorer.html`).

Build:

```
python build_qc_board.py <feed_fit.csv> <feed_released.csv> <control_module.json> qc_template.html qc.html
```

- `feed_fit.csv` and `feed_released.csv` come from `SCHEDULES\Daily Reports\PowerFab API\`.
- `control_module.json` is used for job names only and is optional.
- Test Location mapping: C-Beams and SHOP go to West Bay, Paint Bay goes to East Bay, and Plant 3 is not shown.

Checked against the live board: the week of 9/21 is 187.3 East + 203.4 West + 37.4 P4 = 428.1 t, the same
as the Fabrication Board.

`qc_template.html` is `board/explorer_template.html` lines 1 to 153 (head, refresh engine, CSS) followed by
the QC page body. If the main template's CSS changes, rebuild it the same way.

Still to come: Cut and Paint in the same style, the Production Board links into these pages, and
Command Center tiles.
