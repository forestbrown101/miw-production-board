PRODUCTION BOARD BACKEND
========================
Builds three TV boards from one 5 minute run: the Live board, the Monthly board
and the Fabrication board.

  run_board.bat            the runner. This is what Task Scheduler calls.
  fetch_quickbooks.py      pulls yesterday and today from QuickBooks Time into
                           Daily Reports\QuickBooks Time
  build_production_json.py Live + Monthly. Reads the Daily Reports drop folder.
  build_fab_json.py        Fabrication board. Standalone, so a failure here
                           cannot take the Live board down with it.
  serve_board.bat          serves the boards on port 8080 for Skykit and the TVs
  MIW_Board_Task.xml       import this into Task Scheduler to recreate the task


READS
-----
  SCHEDULES\Daily Reports\QC Final Report\     fit tons
  SCHEDULES\Daily Reports\Cut Daily\           cut tons, one subfolder per saw
  SCHEDULES\Daily Reports\QuickBooks Time\     man hours
  SCHEDULES\2026\PLANT 1, 3 & 4\24 WEEK_PLANT 1, 3 & 4 BACKLOG.xlsx   job names

WRITES  (all into SCHEDULES\Daily Reports\Schedule Board)
------
  production_data.json, production_data_month.json, production_data_fab.json
  index.html, index_month.html, index_fab.html
  board_log.txt          every run appends here. First place to look.

The output folder does not move. The TVs, the launcher and the templates
(board_template.html, fab_template.html) all live there, and the builders read
those templates out of the output folder to stamp the HTML.


WHAT CHANGED WHEN THIS MOVED TO THE SERVER FOLDER
-------------------------------------------------
run_board.bat used to work out every path from where it sat, two folders under
SCHEDULES. It cannot do that from here, so it now finds OneDrive from the
environment and builds the paths explicitly. Same result, no assumption about
where the code lives.

fetch_quickbooks.py used to assume the QuickBooks Time folder was its own
parent. It now takes --qbdir, which run_board.bat passes. With no --qbdir it
falls back to the old behaviour, so the old copy still works where it sits.

TO FINISH THE MOVE: point the "MIW Production Board" scheduled task at
  ...\Merchant Iron Works Team Site - Server\Production Board\run_board.bat
or delete the task and import MIW_Board_Task.xml, which already has it.
