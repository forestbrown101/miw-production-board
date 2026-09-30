@echo off
REM ============================================================================
REM  Production Board build (engine lives here in the Server site). Reads the
REM  PowerFab API CSVs + QuickBooks + newest backlog, builds here, then ships ONLY
REM  index.html, index_month.html, explorer.html to the shared SCHEDULES viewer.
REM ============================================================================
setlocal
set "BK=%~dp0"
set "ONE=%OneDriveCommercial%"
if not defined ONE set "ONE=%OneDrive%"
if not exist "%ONE%\Merchant Iron Works Team Site - SCHEDULES" ( echo ERROR: OneDrive root not auto-detected. Set ONE by hand. & exit /b 2 )
set "SITE=%ONE%\Merchant Iron Works Team Site - SCHEDULES"
set "SRV=%ONE%\Merchant Iron Works Team Site - Server"
REM Daily Reports lives in the SCHEDULES site (moved back 2026-09-04). The board reads it
REM for HOURS only (QuickBooks Time) plus the fab board file; cut and fit come from the API.
REM
REM PINNED TO SCHEDULES 2026-09-08. There used to be a fallback to the Server copy, kept
REM while OneDrive finished syncing. That sync is long done, and the Server copy went stale
REM (newest timesheet 2026-09-04) while looking perfectly healthy. A silent fallback to a
REM stale hours folder does not fail, it just quietly understates every MH/ton on the TV.
REM One source now, and a loud stop if it is not there.
set "DR=%SITE%\Daily Reports"
set "QBDIR=%DR%\QuickBooks Time"
if not exist "%QBDIR%" (
  echo ERROR: QuickBooks Time not found at:
  echo    %QBDIR%
  echo Hours come from the SCHEDULES site only. Refusing to build a board with zero
  echo hours; the previous board stays up until this is fixed.
  exit /b 2
)
REM A folder that EXISTS but has stopped receiving files is the quieter version of the same
REM problem, so say so. Warning only: a Monday morning legitimately trails Friday.
REM  Age of the newest timesheet, read from the FILE NAME (auto_timesheet_YYYY-MM-DD.csv)
REM  rather than its mtime. A PowerShell one-liner was used here first and returned the
REM  999 sentinel on every run: quoting a path with spaces inside a for /f inside a .bat
REM  is three levels of escaping and it lost. Python is already resolved as %PY%, the
REM  filenames carry the date, so this needs no escaping and cannot half-work.
set "QBAGE="
for /f "delims=" %%a in ('"%PY%" -c "import glob,os,datetime;f=sorted(glob.glob(os.path.join(r'%QBDIR%','auto_timesheet_*.csv')));print((datetime.date.today()-datetime.date.fromisoformat(os.path.basename(f[-1])[15:25])).days if f else 999)" 2^>nul') do set "QBAGE=%%a"
if not defined QBAGE set "QBAGE=0"
if %QBAGE% GEQ 3 echo WARNING: newest QuickBooks timesheet is %QBAGE% days old in "%QBDIR%" - hours are understated
REM The engine + all data live and run here on the Server. We only PUBLISH the finished
REM board pages into the launcher folder on SCHEDULES: coworker PCs sync SCHEDULES but not
REM the Server site, so that is the only copy their launcher can open.
set "PUB=%SITE%\MIW Command Center\App"
REM The fab board is now BUILT here every cycle (see below), not copied from a folder.
REM It used to be copied from Schedule Board, which stopped being rebuilt when the engine
REM moved to the Server site, so the launcher showed a frozen 2026-09-03 snapshot. The
REM page embeds its data as a fallback, so a stale file looks alive and is the worst kind.
set "FAB=%BK%index_fab.html"
set "DATA=%SRV%\PowerFab API Pull\data"
set "BLDIR=%SITE%\2026\PLANT 1, 3 & 4"
set "QC_CSV=%DATA%\qc_history.csv"
set "STATION_CSV=%DATA%\station_history.csv"
set "CUT_CSV=%DATA%\cut_by_saw_history.csv"
REM Paint split by plant (Plant 2 paint shop opened 2026-09-15). Dates missing from this
REM file fall back to the station_history Paint total as Plant 1, so old weeks still show.
set "PAINT_CSV=%DATA%\paint_by_plant_history.csv"
REM THE CANONICAL FEED. The same two files the WEEK ## SCHEDULES workbook reads, so the
REM TV and the spreadsheet cannot drift. build_production_json.py prefers these over
REM QC_CSV and over the Cut Daily .xls folders, and falls back automatically if missing.
set "FEED=%SITE%\Daily Reports\PowerFab API"
set "FEED_FIT=%FEED%\feed_fit.csv"
set "FEED_CUT=%FEED%\feed_cut.csv"
set "OUT=%BK%"
REM Use the REAL python.exe directly; the py launcher can intermittently fail "cannot execute".
set "PY="
for %%p in (
  "%LocalAppData%\Python\pythoncore-3.14-64\python.exe"
  "%LocalAppData%\Programs\Python\Python314\python.exe"
  "%LocalAppData%\Programs\Python\Python313\python.exe"
) do if not defined PY if exist "%%~p" set "PY=%%~p"
if not defined PY for /f "delims=" %%p in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%p"
if not defined PY ( where py >nul 2>&1 && set "PY=py -3" )
if not defined PY set "PY=python"
if not exist "%PUB%" mkdir "%PUB%" 2>nul
set "NAMES="
for /f "delims=" %%f in ('dir /b /a-d /o-d "%BLDIR%\*WEEK*BACKLOG*.xlsx" 2^>nul') do if not defined NAMES set "NAMES=%BLDIR%\%%f"
REM explorer-only: build_history_board imports build_production_json for its aggregate,
REM so that .py stays, we just no longer generate the two index files.
%PY% "%BK%build_history_board.py"
REM canonical explorer.html stays here on the Server (%BK%); publish a copy for the launcher.
REM ATOMIC publish: copy to a temp name, then rename into place. A rename is instantaneous,
REM so an open board that reloads mid-publish gets either the old complete file or the new
REM complete file, never a half-written one (which used to kill the page's refresh timer).
copy /y "%OUT%explorer.html" "%PUB%\explorer.html.tmp" >nul
move /y "%PUB%\explorer.html.tmp" "%PUB%\explorer.html" >nul
REM --- Fabrication board: same feed, same 5 minute cycle as the production board ---
REM  SOURCE is %DR% for QuickBooks hours only; fit comes from FEED_FIT. Standalone script
REM  on purpose, so a fault here can never take the production board down with it.
REM  "%BK%." not "%BK%": BK ends in a backslash, and "C:\path\" lets the backslash escape
REM  the closing quote, so the OUT argument arrives mangled. The trailing dot is the same
REM  directory and parses cleanly.
"%PY%" "%BK%build_fab_json.py" "%DR%" "%BK%."
if exist "%FAB%" (
  copy /y "%FAB%" "%PUB%\index_fab.html.tmp" >nul
  move /y "%PUB%\index_fab.html.tmp" "%PUB%\index_fab.html" >nul
  REM publish the JSON beside it so the page's own refresh can re-fetch without a rebuild
  if exist "%BK%production_data_fab.json" (
    copy /y "%BK%production_data_fab.json" "%PUB%\production_data_fab.json.tmp" >nul
    move /y "%PUB%\production_data_fab.json.tmp" "%PUB%\production_data_fab.json" >nul
  )
) else (
  echo WARNING: build_fab_json.py produced no index_fab.html - fab board not republished
)
echo Explorer built (%BK%) and published -^> %PUB%

REM --- Backlog link: pick the newest weekly schedules workbook in SCHEDULES\2026 and
REM  publish its direct SharePoint link into the launcher folder (backlog_latest.json).
REM  The launcher opens that link in desktop Excel, so coworkers no longer need the 2026
REM  folder synced on their own PC. See publish_backlog_link.py.
set "BL2026=%SITE%\2026"
set "SP_BASE=https://merchantironworks.sharepoint.com/SCHEDULES/2026/"
set "SP_FOLDER_URL=https://merchantironworks.sharepoint.com/SCHEDULES/Forms/AllItems.aspx?id=%%2FSCHEDULES%%2F2026"
%PY% "%BK%publish_backlog_link.py"
