@echo off
setlocal enabledelayedexpansion
REM ============================================================================
REM  run_powerfab_pull.bat
REM  One pass of the whole pipeline: PowerFab API -> CSV -> joined with QuickBooks.
REM
REM  Safe to run on a schedule. Every run overwrites a trailing window rather than
REM  appending, so a missed run or a double run cannot corrupt anything.
REM
REM  IMPORTANT: do NOT run this while OneDrive is actively syncing the code folder.
REM  cmd re-reads a .bat from disk as it runs, so if OneDrive swaps the file mid-run
REM  the variables come back empty and paths break. Run only when OneDrive on this
REM  machine shows "Up to date", or run from a LOCAL copy (see the staging note).
REM
REM  Exit codes:  0 clean   1 XMLWarning (result suspect)   2 hard failure
REM ============================================================================

REM --- resolve the OneDrive root for THIS machine (work box or server) ------
set "ONE=%OneDriveCommercial%"
if not defined ONE set "ONE=%OneDrive%"
if not exist "%ONE%\Merchant Iron Works Team Site - SCHEDULES" (
  echo ERROR: OneDrive root not auto-detected. Set it by hand:  set "ONE=C:\Shares\OneDrive - Merchant Iron Works"
  exit /b 2
)

set "OUT=%ONE%\Merchant Iron Works Team Site - Server\PowerFab API Pull\data"
REM Daily Reports (incl. QuickBooks Time) lives in the SCHEDULES site (moved back 2026-09-04).
REM
REM PINNED TO SCHEDULES 2026-09-08. The Server fallback used to be here "while OneDrive
REM finishes syncing". That sync is long done, and the fallback is exactly how the Server
REM copy became a stale second source: this script WRITES timesheets, so one run with
REM SCHEDULES briefly unavailable silently starts a parallel hours folder that then looks
REM perfectly healthy and quietly understates every MH per ton downstream. One source.
REM
REM Missing is a WARNING here, not a fatal error, because tons do not need QuickBooks: the
REM API pull, the history merge and the feed build all still run and stay correct. Only the
REM two hours steps are skipped. (run_board.bat DOES hard-stop on the same folder, because
REM there a missing hours feed would put a wrong MH per ton on a plant-floor TV.)
set "QB=%ONE%\Merchant Iron Works Team Site - SCHEDULES\Daily Reports\QuickBooks Time"
set "QBOK=1"
if not exist "%QB%" (
  set "QBOK="
  call :say "WARNING: QuickBooks Time not found at %QB%"
  call :say "         hours steps SKIPPED this run; tons and the feed are unaffected."
  call :say "         Do NOT point this at the Server copy - fix the SCHEDULES sync."
)

set "SYNC=%ONE%\Merchant Iron Works Team Site - Server\PowerFab API Pull\code"

REM --- create the output folder BEFORE anything tries to log into it ----------
if not exist "%OUT%" mkdir "%OUT%" 2>nul
if not exist "%OUT%" (
  echo ERROR: cannot create output folder:
  echo    %OUT%
  exit /b 2
)
set "LOG=%OUT%\pull_log.txt"

call :say "---- pull start %date% %time% ----"

REM --- find Python -----------------------------------------------------------
REM  Use the REAL python.exe directly. The py launcher / Windows Store python alias can
REM  intermittently fail "The system cannot execute the specified program" mid-run (seen
REM  2026-09-04: the QuickBooks fetch and a date call errored while other calls in the same
REM  run worked). A plain .exe path never has that problem.
set "PY="
for %%p in (
  "%LocalAppData%\Python\pythoncore-3.14-64\python.exe"
  "%LocalAppData%\Programs\Python\Python314\python.exe"
  "%LocalAppData%\Programs\Python\Python313\python.exe"
  "%LocalAppData%\Programs\Python\Python312\python.exe"
) do if not defined PY if exist "%%~p" set "PY=%%~p"
if not defined PY for /f "delims=" %%p in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%p"
if not defined PY for /f "delims=" %%p in ('python -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%p"
if not defined PY ( where py >nul 2>&1 && set "PY=py -3" )
if not defined PY ( where python >nul 2>&1 && set "PY=python" )
if not defined PY (
  call :say "ERROR: no python found on PATH."
  exit /b 2
)
call :say "python: %PY%"

if not exist "%SYNC%\python\pull_daily.py" (
  call :say "ERROR: %SYNC%\python\pull_daily.py not found."
  exit /b 2
)

REM  find MIWFab.exe anywhere under the code folder (build output carries its DLLs)
set "MFX="
for /f "delims=" %%e in ('dir /b /s "%SYNC%\MIWFab.exe" 2^>nul') do if not defined MFX set "MFX=%%e"
if not defined MFX (
  call :say "ERROR: MIWFab.exe not found anywhere under %SYNC%"
  exit /b 2
)

REM --- single-instance guard -------------------------------------------------
REM  The 5-minute scheduled task and a hand run must never overlap: they would share
REM  scratch files and both write the same output CSVs, which scrambles a run. If a
REM  pull is already going, bow out.
set "LOCK=%OUT%\pull.lock"
if exist "%LOCK%" (
  call :say "another pull is already running (lock present) - skipping this run"
  call :say "  if you are certain none is running, delete this file and re-run:"
  call :say "  %LOCK%"
  exit /b 0
)
echo %date% %time% > "%LOCK%"

REM  Scripts are run by FULL PATH so a working directory lost to OneDrive churn cannot
REM  make Python fail to find them. Scratch files are unique per run (%T%).
pushd "%SYNC%\python" 2>nul
set "PYDIR=%SYNC%\python"
set "T=%TEMP%\pf_%RANDOM%%RANDOM%"

REM --- 1. pull from PowerFab -------------------------------------------------
call :say "pulling from PowerFab..."
%PY% "%PYDIR%\pull_daily.py" --out "%OUT%" > "%T%_pull.txt" 2>&1
set RC=%ERRORLEVEL%
type "%T%_pull.txt"
type "%T%_pull.txt" >> "%LOG%"
if %RC% GEQ 2 (
  call :say "PULL FAILED rc=%RC%. Previous CSVs left untouched."
  del "%T%_*.txt" 2>nul
  del "%LOCK%" 2>nul
  popd
  exit /b 2
)

REM --- 1b. QC by plant -------------------------------------------------------
call :say "pulling QC by plant..."
%PY% "%PYDIR%\pull_qc_by_plant.py" --out "%OUT%" > "%T%_qc.txt" 2>&1
set QRC=%ERRORLEVEL%
type "%T%_qc.txt"
type "%T%_qc.txt" >> "%LOG%"
if %QRC% GEQ 2 call :say "QC pull failed rc=%QRC%"

REM --- 1c. cut by saw (-> plant), from the API. Weights come back in kg with
REM  Station Completed By in the GroupBy; pull_cut_by_saw.py converts. Saw 1-4 = P1, Saw 5 = P4.
call :say "pulling cut by saw..."
%PY% "%PYDIR%\pull_cut_by_saw.py" --out "%OUT%" --days 14 > "%T%_cut.txt" 2>&1
set CRC=%ERRORLEVEL%
type "%T%_cut.txt"
type "%T%_cut.txt" >> "%LOG%"
if %CRC% GEQ 2 call :say "cut pull failed rc=%CRC%"

REM --- 1c2. paint by plant, from the API. Same call as cut; Station Completed By is the
REM  only link to a plant. Jvaughn = Plant 2 (paint shop opened 2026-09-15), everyone else = P1.
call :say "pulling paint by plant..."
%PY% "%PYDIR%\pull_paint_by_plant.py" --out "%OUT%" --days 14 > "%T%_paint.txt" 2>&1
set PRC=%ERRORLEVEL%
type "%T%_paint.txt"
type "%T%_paint.txt" >> "%LOG%"
if %PRC% GEQ 2 call :say "paint pull failed rc=%PRC%"

REM --- purchase orders are NOT here on purpose ------------------------------
REM  A full PO pull is ~6 minutes (20,000 POs, 55,000 lines, measured 2026-09-09),
REM  longer than this task's 5 minute cycle; inside it, it cost the boards a refresh
REM  every hour. It runs as its own hourly task: run_po_pull.bat /quiet, scheduled by
REM  MIW_PO_Task_Hourly.xml, own lock and log (data\po_pull_log.txt).

REM --- 1d. merge trailing pulls into permanent history ------------------------
call :say "merging into history..."
%PY% "%PYDIR%\merge_history.py" --qc "%OUT%\qc_by_plant.csv" --qc-hist "%OUT%\qc_history.csv" ^
     --station "%OUT%\daily_by_job_and_station.csv" --station-hist "%OUT%\station_history.csv" ^
     --cut "%OUT%\cut_by_saw.csv" --cut-hist "%OUT%\cut_by_saw_history.csv" ^
     --paint "%OUT%\paint_by_plant.csv" --paint-hist "%OUT%\paint_by_plant_history.csv" > "%T%_merge.txt" 2>&1
type "%T%_merge.txt"
type "%T%_merge.txt" >> "%LOG%"

REM --- 1d2. build and publish the ONE canonical feed -------------------------
REM  Every plant mapping rule lives in build_feed.py, once. The Production Board
REM  and the WEEK ## SCHEDULES workbook both read what it writes, so they cannot
REM  drift apart. rc=2 means it REFUSED to publish and both consumers keep the
REM  last good copy, which is the correct behaviour. Read the message, do not
REM  work around it.
set "FEED=%ONE%\Merchant Iron Works Team Site - SCHEDULES\Daily Reports\PowerFab API"
if not exist "%FEED%" mkdir "%FEED%" 2>nul
call :say "building canonical feed..."
%PY% "%PYDIR%\build_feed.py" --data "%OUT%" --out "%FEED%" > "%T%_feed.txt" 2>&1
set BRC=!ERRORLEVEL!
type "%T%_feed.txt"
type "%T%_feed.txt" >> "%LOG%"
if !BRC! GEQ 2 call :say "FEED NOT PUBLISHED rc=!BRC! - board and workbook keep the last good copy"

REM --- 1e. QuickBooks Time hours (TSheets API) -> auto_timesheet CSVs ---------
REM  Used to ride the retired Schedule Board task; pulled here now so hours stand
REM  alone. Writes yesterday+today. Needs the token at C:\MIW\qbt_token.txt (server).
set "FQB=%ONE%\Merchant Iron Works Team Site - Server\Production Board\fetch_quickbooks.py"
if not defined QBOK (
  call :say "hours fetch skipped, no QuickBooks Time folder"
) else if exist "%FQB%" (
  call :say "pulling QuickBooks Time hours..."
  %PY% "%FQB%" --qbdir "%QB%" > "%T%_qbt.txt" 2>&1
  set FRC=!ERRORLEVEL!
  type "%T%_qbt.txt"
  type "%T%_qbt.txt" >> "%LOG%"
  if !FRC! GEQ 1 call :say "QuickBooks fetch rc=!FRC! (needs the token at C:\MIW\qbt_token.txt on this machine)"
) else (
  call :say "fetch_quickbooks.py not found at %FQB% - hours feed skipped"
)

REM --- 2. join to QuickBooks hours -------------------------------------------
for /f %%d in ('%PY% -c "import datetime;print((datetime.date.today()-datetime.timedelta(days=14)).isoformat())"') do set "START=%%d"
for /f %%d in ('%PY% -c "import datetime;print(datetime.date.today().isoformat())"') do set "END=%%d"
if not defined QBOK (
  call :say "hours join skipped, no QuickBooks Time folder - tons_and_hours.csv left as-is"
) else (
  call :say "joining hours %START% to %END% ..."
  %PY% "%PYDIR%\join_hours.py" --pf "%OUT%\daily_by_job_and_station.csv" --qb "%QB%" ^
      --start !START! --end !END! --out "%OUT%\tons_and_hours.csv" > "%T%_join.txt" 2>&1
  set JRC=!ERRORLEVEL!
  type "%T%_join.txt"
  type "%T%_join.txt" >> "%LOG%"
  if !JRC! GEQ 1 call :say "join_hours rc=!JRC!"
)

del "%T%_*.txt" 2>nul
popd

REM --- 2b. rebuild the API Production Board (explorer) -------------------------
set "BOARD=%ONE%\Merchant Iron Works Team Site - Server\Production Board\run_board.bat"
if exist "%BOARD%" ( call "%BOARD%" >> "%LOG%" 2>&1 ) else ( call :say "board build skipped, run_board.bat not found" )

REM --- 2b2. delivery schedules from Google Drive, once a day -------------------
REM  Needs Google Drive for desktop on this machine. The first cycle of each day
REM  runs the pull; the stamp file records the day it last succeeded. If Drive is
REM  not here the pull refuses (old file stays) and it is tried again next cycle.
set "DS=%ONE%\Merchant Iron Works Team Site - Server\Control Module\tools\pull_delivery_schedules.bat"
set "DSTAMP=%ONE%\Merchant Iron Works Team Site - Server\Daily Reports\Material Flow\delivery\pulled_on_server.txt"
set "DSDAY="
if exist "%DSTAMP%" set /p DSDAY=<"%DSTAMP%"
if exist "%DS%" if not "!DSDAY!"=="%date%" (
  call :say "delivery schedules from Google Drive (daily)..."
  set "MIW_PY=%PY%"
  call "%DS%" >> "%LOG%" 2>&1
  if "!ERRORLEVEL!"=="0" (
    >"%DSTAMP%" echo %date%
    call :say "delivery schedules pulled for %date%"
  ) else (
    call :say "delivery schedules NOT pulled rc=!ERRORLEVEL! (no Google Drive here yet? see Control Module\tools\pull_delivery_schedules_log.txt)"
  )
  set "MIW_PY="
)

REM --- 2c. MIW Control Module: build the numbers, audit them, build the app ----
REM  Reads the feed just published above. Guarded: its own log, its own lock, its
REM  own exit code, and nothing here can fail the pull. What happened is in
REM  Control Module\out\LAST RUN.txt. It is handed this run's python so the two
REM  never disagree about which python.exe to use.
set "CM=%ONE%\Merchant Iron Works Team Site - Server\Control Module\run_control_module.bat"
if exist "%CM%" (
  call :say "control module..."
  set "MIW_PY=%PY%"
  call "%CM%" >> "%LOG%" 2>&1
  call :say "control module rc=!ERRORLEVEL! (details in Control Module\out\LAST RUN.txt)"
  set "MIW_PY="
) else ( call :say "control module skipped, run_control_module.bat not found" )

REM --- 2d. run-once hook --------------------------------------------------------
REM  Drop a file named run_once.bat in the data folder and the next cycle runs it
REM  exactly once (it is renamed before it runs, so it can never run twice), with
REM  its output in the log. For one-off pulls without touching a scheduled task.
set "ONCE=%OUT%\run_once.bat"
if exist "%ONCE%" (
  ren "%ONCE%" "run_once.running.bat" >nul 2>&1
  if exist "%OUT%\run_once.running.bat" (
    call :say "run_once.bat found, running it once..."
    call "%OUT%\run_once.running.bat" >> "%LOG%" 2>&1
    call :say "run_once rc=!ERRORLEVEL!"
    move /y "%OUT%\run_once.running.bat" "%OUT%\run_once.done_%RANDOM%.bat" >nul 2>&1
  )
)

REM --- 3. trim the log --------------------------------------------------------
for %%A in ("%LOG%") do if %%~zA GTR 2000000 (
  powershell -NoProfile -Command "Get-Content -Tail 500 '%LOG%' | Set-Content '%LOG%.tmp'; Move-Item -Force '%LOG%.tmp' '%LOG%'"
)

del "%LOCK%" 2>nul
call :say "---- pull done rc=%RC% ----"
exit /b %RC%

:say
echo %~1
echo [%date% %time%] %~1>> "%LOG%"
goto :eof
