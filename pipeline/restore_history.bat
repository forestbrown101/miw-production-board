@echo off
setlocal
REM ============================================================================
REM  restore_history.bat    RUN THIS ON THE SERVER. Takes a while.
REM
REM  Rebuilds qc_history.csv and station_history.csv from scratch, month by
REM  month, straight from the PowerFab API with the corrected kg->lb conversion.
REM
REM  WHY THIS EXISTS
REM  backfill_history.py writes those two files WHOLESALE. It is a full-history
REM  tool, not a date patcher. On 2026-09-08 it was run with a one-day range and
REM  it replaced 20 months of history with that single day: qc_history went from
REM  73,786 rows to 3,054 and station_history from 7,084 to 432. There is no
REM  partial mode. If a single day in history is wrong, the fix is to rebuild
REM  the whole file, which is cheap enough and always correct.
REM
REM  cut_by_saw_history.csv is NOT touched by this. It is built by
REM  pull_cut_by_saw.py and was unaffected.
REM
REM  Safe to run any time. It only writes at the end of each month's pull, and
REM  a failure leaves the previous file in place.
REM ============================================================================
set "START=%~1"
if not defined START set "START=2025-01-01"

set "ONE=%OneDriveCommercial%"
if not defined ONE set "ONE=%OneDrive%"
if not exist "%ONE%\Merchant Iron Works Team Site - Server" (
  echo ERROR: OneDrive root not auto-detected. Expected C:\Shares\OneDrive - Merchant Iron Works
  pause & exit /b 2
)
set "OUT=%ONE%\Merchant Iron Works Team Site - Server\PowerFab API Pull\data"
set "PYDIR=%ONE%\Merchant Iron Works Team Site - Server\PowerFab API Pull\code\python"

set "PY="
for %%p in (
  "%LocalAppData%\Python\pythoncore-3.14-64\python.exe"
  "%LocalAppData%\Programs\Python\Python314\python.exe"
  "%LocalAppData%\Programs\Python\Python313\python.exe"
  "%LocalAppData%\Programs\Python\Python312\python.exe"
) do if not defined PY if exist "%%~p" set "PY=%%~p"
if not defined PY for /f "delims=" %%p in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%p"
if not defined PY ( where python >nul 2>&1 && set "PY=python" )
if not defined PY ( echo ERROR: no python found. & pause & exit /b 2 )

echo python : %PY%
echo from   : %START%  through today
echo data   : %OUT%
echo.
echo This rebuilds ALL of qc_history.csv and station_history.csv. Several minutes.
echo.

pushd "%PYDIR%"

echo === 1. full history rebuild ===
"%PY%" backfill_history.py --out "%OUT%" --start %START%
if %ERRORLEVEL% GEQ 2 ( echo BACKFILL FAILED. & popd & pause & exit /b 2 )

echo.
echo === 2. folding the current trailing window back in ===
"%PY%" merge_history.py --qc "%OUT%\qc_by_plant.csv" --qc-hist "%OUT%\qc_history.csv" ^
     --station "%OUT%\daily_by_job_and_station.csv" --station-hist "%OUT%\station_history.csv" ^
     --cut "%OUT%\cut_by_saw.csv" --cut-hist "%OUT%\cut_by_saw_history.csv"

echo.
echo === 3. rebuild and publish the feed, with the audit ===
set "FEED=%ONE%\Merchant Iron Works Team Site - SCHEDULES\Daily Reports\PowerFab API"
if not exist "%FEED%" mkdir "%FEED%" 2>nul
"%PY%" build_feed.py --data "%OUT%" --out "%FEED%" --allow-shrink
set BRC=%ERRORLEVEL%

popd
echo.
echo Expect roughly: fit 31,400 t, cut 30,900 t, gate about 0.98,
echo qc_history near 73,800 rows and station_history near 7,100 rows.
echo If the numbers are far below that, STOP and say so before anything reads the feed.
echo.
pause
endlocal & exit /b %BRC%
