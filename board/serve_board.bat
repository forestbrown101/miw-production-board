@echo off
REM ============================================================================
REM  serve_board.bat   serves the Schedule Board folder on port 8080.
REM  Run this at server startup so the TVs and Skykit can reach the boards by URL
REM  instead of by file path.
REM     http://<server-ip>:8080/index.html        Live board
REM     http://<server-ip>:8080/index_month.html  Monthly board
REM     http://<server-ip>:8080/index_fab.html    Fabrication board
REM  It serves the OUTPUT folder, which stays in the SCHEDULES team site.
REM ============================================================================
setlocal
set "ONE=%OneDriveCommercial%"
if not exist "%ONE%\Merchant Iron Works Team Site - SCHEDULES" set "ONE=%OneDrive%"
if not exist "%ONE%\Merchant Iron Works Team Site - SCHEDULES" set "ONE=%USERPROFILE%\OneDrive - Merchant Iron Works"
set "BOARD=%ONE%\Merchant Iron Works Team Site - SCHEDULES\Daily Reports\Schedule Board"
if not exist "%BOARD%" (
  echo ERROR: board folder not found: %BOARD%
  pause & exit /b 2
)
pushd "%BOARD%"
python -m http.server 8080
popd
endlocal
