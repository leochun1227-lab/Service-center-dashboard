@echo off
setlocal EnableExtensions

cd /d "%~dp0"
title Service Center Dashboard - Update Data

echo ============================================================
echo  Service Center Dashboard - Update Data
echo ============================================================
echo.
echo  This will fetch latest C4C tickets, enrich invoice data,
echo  refresh ticket status history, then rebuild dashboard-data.js from:
echo  c4c_ticket_table_z007_z010_checked_hana_final.xlsx
echo.

set "SOURCE_FILE=%CD%\c4c_ticket_table_z007_z010_checked_hana_final.xlsx"
set "C4C_EXPORT_FILE=%CD%\c4c_ticket_table_z007_z010_with_invoice_layout_checked.xlsx"
set "DASHBOARD_DATA_FILE=%CD%\dashboard-data.js"
set "FETCH_SCRIPT_FILE=%CD%\export_filtered_tickets_with_dealer_resolution.py"
set "HANA_SCRIPT_FILE=%CD%\sap_invoice_enrich_from_hana.py"
set "HISTORY_SCRIPT_FILE=%CD%\update_ticket_lastchangedtime.py"
set "SCRIPT_FILE=%CD%\generate_web_data_from_excel.py"
set "LOG_DIR=%CD%\outputs"
set "BACKUP_DIR=%LOG_DIR%\dashboard-data-backups"
set "SOURCE_BACKUP_DIR=%LOG_DIR%\source-workbook-backups"
set "LOG_FILE=%LOG_DIR%\update_dashboard_data.log"

if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"
if not exist "%BACKUP_DIR%" mkdir "%BACKUP_DIR%"
if not exist "%SOURCE_BACKUP_DIR%" mkdir "%SOURCE_BACKUP_DIR%"

if not exist "%FETCH_SCRIPT_FILE%" (
  echo [ERROR] C4C fetch script not found:
  echo %FETCH_SCRIPT_FILE%
  echo.
  pause
  exit /b 1
)

if not exist "%HANA_SCRIPT_FILE%" (
  echo [ERROR] SAP HANA enrich script not found:
  echo %HANA_SCRIPT_FILE%
  echo.
  pause
  exit /b 1
)

if not exist "%HISTORY_SCRIPT_FILE%" (
  echo [ERROR] History update script not found:
  echo %HISTORY_SCRIPT_FILE%
  echo.
  pause
  exit /b 1
)

if not exist "%SCRIPT_FILE%" (
  echo [ERROR] Update script not found:
  echo %SCRIPT_FILE%
  echo.
  pause
  exit /b 1
)

set "PYTHON_EXE="
set "BUNDLED_PY=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%BUNDLED_PY%" set "PYTHON_EXE=%BUNDLED_PY%"

if not defined PYTHON_EXE (
  where py >nul 2>nul
  if not errorlevel 1 set "PYTHON_EXE=py -3"
)

if not defined PYTHON_EXE (
  where python >nul 2>nul
  if not errorlevel 1 set "PYTHON_EXE=python"
)

if not defined PYTHON_EXE (
  echo [ERROR] Python was not found.
  echo Install Python, or run this from Codex after dependencies are available.
  echo.
  pause
  exit /b 1
)

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "TS=%%i"

if exist "%SOURCE_FILE%" (
  copy /y "%SOURCE_FILE%" "%SOURCE_BACKUP_DIR%\c4c_ticket_table_z007_z010_checked_hana_final_%TS%.xlsx" >nul
  echo [OK] Backed up source workbook
)

if exist "%DASHBOARD_DATA_FILE%" (
  copy /y "%DASHBOARD_DATA_FILE%" "%BACKUP_DIR%\dashboard-data_%TS%.js" >nul
  echo [OK] Backed up existing dashboard-data.js
)

echo [RUN] Fetching latest C4C ticket data...
echo [%DATE% %TIME%] Fetching latest C4C ticket data > "%LOG_FILE%"
set "OUTPUT_FILE=%C4C_EXPORT_FILE%"
set "SOURCE_TICKET_FILE="
set "C4C_FETCH_HISTORY=false"
set "C4C_VERIFY_SSL=false"
set "C4C_TIMEOUT=60"
%PYTHON_EXE% "%FETCH_SCRIPT_FILE%" >> "%LOG_FILE%" 2>&1

if errorlevel 1 (
  echo.
  echo [ERROR] C4C ticket fetch failed. Log:
  echo %LOG_FILE%
  echo.
  type "%LOG_FILE%"
  echo.
  pause
  exit /b 1
)

echo [RUN] Enriching invoice data from SAP HANA...
echo.>> "%LOG_FILE%"
echo [%DATE% %TIME%] Enriching invoice data from SAP HANA >> "%LOG_FILE%"
set "WORKBOOK_PATH=%C4C_EXPORT_FILE%"
set "OUTPUT_PATH=%SOURCE_FILE%"
%PYTHON_EXE% "%HANA_SCRIPT_FILE%" >> "%LOG_FILE%" 2>&1

if errorlevel 1 (
  echo.
  echo [ERROR] SAP HANA invoice enrich failed. Log:
  echo %LOG_FILE%
  echo.
  type "%LOG_FILE%"
  echo.
  pause
  exit /b 1
)

echo [RUN] Refreshing lastchangedtime from C4C status history...
echo.>> "%LOG_FILE%"
echo [%DATE% %TIME%] Refreshing lastchangedtime from C4C status history >> "%LOG_FILE%"
set "SOURCE_WORKBOOK=%SOURCE_FILE%"
set "C4C_FETCH_HISTORY=true"
set "C4C_VERIFY_SSL=false"
set "C4C_HISTORY_ENV=PC4C"
set "C4C_HISTORY_SCOPE=all"
set "C4C_HISTORY_WORKERS=16"
set "C4C_TIMEOUT=30"
%PYTHON_EXE% "%HISTORY_SCRIPT_FILE%" >> "%LOG_FILE%" 2>&1

if errorlevel 1 (
  echo.
  echo [ERROR] History update failed. Log:
  echo %LOG_FILE%
  echo.
  type "%LOG_FILE%"
  echo.
  pause
  exit /b 1
)

echo [RUN] Rebuilding dashboard-data.js...
echo.>> "%LOG_FILE%"
echo [%DATE% %TIME%] Rebuilding dashboard-data.js >> "%LOG_FILE%"
%PYTHON_EXE% "%SCRIPT_FILE%" >> "%LOG_FILE%" 2>&1

if errorlevel 1 (
  echo.
  echo [ERROR] Update failed. Log:
  echo %LOG_FILE%
  echo.
  type "%LOG_FILE%"
  echo.
  pause
  exit /b 1
)

powershell -NoProfile -Command "$p='%CD%\overview.html'; $utf8=New-Object System.Text.UTF8Encoding($false); $c=[System.IO.File]::ReadAllText($p,[System.Text.Encoding]::UTF8); $c=$c -replace 'dashboard-data\.js\?v=[^\"'']+', 'dashboard-data.js?v=%TS%'; [System.IO.File]::WriteAllText($p,$c,$utf8)"
echo [OK] Refreshed dashboard-data.js cache version

if exist "%CD%\dist" (
  copy /y "%CD%\overview.html" "%CD%\dist\index.html" >nul
  copy /y "%CD%\overview.html" "%CD%\dist\overview.html" >nul
  copy /y "%CD%\dashboard-data.js" "%CD%\dist\dashboard-data.js" >nul
  if not exist "%CD%\dist\assets" mkdir "%CD%\dist\assets"
  copy /y "%CD%\assets\service-order-icons.js" "%CD%\dist\assets\service-order-icons.js" >nul
  copy /y "%CD%\assets\lucide-LICENSE" "%CD%\dist\assets\lucide-LICENSE" >nul
  echo [OK] Synced files to dist\
)

echo.
echo [SUCCESS] Dashboard data updated.
echo [SOURCE]  %SOURCE_FILE%
echo [OUTPUT]  %DASHBOARD_DATA_FILE%
echo [LOG]     %LOG_FILE%
echo.
echo If Render is already connected to GitHub, commit and push this change,
echo then trigger a new deploy.
echo.
pause
