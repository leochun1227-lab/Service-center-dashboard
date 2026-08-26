@echo off
setlocal EnableExtensions

cd /d "%~dp0"
title Service Center Dashboard - Update Data

echo ============================================================
echo  Service Center Dashboard - Update Data
echo ============================================================
echo.
echo  This will rebuild dashboard-data.js from:
echo  c4c_ticket_table_z007_z010_checked_hana_final.xlsx
echo.

set "SOURCE_FILE=%CD%\c4c_ticket_table_z007_z010_checked_hana_final.xlsx"
set "SYSTEM_SOURCE_FILE=%USERPROFILE%\Documents\ChatGPT\Service centre dashboard\c4c_ticket_table_z007_z010_checked_hana_final.xlsx"
set "OUTPUT_FILE=%CD%\dashboard-data.js"
set "SCRIPT_FILE=%CD%\generate_web_data_from_excel.py"
set "LOG_DIR=%CD%\outputs"
set "BACKUP_DIR=%LOG_DIR%\dashboard-data-backups"
set "LOG_FILE=%LOG_DIR%\update_dashboard_data.log"

if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"
if not exist "%BACKUP_DIR%" mkdir "%BACKUP_DIR%"

if exist "%SYSTEM_SOURCE_FILE%" (
  echo [SYNC] Using latest system-generated workbook:
  echo        %SYSTEM_SOURCE_FILE%
  copy /y "%SYSTEM_SOURCE_FILE%" "%SOURCE_FILE%" >nul
) else (
  echo [WARN] System-generated workbook was not found:
  echo        %SYSTEM_SOURCE_FILE%
  echo        Falling back to workbook in this folder.
)
echo.

if not exist "%SOURCE_FILE%" (
  echo [ERROR] Source Excel not found:
  echo %SOURCE_FILE%
  echo.
  echo Put the latest source workbook in this folder, then run again.
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

if exist "%OUTPUT_FILE%" (
  copy /y "%OUTPUT_FILE%" "%BACKUP_DIR%\dashboard-data_%TS%.js" >nul
  echo [OK] Backed up existing dashboard-data.js
)

echo [RUN] Rebuilding dashboard-data.js...
echo [%DATE% %TIME%] Rebuilding dashboard-data.js > "%LOG_FILE%"
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
  echo [OK] Synced files to dist\
)

echo.
echo [SUCCESS] Dashboard data updated.
echo [OUTPUT]  %OUTPUT_FILE%
echo [LOG]     %LOG_FILE%
echo.
echo If Render is already connected to GitHub, commit and push this change,
echo then trigger a new deploy.
echo.
pause
