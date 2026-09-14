@echo off
call "%~dp0RUN_UPDATE_DATA.bat" --publish-only
exit /b %errorlevel%
