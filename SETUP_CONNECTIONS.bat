@echo off
call "%~dp0RUN_UPDATE_DATA.bat" --setup-connections
exit /b %errorlevel%
