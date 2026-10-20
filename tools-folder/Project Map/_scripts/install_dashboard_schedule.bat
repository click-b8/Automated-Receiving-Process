@echo off
rem One-time setup: refresh the dashboard every 15 minutes in the background (no window).
schtasks /Create /F /SC MINUTE /MO 15 /TN "Receiving Dashboard Refresh" /TR "pythonw.exe \"%~dp0refresh_all.py\""
if errorlevel 1 (echo Could not create the task. & pause & exit /b 1)
schtasks /Run /TN "Receiving Dashboard Refresh"
echo Scheduled: "Receiving Dashboard Refresh" runs every 15 minutes. Remove it with: schtasks /Delete /TN "Receiving Dashboard Refresh" /F
pause
