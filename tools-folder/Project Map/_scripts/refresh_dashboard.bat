@echo off
rem Refresh the Obsidian dashboard and file map now.
python "%~dp0refresh_all.py"
echo.
echo Done. Last results:
powershell -NoProfile -Command "Get-Content -Tail 1 \"%~dp0.refresh_log.txt\""
pause
