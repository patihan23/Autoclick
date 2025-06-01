@echo off
echo กำลังรันโปรแกรม Auto Action Clicker ในฐานะ Administrator...
echo.
powershell -Command "Start-Process python -ArgumentList 'auto_action_clicker.py' -Verb RunAs -WorkingDirectory '%cd%'"
pause
