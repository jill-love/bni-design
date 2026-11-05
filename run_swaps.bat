@echo off
rem BNI pot swaps: replace failed pots with spares in a design map workbook.
cd /d "%~dp0"
py -3.14 swap_app.py %*
if errorlevel 1 pause
