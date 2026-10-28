@echo off
rem BNI experiment setup: plan an experiment and write its design map workbook.
cd /d "%~dp0"
py -3.14 design_app.py
if errorlevel 1 pause
