@echo off
REM Windows convenience wrapper for the schedule-fa-csv CLI.
REM Usage:
REM   schedule-fa-csv.cmd --input holdings.xlsx --output schedule_fa.csv
setlocal
set "PYTHONPATH=%~dp0src;%PYTHONPATH%"
python -m schedule_fa_csv %*
endlocal
