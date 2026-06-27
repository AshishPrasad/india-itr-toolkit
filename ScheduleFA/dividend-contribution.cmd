@echo off
REM Windows convenience wrapper for the dividend-contribution CLI.
REM Usage:
REM   dividend-contribution.cmd --dividends div.xlsx --stocks stocks.xlsx --output out.xlsx
setlocal
set "PYTHONPATH=%~dp0src;%PYTHONPATH%"
python -m dividend_contribution %*
endlocal
