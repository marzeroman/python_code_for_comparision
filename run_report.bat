@echo off
REM ============================================================
REM  Runs a saved report profile. Used by Windows Task Scheduler.
REM  1) Change "example_monthly_report" below to your profile name.
REM  2) Schedule it (run once in Command Prompt, adjust the path):
REM     schtasks /create /tn "Monthly report" /tr "C:\path\to\run_report.bat" /sc monthly /d 1 /st 08:00
REM  Results go to the results folder; every run is listed in logs\run_log.csv
REM ============================================================
cd /d "%~dp0"
python report_tool.py --profile example_monthly_report
exit /b %ERRORLEVEL%