@echo off
:: Launcher for PowerBI Assessment scripts
:: Uses -ExecutionPolicy Bypass scoped to the process only (no admin, no system changes)

echo Select the script to run:
echo  1. Get PowerBI Workspaces
echo  2. Get PowerBI Reports By Workspace
echo  3. Get Usage Metric By Report
echo  4. Get Fabric Audit Logs
echo.
set /p choice=Enter number (1-4): 

if "%choice%"=="1" powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp001-Get-PowerBIWorkspaces.ps1"
if "%choice%"=="2" powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp002-Get-PowerBIReportsByWorkspace.ps1"
if "%choice%"=="3" powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp003-Get-UsageMetricByReport.ps1"
if "%choice%"=="4" powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp004-Get-FabricAuditLogs.ps1"

pause
