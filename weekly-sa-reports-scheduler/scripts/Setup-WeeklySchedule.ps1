<#
.SYNOPSIS
    Creates/updates the input-output folder structure and the Windows weekly
    scheduled task that runs Run-WeeklyReports.ps1.

.DESCRIPTION
    Idempotent: safe to re-run any time you want to change the day, time, or
    reports root folder. Deletes and recreates the named scheduled task with
    the new settings.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Setup-WeeklySchedule.ps1 `
        -ReportsRoot "$env:USERPROFILE\WeeklyReports" -DayOfWeek THU -Time 08:15
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ReportsRoot,
    [ValidateSet('MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN')]
    [string]$DayOfWeek = 'WED',
    [string]$Time = '08:15',
    [string]$TaskName = 'WeeklyReports-SARB-TechPass-SAPerformance',
    [switch]$RunNow
)

$ErrorActionPreference = 'Stop'
$scriptDir = $PSScriptRoot
$runnerPath = Join-Path $scriptDir 'Run-WeeklyReports.ps1'

foreach ($sub in @('inputs\work-queue', 'inputs\techpass', 'outputs', 'logs')) {
    New-Item -ItemType Directory -Force -Path (Join-Path $ReportsRoot $sub) | Out-Null
}
Write-Host "Folder structure ready under: $ReportsRoot"
Write-Host "  Drop the weekly Work Queue export in:  $ReportsRoot\inputs\work-queue\"
Write-Host "  Drop the weekly TechPass export in:    $ReportsRoot\inputs\techpass\"

$action = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$runnerPath`" -ReportsRoot `"$ReportsRoot`""

cmd /c "schtasks /Query /TN `"$TaskName`" >nul 2>nul"
if ($LASTEXITCODE -eq 0) {
    schtasks /Delete /TN $TaskName /F | Out-Null
}
schtasks /Create /SC WEEKLY /D $DayOfWeek /ST $Time /TN $TaskName /TR $action /F | Out-Null

Write-Host ""
Write-Host "Scheduled task '$TaskName' set for every $DayOfWeek at $Time." -ForegroundColor Green
schtasks /Query /TN $TaskName /V /FO LIST | Select-String "TaskName|Next Run Time|Status"

if ($RunNow) {
    Write-Host ""
    Write-Host "Running now to validate..." -ForegroundColor Cyan
    & $runnerPath -ReportsRoot $ReportsRoot -OpenOutputs $false
}
