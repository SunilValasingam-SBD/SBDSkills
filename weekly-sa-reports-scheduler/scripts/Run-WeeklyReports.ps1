<#
.SYNOPSIS
    Weekly runner for the TechPass Aging, SARB SLA Report, SARB SLA Reasoning,
    and SA KPI Heatmap (performance) skills.

.DESCRIPTION
    Picks the newest file dropped in:
      inputs\work-queue\  -> used by SARB SLA Report, SARB SLA Reasoning, SA KPI Heatmap
      inputs\techpass\    -> used by TechPass Aging Report
    and writes all generated reports into a dated folder under .\outputs\.

    Drop a new weekly export into each inputs subfolder (any filename, .csv or
    .xlsx) before this runs — old files can stay for history, only the most
    recently modified one in each folder is used.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Run-WeeklyReports.ps1 -ReportsRoot "$env:USERPROFILE\WeeklyReports"
    powershell -ExecutionPolicy Bypass -File .\Run-WeeklyReports.ps1 -ReportsRoot "$env:USERPROFILE\WeeklyReports" -AsOfDate 2026-10-14
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ReportsRoot,
    [string]$AsOfDate = (Get-Date).ToString('yyyy-MM-dd'),
    [bool]$OpenOutputs = $true
)

$ErrorActionPreference = 'Stop'
$root = $ReportsRoot
$skillsRoot = Split-Path $PSScriptRoot -Parent | Split-Path -Parent
$logPath = Join-Path $root "logs\run-$AsOfDate.log"

function Write-Log($message) {
    $line = "[$(Get-Date -Format 'HH:mm:ss')] $message"
    Write-Host $line
    Add-Content -Path $logPath -Value $line
}

function Get-NewestInput($folder, $label) {
    $file = Get-ChildItem $folder -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch '^~\$' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if (-not $file) {
        Write-Log "WARNING: No $label file found in $folder - skipping dependent reports."
        return $null
    }
    $ageDays = [math]::Round(((Get-Date) - $file.LastWriteTime).TotalDays, 1)
    if ($ageDays -gt 8) {
        Write-Log "WARNING: Newest $label file '$($file.Name)' is $ageDays days old - may be stale. Drop a fresh export before running."
    } else {
        Write-Log "Using $label file: $($file.Name) (modified $ageDays day(s) ago)"
    }
    return $file.FullName
}

New-Item -ItemType Directory -Force -Path (Join-Path $root 'logs') | Out-Null
"==== Weekly Reports run: $AsOfDate ====" | Out-File -FilePath $logPath -Encoding UTF8

$runDir = Join-Path $root "outputs\$AsOfDate"
$techpassOut = Join-Path $runDir 'techpass-aging'
$sarbOut = Join-Path $runDir 'sarb-sla-report'
$reasoningOut = Join-Path $runDir 'sarb-sla-reasoning'
$perfOut = Join-Path $runDir 'sa-performance'
foreach ($d in @($runDir, $techpassOut, $sarbOut, $reasoningOut, $perfOut)) {
    New-Item -ItemType Directory -Force -Path $d | Out-Null
}

$workQueueFile = Get-NewestInput (Join-Path $root 'inputs\work-queue') 'Work Queue'
$techpassFile = Get-NewestInput (Join-Path $root 'inputs\techpass') 'TechPass'

# --- TechPass Aging Report -------------------------------------------------
if ($techpassFile) {
    Write-Log "Running TechPass Aging Report..."
    try {
        Push-Location (Join-Path $skillsRoot 'techpass-aging')
        & py .\techpass_aging_chart.py --input $techpassFile --as-of $AsOfDate `
            --output (Join-Path $techpassOut 'techpass_aging_by_architect.html') `
            --pptx-output (Join-Path $techpassOut 'techpass_aging_summary.pptx') 2>&1 |
            ForEach-Object { Write-Log "  $_" }
        Pop-Location
    } catch {
        Write-Log "ERROR running TechPass Aging Report: $_"
    }
}

if ($workQueueFile) {
    # run-sarb-sla.ps1 reads with Import-Csv, so convert Excel input to CSV first.
    $workQueueCsv = $workQueueFile
    if ($workQueueFile -match '\.xlsx?$') {
        $workQueueCsv = Join-Path $runDir 'work-queue-source.csv'
        Write-Log "Converting Work Queue Excel export to CSV..."
        try {
            & py -c "import pandas as pd; pd.read_excel(r'$workQueueFile', dtype=str).to_csv(r'$workQueueCsv', index=False)"
        } catch {
            Write-Log "ERROR converting Work Queue file to CSV: $_"
            $workQueueCsv = $null
        }
    }

    # --- SARB SLA Report ----------------------------------------------------
    if ($workQueueCsv) {
        Write-Log "Running SARB SLA Report..."
        try {
            powershell -ExecutionPolicy Bypass -File (Join-Path $skillsRoot 'sarb-sla-report\scripts\run-sarb-sla.ps1') `
                -Source $workQueueCsv -AsOfDate $AsOfDate -OutputDirectory $sarbOut 2>&1 |
                ForEach-Object { Write-Log "  $_" }
        } catch {
            Write-Log "ERROR running SARB SLA Report: $_"
        }
    }

    # --- SARB SLA Reasoning Analysis ----------------------------------------
    Write-Log "Running SARB SLA Reasoning Analysis..."
    try {
        Push-Location (Join-Path $skillsRoot 'sarb-sla-reasoning')
        & py .\sarb_sla_reasoning.py --input $workQueueFile --as-of $AsOfDate --output-dir $reasoningOut `
            --output (Join-Path $reasoningOut 'sarb-sla-reasoning.html') 2>&1 |
            ForEach-Object { Write-Log "  $_" }
        Pop-Location
    } catch {
        Write-Log "ERROR running SARB SLA Reasoning Analysis: $_"
    }

    # --- SA KPI Heatmap (SA Performance) ------------------------------------
    Write-Log "Running SA KPI Heatmap..."
    try {
        Push-Location (Join-Path $skillsRoot 'sa-kpi-heatmap')
        & py .\sa_kpi_heatmap.py --input $workQueueFile `
            --output (Join-Path $perfOut 'sa_kpi_heatmap.html') `
            --summary-output (Join-Path $perfOut 'sa_kpi_summary.csv') 2>&1 |
            ForEach-Object { Write-Log "  $_" }
        Pop-Location
    } catch {
        Write-Log "ERROR running SA KPI Heatmap: $_"
    }
}

Write-Log "Done. Outputs in: $runDir"

if ($OpenOutputs) {
    try { Invoke-Item $runDir } catch { Write-Log "Could not auto-open output folder (no interactive session?): $_" }
}
