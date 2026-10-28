<#
.SYNOPSIS
    Shared Inputs/Outputs path helpers for every SA skill (PowerShell skills).

.DESCRIPTION
    Dot-source this file from a skill's script to get $SAInputsDir /
    $SAOutputsDir plus helper functions, mirroring sa_paths.py (the Python
    equivalent) so every skill shares one drop folder for source files and
    one destination root for generated reports.

.EXAMPLE
    . (Join-Path $PSScriptRoot '..\_shared\SA-Paths.ps1')
    $outDir = Get-SAOutputDir -SkillName 'my-skill-name'
    $src = if ($Source) { $Source } else { Get-SALatestInput -Patterns '*.xlsx','*.csv' }
#>

$SARoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$SAInputsDir = Join-Path $SARoot 'Inputs'
$SAOutputsDir = Join-Path $SARoot 'Outputs'

function Get-SAOutputDir {
    param([Parameter(Mandatory)][string]$SkillName)
    $dir = Join-Path $SAOutputsDir $SkillName
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    return $dir
}

function Get-SALatestInput {
    param([string[]]$Patterns = @('*.xlsx', '*.csv'))
    if (-not (Test-Path $SAInputsDir)) { return $null }
    $candidates = foreach ($pat in $Patterns) {
        Get-ChildItem -Path $SAInputsDir -Filter $pat -File -ErrorAction SilentlyContinue
    }
    if (-not $candidates) { return $null }
    return ($candidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
}
