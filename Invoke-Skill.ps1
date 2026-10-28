<#
.SYNOPSIS
    Interactive picker/launcher for all skills under this folder.

.DESCRIPTION
    Reads skills-manifest.json (one entry per skill, built from each skill's SKILL.md
    and underlying script arguments), lets you pick a skill, shows what input it
    expects on one page, prompts for each parameter, then runs it.

    "script" mode skills run immediately (python/powershell). "assisted" mode skills
    (content-authoring, e.g. SAD generator, swimlane diagram) are run by the Copilot
    model, not a single script — this launcher shows the inputs it needs and prints
    the exact instruction to paste into a Copilot CLI chat.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Invoke-Skill.ps1
#>

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$manifestPath = Join-Path $root 'skills-manifest.json'

if (-not (Test-Path $manifestPath)) {
    Write-Error "Manifest not found: $manifestPath"
    exit 1
}

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$manifest = Get-Content $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$skills = $manifest.skills

function Show-Menu {
    Write-Host ""
    Write-Host "=== Available Skills ===" -ForegroundColor Cyan
    for ($i = 0; $i -lt $skills.Count; $i++) {
        $s = $skills[$i]
        $modeTag = if ($s.mode -eq 'assisted') { '[assisted]' } else { '[script]' }
        Write-Host ("  {0,2}. {1,-32} {2}" -f ($i + 1), $s.name, $modeTag)
    }
    Write-Host "   0. Exit"
    Write-Host ""
}

function Show-SkillDetails($s) {
    Write-Host ""
    Write-Host "=== $($s.name) ===" -ForegroundColor Yellow
    Write-Host "Mode: $($s.mode)"
    Write-Host ""
    Write-Host "Description:"
    Write-Host "  $($s.description)"
    if ($s.note) {
        Write-Host ""
        Write-Host "Note:" -ForegroundColor DarkYellow
        Write-Host "  $($s.note)"
    }
    Write-Host ""
    Write-Host "Inputs expected:"
    foreach ($p in $s.params) {
        $req = if ($p.required) { 'required' } else { 'optional' }
        Write-Host ("  - {0,-20} ({1}, {2}) - {3}" -f $p.flag, $p.type, $req, $p.prompt)
    }
    Write-Host ""
}

function Read-ParamValue($p) {
    $suffix = if ($p.required) { " [required]" } else { " [optional]" }
    $value = Read-Host "  $($p.prompt)$suffix"
    while ($p.required -and [string]::IsNullOrWhiteSpace($value)) {
        Write-Host "    This value is required." -ForegroundColor Red
        $value = Read-Host "  $($p.prompt)$suffix"
    }
    return $value
}

function Invoke-ScriptSkill($s) {
    $skillDir = Join-Path $root $s.folder
    $entryPath = Join-Path $skillDir $s.entry
    if (-not (Test-Path $entryPath)) {
        Write-Error "Entry script not found: $entryPath"
        return
    }

    # Python (argparse) scripts take a flat array of tokens ("--flag", "value", ...).
    # PowerShell scripts have a typed param() block, so named args must be splatted
    # from a hashtable (@{ Name = value }) — splatting a flat array binds positionally
    # instead of by name and silently corrupts the values.
    $argList = @()
    $argHash = @{}
    foreach ($p in $s.params) {
        $value = Read-ParamValue $p
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        $paramName = $p.flag.TrimStart('-')

        switch ($p.type) {
            'switch' {
                if ($value -match '^(y|yes)$') {
                    $argList += $p.flag
                    $argHash[$paramName] = $true
                }
            }
            'repeat-csv' {
                $items = @($value -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                foreach ($item in $items) { $argList += @($p.flag, $item) }
                $argHash[$paramName] = $items
            }
            default {
                $argList += @($p.flag, $value)
                $argHash[$paramName] = $value
            }
        }
    }

    Write-Host ""
    Write-Host "Running:" -ForegroundColor Green
    if ($s.runner -eq 'python') {
        Write-Host "  py `"$entryPath`" $($argList -join ' ')"
        Push-Location $skillDir
        try { & py $entryPath @argList } finally { Pop-Location }
    } else {
        Write-Host "  powershell -File `"$entryPath`" $($argList -join ' ')"
        Push-Location $skillDir
        try { & $entryPath @argHash } finally { Pop-Location }
    }
}

function Invoke-AssistedSkill($s) {
    Write-Host ""
    Write-Host "This skill needs to be run by the Copilot model (not a standalone script)." -ForegroundColor Yellow
    $answers = @()
    foreach ($p in $s.params) {
        $value = Read-ParamValue $p
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $answers += "- $($p.flag): $value"
        }
    }
    Write-Host ""
    Write-Host "Paste this into a Copilot CLI chat to run it:" -ForegroundColor Green
    Write-Host "----------------------------------------------"
    Write-Host "Use the $($s.id) skill."
    if ($answers.Count -gt 0) {
        Write-Host ($answers -join "`n")
    }
    Write-Host "----------------------------------------------"
}

while ($true) {
    Show-Menu
    $choice = Read-Host "Pick a skill number"
    if ($choice -eq '0' -or [string]::IsNullOrWhiteSpace($choice)) { break }

    $index = 0
    if (-not [int]::TryParse($choice, [ref]$index) -or $index -lt 1 -or $index -gt $skills.Count) {
        Write-Host "Invalid choice." -ForegroundColor Red
        continue
    }

    $selected = $skills[$index - 1]
    Show-SkillDetails $selected

    $confirm = Read-Host "Run this skill now? (Y/n)"
    if ($confirm -match '^(n|no)$') { continue }

    if ($selected.mode -eq 'assisted') {
        Invoke-AssistedSkill $selected
    } else {
        Invoke-ScriptSkill $selected
    }
}

Write-Host "Goodbye."
