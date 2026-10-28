param(
    [string]$Source,
    [datetime]$AsOfDate = (Get-Date).Date,
    [string]$OutputDirectory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot '..\..\_shared\SA-Paths.ps1')

if (-not $Source) {
    $Source = Get-SALatestInput -Patterns '*.csv'
    if (-not $Source) {
        throw "No -Source given and no .csv file found in the shared SA\Inputs folder."
    }
    Write-Host "No -Source given; using newest file in SA\Inputs: $Source"
}
if (-not $OutputDirectory) {
    $OutputDirectory = Get-SAOutputDir -SkillName 'sarb-sla-report'
}

function Parse-Date([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    return ([datetime]::Parse($Value, [Globalization.CultureInfo]::InvariantCulture)).Date
}

function Get-ExpectedSarbDate([datetime]$Created, [string]$WorkRef) {
    # Projects (PRJ) get a 14-calendar-day SLA window; Ideas (ID) get 28 calendar days.
    $days = if ($WorkRef -match '^ID') { 28 } else { 14 }
    $date = $Created.AddDays($days)
    while ($date.DayOfWeek -ne [DayOfWeek]::Tuesday) { $date = $date.AddDays(1) }
    return $date
}

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$allRows = Import-Csv $Source
# Ideas that have already been converted to a Project are referenced by a PRJ
# row's 'Idea ID' value. Exclude those Idea rows so the same initiative is not
# counted twice (once as an Idea, once as its resulting Project).
$convertedIdeaIds = @($allRows | Where-Object { $_.'Work Ref' -match '^PRJ' -and -not [string]::IsNullOrWhiteSpace($_.'Idea ID') } | Select-Object -ExpandProperty 'Idea ID' | Sort-Object -Unique)
$filtered = $allRows | Where-Object {
    $created = Parse-Date $_.Created
    $isPrj = $_.'Work Ref' -match '^PRJ'
    $isStandaloneIdea = ($_.'Work Ref' -match '^ID') -and ($convertedIdeaIds -notcontains $_.'Work Ref')
    $created.Year -eq 2026 -and ($isPrj -or $isStandaloneIdea) -and $_.Status -ne 'Cancelled'
} | ForEach-Object {
    $created = Parse-Date $_.Created
    $expected = Get-ExpectedSarbDate $created $_.'Work Ref'
    $approval = Parse-Date $_.'SARB Approval Date'
    if ($approval) {
        $status = if ($approval -le $expected) { 'Met' } else { 'Not met' }
    } elseif ($AsOfDate.Date -gt $expected) {
        $status = 'Not met'
    } else {
        $status = 'Not yet due'
    }
    $_ | Add-Member -NotePropertyName 'Expected SARB Date' -NotePropertyValue $expected.ToString('yyyy-MM-dd') -PassThru |
        Add-Member -NotePropertyName 'SLA Status' -NotePropertyValue $status -PassThru
}

$analyzedPath = Join-Path $OutputDirectory 'analyzed-2026-prj.csv'
$filtered | Export-Csv $analyzedPath -NoTypeInformation -Encoding UTF8
$summary = $filtered | Group-Object 'Work Ref: Solution Architect' | ForEach-Object {
    $group = $_.Group
    $met = @($group | Where-Object 'SLA Status' -eq 'Met').Count
    $notMet = @($group | Where-Object 'SLA Status' -eq 'Not met').Count
    $pending = @($group | Where-Object 'SLA Status' -eq 'Not yet due').Count
    [pscustomobject]@{
        'Solution Architect' = if ([string]::IsNullOrWhiteSpace($_.Name)) { '(Unassigned)' } else { $_.Name }
        Total = $group.Count
        Met = $met
        'Not met' = $notMet
        'Not yet due' = $pending
        'SLA %' = if ($met + $notMet) { [math]::Round(100 * $met / ($met + $notMet), 1) } else { $null }
    }
} | Sort-Object 'SLA %' -Descending
$summary | Export-Csv (Join-Path $OutputDirectory 'summary-2026-prj.csv') -NoTypeInformation -Encoding UTF8

$metCount = @($filtered | Where-Object 'SLA Status' -eq 'Met').Count
$notMetCount = @($filtered | Where-Object 'SLA Status' -eq 'Not met').Count
$pendingCount = @($filtered | Where-Object 'SLA Status' -eq 'Not yet due').Count
$rate = if ($metCount + $notMetCount) { [math]::Round(100 * $metCount / ($metCount + $notMetCount), 1) } else { 0 }
$maxCount = [math]::Max(1, @($summary | ForEach-Object { [int]$_.Met; [int]$_.'Not met' } | Measure-Object -Maximum).Maximum)
$chartRows = foreach ($sa in $summary) {
    $metHeight = [math]::Round(100 * [int]$sa.Met / $maxCount, 1)
    $notHeight = [math]::Round(100 * [int]$sa.'Not met' / $maxCount, 1)
    $label = $sa.'Solution Architect'
    "<div class=""chart-row""><div class=""bar-column"" title=""$($sa.'Solution Architect'): $($sa.Met) met""><div class=""bar-fill bar-fill-met"" style=""height:$metHeight%""><span>$($sa.Met)</span></div></div><div class=""bar-column"" title=""$($sa.'Solution Architect'): $($sa.'Not met') not met""><div class=""bar-fill bar-fill-not"" style=""height:$notHeight%""><span>$($sa.'Not met')</span></div></div><small>$label</small></div>"
}
$chartHtml = @"
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SARB SLA Performance - $($AsOfDate.ToString('yyyy-MM-dd'))</title>
<style>
body{font-family:Segoe UI,sans-serif;max-width:1100px;margin:32px auto;padding:0 20px;color:#17212b;background:#f4f7f8}
h1{margin-bottom:4px}.muted{color:#5c6b73}.cards{display:flex;gap:12px;flex-wrap:wrap;margin:24px 0}.card{background:#fff;border:1px solid #d9e2e5;border-radius:8px;padding:14px 18px;min-width:120px}.card strong{display:block;font-size:25px;margin-top:4px}
.chart{background:#fff;border:1px solid #d9e2e5;border-radius:8px;padding:20px}.legend{margin-bottom:16px;color:#5c6b73}.legend span{margin-right:18px}.met{color:#1c8c72}.not{color:#c24b3d}
.plot{height:360px;min-width:1080px;display:flex;align-items:end;gap:18px;border-left:1px solid #9fb1b5;border-bottom:1px solid #9fb1b5;padding:0 18px;background:repeating-linear-gradient(to bottom,transparent 0,transparent 19.8%,#dce6e8 20%,#dce6e8 calc(20% + 1px))}
.chart{overflow-x:auto}.chart-row{height:100%;min-width:110px;flex:1;display:flex;align-items:end;justify-content:center;gap:5px;position:relative;padding-bottom:42px}.bar-column{height:100%;width:24px;display:flex;align-items:center;justify-content:end;flex-direction:column}.bar-fill{width:100%;border-radius:3px 3px 0 0;position:relative;min-height:1px}.bar-fill-met{background:#1c8c72}.bar-fill-not{background:#c24b3d}.bar-fill span{position:absolute;top:-18px;font-size:12px}.chart-row small{position:absolute;bottom:8px;white-space:normal;width:110px;text-align:center;line-height:1.2;color:#5c6b73}
table{width:100%;border-collapse:collapse;background:#fff;margin-top:24px}th,td{padding:10px 12px;border-bottom:1px solid #e5ecee;text-align:left}th{background:#173f4a;color:#fff}
</style>
</head>
<body>
<h1>SARB SLA Performance</h1>
<div class="muted">As of $($AsOfDate.ToString('yyyy-MM-dd')) | Created in 2026 | PRJ projects and standalone Ideas (not yet converted to a Project) | Cancelled items excluded</div>
<div class="cards"><div class="card">Assigned projects<strong>$($filtered.Count)</strong></div><div class="card">SLA met<strong>$metCount</strong></div><div class="card">SLA not met<strong>$notMetCount</strong></div><div class="card">Not yet due<strong>$pendingCount</strong></div><div class="card">Team SLA rate<strong>$rate%</strong></div></div>
<h2>Met versus not met by Solution Architect</h2>
<div class="chart"><div class="legend"><span class="met">SLA met</span><span class="not">SLA not met</span></div><div class="plot">$($chartRows -join [Environment]::NewLine)</div></div>
<h2>Solution Architect performance</h2>
$($summary | ConvertTo-Html -Fragment)
</body>
</html>
"@
$chartHtml | Set-Content (Join-Path $OutputDirectory 'SARB-SLA-report.html') -Encoding UTF8
"Analyzed $($filtered.Count) rows. Output: $analyzedPath"
