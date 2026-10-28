param(
    [Parameter(Mandatory = $true)]
    [string]$ArchitectName,
    [string]$Source,
    [string]$OutputDirectory,
    [int]$Year = (Get-Date).Year
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
    $OutputDirectory = Get-SAOutputDir -SkillName 'sa-project-status-review'
}

function Parse-Date([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    try {
        return ([datetime]::Parse($Value, [Globalization.CultureInfo]::InvariantCulture)).Date
    } catch {
        return $null
    }
}

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$rows = Import-Csv $Source
$architectStatus = @($rows | Where-Object {
    try {
        $created = Parse-Date $_.'Created'
        $created -and $created.Year -eq $Year -and
            $_.'Work Ref' -match '^PRJ' -and
            $_.Status -ne 'Cancelled' -and
            $_.'Work Ref: Solution Architect' -eq $ArchitectName
    } catch {
        $false
    }
} | ForEach-Object {
    $_ | Add-Member -NotePropertyName 'SAD Approval Status' -NotePropertyValue $(
        if ([string]::IsNullOrWhiteSpace($_.'SAD Approved')) { 'Missing SAD Approval' } else { 'Present' }
    ) -PassThru |
        Add-Member -NotePropertyName 'Scenario Status' -NotePropertyValue $(
        if ($_.'Publication Status' -match '^Published') { 'Scenario published' } else { 'Not scenario published' }
    ) -PassThru
})

$outputPath = Join-Path $OutputDirectory 'sa-project-status.csv'
$architectStatus | Select-Object 'Work Ref', 'Work Ref: Name', 'Work Ref: Solution Architect',
    'SAD Approved', 'Publication Status', 'SAD Approval Status', 'Scenario Status' |
    Export-Csv $outputPath -NoTypeInformation -Encoding UTF8

$missingSad = @($architectStatus | Where-Object 'SAD Approval Status' -eq 'Missing SAD Approval').Count
$sadPresent = @($architectStatus | Where-Object 'SAD Approval Status' -eq 'Present').Count
$published = @($architectStatus | Where-Object 'Scenario Status' -eq 'Scenario published').Count
$notPublished = @($architectStatus | Where-Object 'Scenario Status' -eq 'Not scenario published').Count

"Solution Architect project analysis: $ArchitectName"
"Reporting year: $Year"
"Total projects: $($architectStatus.Count)"
"Missing SAD Approval: $missingSad"
"SAD Approval present: $sadPresent"
"Scenario published: $published"
"Not scenario published: $notPublished"
"Output: $outputPath"
