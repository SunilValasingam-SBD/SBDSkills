# SARB SLA Report

## Executive Summary

Every project and idea submitted to the Solution Architecture Review Board (SARB) is expected to receive an architecture decision within a defined turnaround window. This skill scores that commitment for the team and each Solution Architect using a Work Queue export. It replaces manual due-date and approval-date tracking with a repeatable report for leadership.

## Business Process Flow

1. A team member exports the current Solution Architecture Work Queue list from the SharePoint site to a CSV file (this is a manual SharePoint export step; the tool does not log into SharePoint itself).
2. The exported CSV is dropped into the shared `SA\Inputs` folder (or an explicit file path is supplied), and the skill is run — either via the `/sarb-sla-report` Copilot CLI command or directly from PowerShell.
3. The tool automatically decides which records count: only Projects and Ideas created in the current (or requested) calendar year, excluding Cancelled items, and excluding an Idea once it has been converted into a Project (so the same piece of work is never counted twice).
4. For each in-scope record, the tool calculates the expected SARB decision date — 14 calendar days after creation for a Project, or 28 calendar days for an Idea, moved forward to the next Tuesday if it doesn't already land on one — and compares it against the actual approval date (or today's date, if not yet approved) to classify each record as **Met**, **Not met**, or **Not yet due**.
5. The tool rolls those classifications up into a team-wide SLA percentage and a per-Solution-Architect scorecard.
6. The deliverables — detailed and summary spreadsheets plus a visual HTML report with a horizontal bar chart — are written automatically to the shared `SA\Outputs\sarb-sla-report` folder, ready to open, share, or paste into a leadership update with no further manipulation.

## Technical Flow

- **Entry point:** `scripts\run-sarb-sla.ps1`, a PowerShell script invoked either through the `sarb-sla-report` Copilot CLI skill or directly (`powershell -ExecutionPolicy Bypass -File .\scripts\run-sarb-sla.ps1`). It uses `Set-StrictMode -Version Latest` and `$ErrorActionPreference = 'Stop'` for fail-fast behavior, and relies entirely on built-in PowerShell cmdlets (`Import-Csv`, `Export-Csv`, `ConvertTo-Html`, `Group-Object`, `Add-Member`) — no external modules or packages are required.
- **Shared path resolution:** The script dot-sources `..\..\_shared\SA-Paths.ps1`, which exposes `Get-SALatestInput` and `Get-SAOutputDir`. If `-Source` is not supplied on the command line, `Get-SALatestInput -Patterns '*.csv'` scans the shared `SA\Inputs` folder and selects the file with the most recent `LastWriteTime`; the script throws if none is found. If `-OutputDirectory` is not supplied, `Get-SAOutputDir -SkillName 'sarb-sla-report'` resolves (and creates if necessary) `SA\Outputs\sarb-sla-report`.
- **Scope filtering:** All rows are read via `Import-Csv`. Converted Ideas are identified by collecting every `Idea ID` value referenced by a `PRJ` row (`$convertedIdeaIds`), and any `ID`-prefixed row whose `Work Ref` appears in that set is dropped. The remaining filter keeps rows where `Created` falls in calendar year 2026, `Work Ref` starts with `PRJ` or is a standalone `ID`, and `Status` is not `Cancelled`.
- **SLA date math (`Get-ExpectedSarbDate`):** Adds 28 days to `Created` for `ID`-prefixed work refs, or 14 days for everything else, then advances the date one day at a time until `DayOfWeek -eq 'Tuesday'`.
- **Status classification:** If `SARB Approval Date` is present, the record is `Met` when the approval is on or before the expected date, else `Not met`. If no approval date is present, the record is `Not met` when `$AsOfDate` (defaults to today, overridable via `-AsOfDate`) is after the expected date, otherwise `Not yet due`.
- **Aggregation:** `Group-Object 'Work Ref: Solution Architect'` produces per-SA counts of Met/Not met/Not yet due, with `SLA %` computed as `Met / (Met + Not met)` rounded to one decimal (blank when the denominator is zero), sorted descending by `SLA %`.
- **HTML report generation:** The script hand-builds a self-contained HTML file (no external chart library) using inline CSS and a hand-rolled flex-based bar chart, where each Solution Architect's `Met` and `Not met` bar heights are scaled as a percentage of the largest count across all SAs (`$maxCount`), plus a `ConvertTo-Html -Fragment` table for the full SA performance table.
- **Outputs:** `New-Item -ItemType Directory -Force` ensures the output directory exists, then `analyzed-2026-prj.csv`, `summary-2026-prj.csv`, and `SARB-SLA-report.html` are written via `Export-Csv`/`Set-Content` into `SA\Outputs\sarb-sla-report`.
- **Edge cases handled:** blank `Work Ref: Solution Architect` values are relabeled `(Unassigned)` in the summary rather than dropped; a zero `Met + Not met` denominator yields a blank `SLA %` instead of a divide-by-zero error; the `$maxCount` chart scale is floored at 1 to avoid a zero-height bar chart when no data exists.

## Inputs

- A CSV export of the Solution Architecture Work Queue SharePoint list (explicit `-Source` path, or auto-detected as the newest `*.csv` file in the shared `SA\Inputs` folder). Required columns:
  - `Created` — record creation date, used as the SLA clock start.
  - `Work Ref` — identifier prefix (`PRJ` or `ID`) that determines record type and SLA window.
  - `Idea ID` — used to detect Ideas that have already been converted into a Project.
  - `Status` — used to exclude `Cancelled` records.
  - `SARB Approval Date` — actual approval date, if any.
  - `Work Ref: Solution Architect` — the assigned SA, used for grouping.
  - Other columns may be present in the export; the SLA report does not use them.
- Optional `-AsOfDate` (defaults to today) for historical/point-in-time reporting.
- Optional `-OutputDirectory` (defaults to the shared `SA\Outputs\sarb-sla-report` folder).

## Outputs

All written to `SA\Outputs\sarb-sla-report\` (or a custom `-OutputDirectory`):

- `analyzed-2026-prj.csv` — every in-scope record with its computed `Expected SARB Date` and `SLA Status` (Met / Not met / Not yet due).
- `summary-2026-prj.csv` — one row per Solution Architect with Total, Met, Not met, Not yet due, and `SLA %`.
- `SARB-SLA-report.html` — a self-contained dashboard with summary cards, a horizontal Met/Not-met bar chart by Solution Architect, and the full SA performance table.
