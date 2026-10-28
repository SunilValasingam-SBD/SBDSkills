# Weekly SA Reports Scheduler

## Executive Summary

Four of the team's most valuable recurring reports — TechPass aging, SARB SLA performance, SLA-miss reasoning, and the SA KPI heatmap — currently depend on someone remembering to re-run them every week with the latest data. This skill removes that manual burden by registering a native, unattended Windows schedule that automatically regenerates all four reports on a chosen day and time, pulling whatever the latest Work Queue and TechPass exports are, and filing the results into a dated folder. Leadership gets a current, consistent set of SA performance and SLA reports delivered on a predictable cadence, with no one needing to remember to run anything, and a clear log of exactly what data each week's reports were built from.

## Business Process Flow

1. The team drops the latest weekly exports — the SARB Work Queue export and the TechPass export — into two designated shared drop folders. Older exports can be left in place for history; only the most recently modified file in each folder is used.
2. The team (once, during setup) tells the tool which folder to operate from, and what day of the week and time the reports should run.
3. The tool registers a recurring job with the Windows operating system itself — independent of any particular chat session — so the four reports regenerate automatically at the chosen time, as long as the computer is on and the user is logged in.
4. Each time the job runs, it automatically picks the newest Work Queue and TechPass files from the drop folders, regenerates all four reports, and warns (in a log file) if the newest available file looks stale (more than 8 days old) so the team knows to refresh it.
5. All four reports for that run are filed together into a single folder labeled with that run's date, so a given week's full report set is always easy to find as one bundle.
6. The schedule, day, time, or drop-folder location can be changed at any time by simply re-running the setup step with new values — it automatically replaces the old schedule with the new one. The automation can also be removed entirely on request.
7. The team can also trigger an immediate, off-schedule run at any time (for example right after dropping in a fresh export) without disturbing the recurring schedule.

## Technical Flow

- **Setup**: `scripts\Setup-WeeklySchedule.ps1 -ReportsRoot "<folder>" -DayOfWeek <DOW> -Time <HH:MM> [-RunNow]` creates the folder structure under `<ReportsRoot>` (`inputs\work-queue\`, `inputs\techpass\`, `outputs\`, `logs\`) and registers/replaces a Windows Task Scheduler task (`schtasks /Create /SC WEEKLY ...`) named `WeeklyReports-SARB-TechPass-SAPerformance` that runs `Run-WeeklyReports.ps1 -ReportsRoot "<folder>"` on the configured day/time. The script is idempotent: re-running it with new values deletes the existing task and recreates it.
- **Recommended location**: because this skill's own `inputs\`/`outputs\` subfolders are created directly under whatever `-ReportsRoot` is given, and Windows paths are not case-sensitive, pointing `-ReportsRoot` at the shared SA root folder (the parent of `SA\Skills\`, i.e. `...\SA\`) makes `<ReportsRoot>\inputs\work-queue` and `<ReportsRoot>\inputs\techpass` land inside the same shared `SA\Inputs\` folder the other skills already drop files into, and `<ReportsRoot>\outputs\<date>\` land inside the same shared `SA\Outputs\` folder the other skills already write to (as dated subfolders alongside each skill's own named output folder). This keeps one shared Inputs/Outputs location for the whole SA toolset instead of a separate reports-only folder tree.
- **Runner**: `Run-WeeklyReports.ps1 -ReportsRoot "<folder>" [-AsOfDate yyyy-MM-dd] [-OpenOutputs $true|$false]` creates `outputs\<AsOfDate>\` with four subfolders (`techpass-aging`, `sarb-sla-report`, `sarb-sla-reasoning`, `sa-performance`) and a `logs\run-<AsOfDate>.log`. It independently selects the most recently modified file in `inputs\work-queue\` and `inputs\techpass\` (skipping Office lock files matching `^~\$`), logging a warning if the selected file is more than 8 days old.
- It then shells out to each underlying skill directly, writing into that run's dated subfolders:
  - `techpass_aging_chart.py --input <techpass file> --as-of <date> --output ...\techpass_aging_by_architect.html --pptx-output ...\techpass_aging_summary.pptx` (skipped if no TechPass file is found).
  - If the Work Queue file is Excel, it is first converted to CSV via a `pandas`/`py -c` one-liner (because the SARB SLA report runner reads with `Import-Csv`), then `sarb-sla-report\scripts\run-sarb-sla.ps1 -Source <csv> -AsOfDate <date> -OutputDirectory ...\sarb-sla-report` is invoked.
  - `sarb_sla_reasoning.py --input <work queue file> --as-of <date> --output-dir ...\sarb-sla-reasoning --output ...\sarb-sla-reasoning.html`.
  - `sa_kpi_heatmap.py --input <work queue file> --output ...\sa_kpi_heatmap.html --summary-output ...\sa_kpi_summary.csv`.
- All four calls are wrapped in try/catch and logged (`Write-Log`) so one report failing does not stop the others from running; the completed run folder is optionally auto-opened (`Invoke-Item`), which should be disabled (`-OpenOutputs $false`) for unattended/non-interactive scheduled runs.
- This skill contains no report-generation logic itself — it only resolves which input files to use, orchestrates the four underlying skills' existing scripts in sequence, and manages the Windows Task Scheduler registration.

## Inputs

- **Work Queue export** (CSV or Excel) dropped in `inputs\work-queue\` — used by SARB SLA Report, SARB SLA Reasoning, and SA KPI Heatmap.
- **TechPass/`sc_task` export** (Excel) dropped in `inputs\techpass\` — used by TechPass Aging Report.
- **`-ReportsRoot`** — the root folder containing the `inputs\`/`outputs\`/`logs\` structure (recommended: the shared `SA` root folder so these land alongside the shared `Inputs\`/`Outputs\` folders).
- **`-DayOfWeek`** / **`-Time`** — the weekly schedule for the Windows scheduled task (setup only).
- Optional: **`-AsOfDate`** for a historical snapshot run; **`-OpenOutputs`** to control auto-opening the output folder; **`-RunNow`** to validate immediately after setup.

## Outputs

- **`outputs\<yyyy-MM-dd>\techpass-aging\`** — `techpass_aging_by_architect.html` and `techpass_aging_summary.pptx` for that run's date.
- **`outputs\<yyyy-MM-dd>\sarb-sla-report\`** — that run's SARB SLA Report output files.
- **`outputs\<yyyy-MM-dd>\sarb-sla-reasoning\`** — `sarb-sla-reasoning.html` and supporting files for that run's date.
- **`outputs\<yyyy-MM-dd>\sa-performance\`** — `sa_kpi_heatmap.html` and `sa_kpi_summary.csv` for that run's date.
- **`outputs\<yyyy-MM-dd>\work-queue-source.csv`** — intermediate CSV conversion of the Work Queue Excel export, when applicable.
- **`logs\run-<yyyy-MM-dd>.log`** — a log of which input files were used, any staleness warnings, and any errors from that run.
- **Windows Task Scheduler task** (`WeeklyReports-SARB-TechPass-SAPerformance`) — the registered recurring job itself, visible via `schtasks /Query`.
