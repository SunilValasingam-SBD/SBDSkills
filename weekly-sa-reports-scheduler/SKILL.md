---
name: weekly-sa-reports-scheduler
description: 'Set up (or change) a weekly Windows scheduled task that automatically runs the techpass-aging, sarb-sla-report, sarb-sla-reasoning, and sa-kpi-heatmap skills against the newest files dropped in a shared inputs folder, writing all outputs to a dated folder. Use when the user wants these reports to run automatically on a recurring schedule (e.g. every Wednesday/Thursday morning) instead of being run manually each time.'
argument-hint: '[reports root folder] [day of week] [time HH:MM]'
user-invocable: true
disable-model-invocation: false
---

# Weekly SA Reports Scheduler

Automates a weekly batch of four reports — TechPass Aging, SARB SLA Report,
SARB SLA Reasoning, and SA KPI Heatmap (SA Performance) — by registering a
native Windows Task Scheduler job that runs them unattended and writes all
outputs to a dated folder.

## Use When

- The user wants to stop manually re-running `techpass-aging`,
  `sarb-sla-report`, `sarb-sla-reasoning`, and `sa-kpi-heatmap` every week.
- The user wants one folder to drop the latest Work Queue / TechPass exports
  into, instead of hunting for file paths each time.
- The user wants to change the day or time of an already-scheduled run.

## How It Works

- `inputs\work-queue\` — drop the latest SARB Work Queue export here (CSV or
  Excel). Used by SARB SLA Report, SARB SLA Reasoning, and SA KPI Heatmap.
- `inputs\techpass\` — drop the latest TechPass/`sc_task` export here (Excel).
  Used by TechPass Aging Report.
- Old files can stay in those folders for history — the runner always picks
  the **most recently modified** file in each folder, and logs a warning if
  the newest file is more than 8 days old (likely stale).
- `outputs\<yyyy-MM-dd>\` — each run creates a dated folder with four
  subfolders (`techpass-aging`, `sarb-sla-report`, `sarb-sla-reasoning`,
  `sa-performance`) containing that skill's normal output files.
- `logs\run-<date>.log` — a log of what ran, which input files were used, and
  any warnings or errors.
- A Windows Task Scheduler task (native OS scheduler, independent of any
  Copilot CLI session) runs [Run-WeeklyReports.ps1](./scripts/Run-WeeklyReports.ps1)
  on the configured day/time. **Important:** the PC must be on and the user
  logged in at that time for a standard scheduled task to fire.

## Procedure

1. Confirm the Work Queue and TechPass export files/paths exist, or ask the
   user to export them first.
2. Ask the user (if not already known from context):
   - Reports root folder (default: reuse an existing one if the user has run
     this before, e.g. `C:\Users\<user>\WeeklyReports`).
   - Day of week (MON/TUE/WED/THU/FRI/SAT/SUN).
   - Time of day (24h `HH:MM`).
3. Run the setup script — it is idempotent, so re-running it with new
   day/time values simply replaces the existing scheduled task:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\Setup-WeeklySchedule.ps1 `
       -ReportsRoot "<reports root folder>" -DayOfWeek THU -Time 08:15
   ```

   Add `-RunNow` to validate immediately against whatever is currently in the
   inputs folders:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\Setup-WeeklySchedule.ps1 `
       -ReportsRoot "<reports root folder>" -DayOfWeek THU -Time 08:15 -RunNow
   ```

4. If the user wants to run the batch manually/off-schedule at any time
   (without touching the scheduled task), run the runner directly:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\Run-WeeklyReports.ps1 `
       -ReportsRoot "<reports root folder>"
   ```

   Optional: `-AsOfDate YYYY-MM-DD` for a historical snapshot run,
   `-OpenOutputs $false` to skip auto-opening the output folder (recommended
   when running non-interactively).
5. Report back: the task name, next run time, and the reports-root folder
   paths the user should drop weekly exports into.
6. To change the schedule later (different day/time/folder), just re-run step
   3 with the new values — it deletes and recreates the task.
7. To remove the automation entirely:

   ```powershell
   schtasks /Delete /TN "WeeklyReports-SARB-TechPass-SAPerformance" /F
   ```

## Requirements

Same Python/package requirements as the four underlying skills
(`techpass-aging`, `sarb-sla-report`, `sarb-sla-reasoning`, `sa-kpi-heatmap`):
Python 3.10+, `pandas`, `openpyxl`, `python-pptx`, `plotly`. This skill does
not duplicate their logic — it only orchestrates when/how they run and where
inputs/outputs live.
