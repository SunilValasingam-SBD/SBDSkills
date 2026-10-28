---
name: sa-workload-dashboard
description: 'Generate a dashboard comparing, per Solution Architect, the count of Worklist items by classification (Project/Idea/TechPass/SA NPD/NPD), total hours spent by classification (Project/Idea/Time Off/Training/Admin-Other) from a timesheet export, and a third panel that imputes actual Idea/TechPass effort (since they share no dedicated timesheet Project ID) to show true time allocation with %. Use for SA workload comparisons, worklist-vs-timesheet analysis, or time allocation by classification.'
argument-hint: 'Provide the Worklist CSV and Timesheet Excel source paths when prompted'
user-invocable: true
disable-model-invocation: false
---

# SA Workload Dashboard (Worklist vs. Timesheet)

Build an interactive three-panel HTML dashboard per Solution Architect (SA):

- **Panel 1 (left)** — count of Worklist items by classification (Project, Idea, TechPass, SA NPD, NPD)
- **Panel 2 (right)** — total hours spent by classification (Project, Idea, Time Off, Training & Development, Admin/Other), from a timesheet export
- **Panel 3 (bottom, full-width)** — actual + imputed time allocation by classification (Project, Idea, TechPass, Admin/Other, Training & Development), with each bar segment labeled with both hours and % of that SA's total. Time Off is excluded entirely (not "work" time).

**Why Panel 3 exists:** Ideas and TechPasses may have no dedicated Project ID in a timesheet system. If that time is logged under an administrative project, it is indistinguishable from administrative work and time off in Panel 2. Panel 3 estimates Idea/TechPass effort from the Worklist instead:

- Every `ID`-prefixed item accrues **1 hour/week** (configurable) for every week it was active.
- Every `TechPASS`-prefixed item accrues **2 hours/week** (configurable) for every week it was active.
- "Weeks" = the distinct `Week Start` values actually present in the Timesheet export (so the imputed period matches exactly the period real hours were logged for).
- An item is "active" in a given week if its `Created` date (assignment-date proxy) is on/before that week's start, **and** — if the item has reached a terminal status (`Complete`/`Cancelled`) — its `Modified` date (completion/cancellation-date proxy) is on/after that week's start. Items not yet Complete/Cancelled are treated as active through the end of the available week range.
- The imputed Idea/TechPass hours are **subtracted from that SA's Admin/Other total** (floored at 0), since that time is already counted there today and would otherwise be double-counted.

All three bar charts are horizontal, stacked, color coded with a distinct high-contrast color per category, and show the value printed on each bar segment (auto white/black text + outline so numbers stay readable on any bar color). Panel 3 additionally shows each segment's % of that SA's total allocated time (in the label and the tooltip).

## Required Inputs (always ask for both, every run)

This skill requires **two source files every run** — never assume a filename or reuse a prior run's path without confirming it's still current:

1. **Worklist CSV export** — must contain: `Work Ref`, `Work Ref: Solution Architect`, `Status`, `Work Ref: Finish Date`, `Modified`.
2. **Timesheet Excel export** — must contain (after its header row, which may follow a couple of preamble rows): `Resource Name`, `Project ID`, `Timesheet Line Class`, `Hours Reported`.

If the user hasn't provided both paths, ask for them before running.

## Classification Rules

**Worklist** (derived from the `Work Ref` column prefix):
- `PRJ` → Project
- `ID` → Idea
- `TechPASS` → TechPass
- `SA` → SA NPD
- `NPD` → NPD

**Timesheet** (derived from `Project ID`):
- If `Project ID` equals the configured time-off project id, classify by the `Timesheet Line Class` column: contains "time off" → Time Off; contains "training" → Training & Development; anything else → Admin / Other.
- Otherwise classify by `Project ID` prefix: `PRJ` → Project, `ID` → Idea.

## Default Filters (apply unless the user says otherwise for a given run)

- All Solution Architects are included by default. Use `--exclude-sa` only when a specific exclusion is intended.
- **Worklist recency**: include items that are `Status = WIP`, or `Status = Complete` closed within the last 28 days (using `Work Ref: Finish Date`, falling back to `Modified` when Finish Date is blank).

## Procedure

1. Ask the user for the Worklist CSV path and the Timesheet Excel path (every run — do not reuse stale paths without confirming).
2. Run the bundled generator:

   ```powershell
   py .\sa_workload_dashboard.py --worklist "<worklist csv path>" --timesheet "<timesheet xlsx path>"
   ```

   If `--worklist`/`--timesheet` are omitted, the script prompts interactively for the paths.

   Override defaults only if the user explicitly asks for different filters:

   - `--exclude-sa "<Last, First>"` — exclude a specific SA (exact match; repeatable; omitted means no exclusions).
   - `--lookback-days <n>` — change the recency window for closed Worklist items (default 28). Only affects Panel 1.
   - `--timeoff-project-id "<id>"` — set the organization-specific Project ID that represents time off/training/admin rows (no default).
   - `--idea-hours-per-week <n>` — imputed hours/week per active Idea in Panel 3 (default 1.0).
   - `--techpass-hours-per-week <n>` — imputed hours/week per active TechPass in Panel 3 (default 2.0).
   - `--no-net-admin` — do NOT subtract imputed Idea/TechPass hours from Admin/Other in Panel 3 (default: subtract, to avoid double-counting).
   - `--output-dir "<path>"` — change where `sa_dashboard.html` is written (default: this skill's folder).

3. Open `sa_dashboard.html` to review all three panels.

## Rules

- Source files are never modified.
- Blank/unrecognized `Work Ref` or `Project ID` prefixes are kept as their own "Unclassified"/raw-prefix category rather than silently dropped, so data issues stay visible.
- SAs appearing in only one of the two sources still appear in all three panels (with zero/blank bars on the side where they have no data).
- Panel 3's Idea/TechPass imputation requires a `Created` date on the Worklist row and at least one `Week Start` date in the Timesheet; items or SAs missing these are simply excluded from the imputed total (not from Panels 1/2).
