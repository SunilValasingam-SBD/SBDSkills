---
name: sa-kpi-heatmap
description: 'Generate an interactive KPI performance heatmap per Solution Architect (TechPass Created, Scenario Published, SARB Approved, SAD Approved) with click-to-drill-down project lists. Use for Solution Architect KPI scorecards, performance heatmaps, or project-level drill-down by KPI.'
argument-hint: 'Provide the CSV/Excel source path when prompted'
user-invocable: true
disable-model-invocation: false
---

# Solution Architect KPI Heatmap

Build an interactive heatmap of four KPIs per Solution Architect (SA), color coded by performance band, with click-to-drill-down into the contributing projects.

## KPIs

1. **TechPass Created** — `TechPass` column is populated
2. **Scenario Published** — `Publication Status` is populated and not "Not Published"
3. **SARB Approved** — `SARB Approval Date` is populated
4. **SAD Approved** — `SAD Approved` column is populated

Each KPI is the percentage of that SA's total (non-cancelled, or filtered) projects meeting the rule above.

Only projects/ideas are considered — `Work Ref` must start with `PRJ` or `ID`.

## Color Bands

- **Dark green**: 70% and above
- **Light green**: 50% – 70%
- **Yellow**: 30% – 50%
- **Light orange**: below 30%

## Default Filters

Unless the user says otherwise for a given run, always apply these defaults:

- Exclude unassigned records: `Unassigned` (substring, case-insensitive). Add any other exclusions only when requested for that run.
- Restrict to `Status` = `WIP` only.

## Procedure

1. Ask the user for the source file path (CSV or Excel export) if not already provided — do not assume a filename.
2. Confirm the file contains these columns: `Work Ref`, `Work Ref: Name`, `Work Ref: Solution Architect`, `Status`, `TechPass`, `Publication Status`, `SARB Approval Date`, `SAD Approved`.
3. Run the bundled generator, applying the [Default Filters](#default-filters) above:

   ```powershell
   py .\sa_kpi_heatmap.py --input "<source file path>" --exclude-sa Unassigned --status WIP
   ```

   If `--input` is omitted, the script prompts interactively for the path.

   Override the defaults only if the user explicitly asks for different filters:

   - `--exclude-sa "<name>"` — exclude a Solution Architect (substring, case-insensitive match; repeatable). Use to drop people who left the org.
   - `--status "<status>"` — restrict to one or more exact `Status` values (repeatable). If omitted entirely (no `--status` at all), all non-cancelled projects are included instead of just `WIP`.

4. Open `sa_kpi_heatmap.html` — click any heatmap cell to see two lists: projects that **met** the KPI and projects that are **missing**/did not meet it (Project ID, Project Name, Status) for that SA/KPI.
5. Review `sa_kpi_summary.csv` for the per-SA percentage and met/total counts for all four KPIs.

## Rules

- Only Projects and Ideas are considered — `Work Ref` must start with `PRJ` or `ID` (case-insensitive); other work types are excluded.
- By default, projects with `Status` = `Cancelled` are excluded from all totals; use `--status` to instead restrict to specific status value(s) such as `WIP`.
- `Work Ref: Solution Architect` is used as the SA label (blank values are grouped as `Unassigned`).
- `Work Ref` is used as the Project ID shown in the drill-down table.
- The source file is never modified.

The bundled generator requires Python 3.10+ and the `pandas`, `plotly`, and `openpyxl` packages. Install them with `py -m pip install -r requirements.txt`.
