---
name: sa-engagement-analysis
description: 'Analyze Solution Architect engagement in initiatives from a CSV, comparing end-to-end idea-to-project engagement with project-only engagement by portfolio and SA, using an interactive dashboard and CSV summaries.'
argument-hint: 'Provide the CSV path when prompted; optionally provide the initiative creation year'
user-invocable: true
disable-model-invocation: false
---

# SA Engagement Analysis

Build an interactive engagement dashboard from a CSV export containing Solution Architect workload for ideas and projects.

## Business definitions

- **Idea** — `Work Ref` starts with `ID` (case-insensitive).
- **Project** — `Work Ref` starts with `PRJ` (case-insensitive).
- **End-to-end (E2E) engagement** — a project row (`Work Ref` starts with `PRJ`) with a populated `Idea ID`, meaning the SA was engaged before the idea became a project.
- **Project-only engagement** — a project row with a blank `Idea ID`.
- **Total engagement** — all included idea and project records. Counts use distinct `Work Ref` values within each grouping where possible.

The default reporting period is the current calendar year. The creation date is detected from common column names (`Created`, `Created Date`, `Initiative Created`, `Created On`, or `Date Created`). If no unambiguous date column is found, the generator stops with an actionable error; pass `--created-column` to identify it explicitly.

## Required columns

- `Work Ref`
- `Idea ID`
- `Program Owning Portfolio`
- A creation-date column, unless supplied with `--created-column`

An SA column is detected from `Work Ref: Solution Architect`, `Solution Architect`, `SA`, or `Assigned to`. Use `--sa-column` when the export uses another name.

## Procedure

1. Prompt for the CSV path if the user has not supplied one. Do not assume a filename.
2. Ask for or use the current calendar year unless the user requests another year.
3. Validate the required columns and parse the creation date.
4. Run:

   ```powershell
   py .\sa_engagement_analysis.py --input "<CSV path>" --year 2026
   ```

   The script creates:

   - `sa_engagement_dashboard.html` — interactive chart with a Portfolio/SA view selector, hover details, and E2E percentage.
   - `sa_engagement_summary.csv` — one row per Portfolio and one row per SA.
   - `sa_engagement_detail.csv` — normalized included records with classification fields.

5. Report the included record count, date range, detected columns, and output paths.

## Dashboard interpretation

For each Portfolio or SA, the dashboard shows total engagement, ideas, projects, E2E projects, and the E2E percentage of total engagement. Use the dropdown to switch from the portfolio view to the SA view. Hover over a bar for counts and percentages.

## Rules

- Only rows whose `Work Ref` starts with `ID` or `PRJ` are included.
- Blank `Idea ID` values include empty strings, whitespace, `NaN`, `None`, and `null`.
- A project with a populated `Idea ID` is E2E; an idea row is not counted as E2E until it has a project `Work Ref`.
- Blank Portfolio or SA values are grouped as `Unassigned`.
- The source CSV is never modified.

The bundled generator requires Python 3.10+ and `pandas` and `plotly`. Install them with:

```powershell
py -m pip install -r requirements.txt
```
