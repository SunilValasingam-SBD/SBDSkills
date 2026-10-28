# Solution Architect KPI Heatmap

## Executive Summary

This tool gives leadership an at-a-glance view of how consistently each Solution
Architect is completing the key governance milestones every project is expected
to pass through — standing up a TechPass, publishing a scenario, clearing SARB
approval, and obtaining SAD approval. Instead of manually auditing spreadsheets
project by project, a manager can open one color-coded report and immediately
see which architects are on track (green) and which need coaching or follow-up
(yellow/orange). A companion workbook also gives a simple, print-friendly list
of every work-in-progress project grouped by architect, useful for 1:1s and
capacity conversations. Together these remove hours of manual spreadsheet
filtering from every KPI/status review cycle.

## Business Process Flow

1. A business user exports the Solution Architect project/work list (CSV or
   Excel) from the source system and drops it in the shared `SA\Inputs`
   folder (or has it ready to point to directly).
2. The user runs the skill, optionally telling it to exclude specific
   architects (e.g., people who've left the team) or restrict to a specific
   project status such as "WIP." If nothing is specified, the tool
   automatically picks up the newest matching file in `SA\Inputs` and applies
   the standing defaults (exclude a few known non-architect names, and
   restrict to WIP projects only).
3. The tool evaluates, for every project and idea owned by each architect,
   whether four milestones were reached: TechPass created, scenario
   published, SARB approved, and SAD approved.
4. It calculates each architect's completion percentage for all four
   milestones and assigns a performance color band (dark green, light green,
   yellow, or light orange) so trends are visible instantly.
5. The deliverables — an interactive heatmap webpage and a summary
   spreadsheet — are written to the shared `SA\Outputs\sa-kpi-heatmap` folder.
6. A reviewer opens the heatmap, clicks any colored cell, and instantly sees
   the exact list of projects that did and did not meet that milestone for
   that architect — turning a percentage into an actionable follow-up list.
7. Separately, a reviewer can generate a simple "who's working on what"
   workbook: a list of all open (WIP) projects/ideas, grouped and
   color-banded by architect, plus a one-page count summary per architect —
   useful for workload check-ins without needing the full heatmap.

## Technical Flow

This skill folder contains **two independent scripts** that share the same
input conventions but produce different deliverables:

### `sa_kpi_heatmap.py` (the KPI heatmap)

- Built with **pandas** (data loading/cleaning/aggregation) and **plotly**
  (interactive heatmap rendering); reads `.xlsx`/`.xls` via **openpyxl** or
  `.csv` via `pandas.read_csv`.
- **Input resolution**: `--input` if given; otherwise it calls the shared
  `sa_paths.latest_input("*.csv", "*.xlsx", "*.xls")` helper to pick the
  newest matching file in `SA\Inputs`; if nothing is found there, it falls
  back to an interactive prompt for a path.
- **Validation**: on load, column headers are stripped and the frame is
  checked against `REQUIRED_COLUMNS`; missing columns raise an error before
  any processing happens.
- **Filtering (`clean_projects`)**:
  - `Status` is filtered to the exact value(s) passed via repeatable
    `--status` (case-insensitive); if `--status` is omitted entirely, all
    rows except `Status == "Cancelled"` are kept instead.
  - Only rows whose `Work Ref` (uppercased) starts with `PRJ` or `ID` are
    kept (Projects and Ideas only).
  - `Work Ref: Solution Architect` becomes the `Solution Architect` label;
    blank values are normalized to `"Unassigned"`.
  - Repeatable `--exclude-sa <substring>` performs a case-insensitive
    substring match against the SA label and drops matching rows.
- **KPI computation**: each of the four KPI boolean columns is computed with
  an `is_blank()` helper (handles `None`/NaN/empty-string/"nan"/"none") against
  `TechPass`, `Publication Status` (populated and not equal to
  `"Not Published"`), `SARB Approval Date`, and `SAD Approved`.
- **Aggregation (`build_summary`)**: grouped by `Solution Architect`, each KPI's
  `% = met / total * 100` (rounded to 1 decimal) plus a `"met/total"` string are
  computed and written to one row per architect.
- **Drill-down (`build_drilldown`)**: for every architect × KPI combination, two
  lists (`met`/`missing`) of `{Project ID, Project Name, Status}` are built and
  serialized to JSON, embedded directly in the generated HTML so clicking a
  heatmap cell renders the matching table client-side (no server needed).
- **Rendering (`create_heatmap`)**: a `plotly.graph_objects.Heatmap` is built
  with a custom banded colorscale (`<30%` orange, `30–50%` yellow, `50–70%`
  light green, `≥70%` dark green) and a `plotly_click` JS listener that looks
  up the clicked architect/KPI in the embedded drill-down JSON.
- **Outputs**: `summary.to_csv(...)` writes the CSV summary; the HTML heatmap
  is assembled as a full standalone page (Plotly CDN + inline JS) and written
  with `Path.write_text`. Both default to
  `SA\Outputs\sa-kpi-heatmap\` via `sa_paths.output_dir_for("sa-kpi-heatmap")`,
  overridable with `--output`/`--summary-output`.
- **Edge cases handled**: an empty post-filter frame raises a clear error
  instead of producing a blank report; blank SA values are grouped rather than
  dropped; status filtering and SA exclusion are both case-insensitive.

### `build_wip_by_sa.py` (WIP-by-SA workbook builder)

- Built with the standard library **csv** module (input) and **openpyxl**
  (`Workbook`, `PatternFill`, `Font`, `Alignment`) for Excel output — no
  pandas/plotly dependency.
- **Input resolution**: `--input` if given; otherwise `sa_paths.latest_input("*.csv")`
  picks the newest `.csv` in `SA\Inputs`; if none found, the script exits
  with an error (no interactive prompt for this script).
- **Filtering (`filter_wip`)**: keeps only rows where `Status` (case-insensitive)
  equals exactly `"WIP"` and `Work Ref` starts with `PRJ` or `ID`.
- **Sorting/banding**: rows are sorted by Solution Architect (blank → `"Unassigned"`)
  then by `Work Ref: Start Date` string; each time the SA changes, the row
  fill color alternates between two bands (`BAND_COLORS`) so groups are
  visually distinguishable without gridlines.
- **Output construction**: writes a two-tab workbook — `"WIP Projects"` (one
  row per project with SA, Work Ref, Project Name, Program Owning Portfolio,
  BRM, Start Date, Finish Date; styled header row, frozen header, set column
  widths) and `"Summary"` (count of WIP items per SA plus a bold Total row).
- **Output path**: defaults to
  `SA\Outputs\sa-kpi-heatmap\wip_projects_by_sa.xlsx` via
  `sa_paths.output_dir_for("sa-kpi-heatmap")`, overridable with `--output`.

## Inputs

**`sa_kpi_heatmap.py`** — one CSV or Excel export (`--input`, auto-detects
newest `*.csv`/`*.xlsx`/`*.xls` in `SA\Inputs` when omitted) containing:

- `Work Ref` — Project/Idea ID; must start with `PRJ` or `ID` to be included.
- `Work Ref: Name` — project/idea name.
- `Work Ref: Solution Architect` — SA owner (blank → "Unassigned").
- `Status` — used for the Cancelled exclusion / `--status` filter.
- `TechPass` — populated when a TechPass exists.
- `Publication Status` — scenario publication state.
- `SARB Approval Date` — populated when SARB-approved.
- `SAD Approved` — populated when SAD-approved.

Optional flags: `--exclude-sa <name>` (repeatable, substring match),
`--status <value>` (repeatable), `--output`, `--summary-output`.

**`build_wip_by_sa.py`** — one raw work list CSV (`--input`, auto-detects
newest `*.csv` in `SA\Inputs` when omitted) containing: `Status`, `Work Ref`,
`Work Ref: Solution Architect`, `Work Ref: Name`,
`Work Ref: Program Owning Portfolio`, `Work Ref: BRM`,
`Work Ref: Start Date`, `Work Ref: Finish Date`. Optional: `--output`.

## Outputs

Written by default to `SA\Outputs\sa-kpi-heatmap\`:

- `sa_kpi_heatmap.html` — interactive click-to-drill-down KPI heatmap per
  Solution Architect.
- `sa_kpi_summary.csv` — per-SA percentage and met/total counts for all four
  KPIs.
- `wip_projects_by_sa.xlsx` — color-banded workbook of all WIP projects/ideas
  grouped by Solution Architect, plus a per-SA count summary tab.
