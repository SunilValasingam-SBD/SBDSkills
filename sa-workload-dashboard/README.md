# SA Workload Dashboard (Worklist vs. Timesheet)

## Executive Summary

This tool reconciles two things that normally live in separate systems and
never get compared: what Solution Architects say they are working on (the
Worklist) and what hours they actually log against each type of work (the
Timesheet). It reveals a blind spot — Ideas and TechPasses have no dedicated
timesheet code, so that real effort is normally buried and invisible inside
generic "Admin" hours. By estimating that hidden effort from the Worklist and
pulling it back out of Admin, leadership gets a much more honest, single-page
picture of how each architect's time is truly allocated across delivery work,
administrative overhead, and training — supporting fairer workload and
staffing conversations.

## Business Process Flow

1. A business user provides two source files every run: a Worklist export
   (CSV) and a Timesheet export (Excel) — both must be current, since the
   tool never assumes or reuses a stale file.
2. The user can optionally tune a few settings (which architects to exclude,
   how far back to look for recently closed items, and how many hours/week
   to assume an Idea or TechPass consumes); sensible defaults are applied
   automatically otherwise.
3. The tool evaluates each architect's Worklist items (counts by type:
   Project, Idea, TechPass, SA NPD, NPD) and separately totals their logged
   Timesheet hours by type (Project, Idea, Time Off, Training, Admin/Other).
4. Because Idea and TechPass hours are not separately tracked in the
   Timesheet system, the tool estimates how much time was likely spent on
   each open Idea/TechPass (based on how many weeks it has been active) and
   removes that estimated amount from the architect's Admin/Other bucket so
   it isn't hidden there — revealing a more realistic time allocation.
5. The deliverable — a single interactive dashboard web page with three
   side-by-side panels — is written to the shared
   `SA\Outputs\sa-workload-dashboard` folder.
6. A reviewer opens the dashboard to compare, per architect: how many items
   of each type they're carrying, how many hours they logged by type, and
   (in the third panel) a truer picture of time allocation once Idea/TechPass
   effort is accounted for — supporting workload-balancing and capacity
   discussions.

## Technical Flow

- Single script `sa_workload_dashboard.py`; uses **openpyxl** to read the
  Timesheet `.xlsx`, the standard library **csv** module to read the
  Worklist CSV, and hand-rolled HTML/JS (Chart.js + chartjs-plugin-datalabels
  via CDN) to render the three-panel dashboard — no pandas/plotly dependency.
- **Input resolution**: this skill requires **two explicit inputs**
  (`--worklist`, `--timesheet`) since auto-detecting between two ambiguous
  files in the shared `SA\Inputs` folder would be unsafe; if either flag is
  omitted the script prompts interactively for the path. `--output-dir`
  defaults to the shared `SA\Outputs\sa-workload-dashboard` folder via
  `sa_paths.output_dir_for("sa-workload-dashboard")`.
- **Worklist classification (`load_worklist`)**: the leading alphabetic
  prefix of `Work Ref` is matched against `WORKREF_PREFIX_MAP`
  (`PRJ`→Project, `ID`→Idea, `TECHPASS`→TechPass, `SA`→SA NPD, `NPD`→NPD;
  unmatched prefixes become their own category rather than being dropped).
  Rows are kept only if `Status == "WIP"`, or `Status == "Complete"` and the
  item's closure date (`Work Ref: Finish Date`, falling back to `Modified`)
  is on/after a cutoff date (`--lookback-days`, default 28 days, via
  `is_recently_active`). This recency filter affects Panel 1 only.
- **Timesheet classification (`load_timesheet`)**: the workbook's header row
  is located by scanning the first 10 rows for one whose first cell equals
  `"Segment Name"` (handles preamble rows above the real header). Rows with
  `Project ID` equal to the user-supplied `--timeoff-project-id` are
  reclassified by `Timesheet Line Class` text (`classify_timeclass`:
  contains "time off" → Time Off, contains "training" → Training &
  Development, else → Admin/Other); all other rows are classified Project/Idea
  by the `Project ID` prefix. Distinct `Week Start` values across all rows are
  collected to define the imputation period.
- **Imputation (`load_impute_items` + `compute_imputed_hours`)**: every
  Idea/TechPass Worklist row with a parseable `Created` date is treated as
  "active" starting that date; if its `Status` is terminal (`Complete` or
  `Cancelled`), it stops being active after its `Modified` date, otherwise it
  stays active through the end of the timesheet's week range. For every week
  in the timesheet's observed week set where the item is active, it accrues
  `--idea-hours-per-week` (default 1.0) or `--techpass-hours-per-week`
  (default 2.0) hours. `_parse_date` tries several formats
  (`%m/%d/%Y %I:%M %p`, `%m/%d/%Y`, `%d-%b-%y`, `%d-%b-%Y`).
- **Panel 3 assembly (`main`)**: for each SA, Panel 3 = actual Project hours
  (from Timesheet) + imputed Idea hours + imputed TechPass hours + Admin/Other
  (netted of the imputed Idea/TechPass hours, floored at 0, unless
  `--no-net-admin` is passed) + Training & Development. Time Off is excluded
  entirely from Panel 3. `build_dataset_with_pct` attaches each segment's % of
  that SA's Panel-3 total for on-bar/tooltip labeling.
- **Rendering**: `build_html` embeds the three datasets as JSON and renders
  three horizontal stacked Chart.js bar charts; a `contrastColor` helper
  picks readable (white/black) label text and a matching outline color per
  bar segment based on its background luminance.
- **Output**: a single self-contained HTML file, `sa_dashboard.html`, written
  via `Path.write_text` to `--output-dir` (default
  `SA\Outputs\sa-workload-dashboard\`).
- **Edge cases handled**: SAs present in only one source still appear across
  all panels with zero/blank values on the side they're missing; unrecognized
  `Work Ref`/`Project ID` prefixes become their own visible category instead
  of being silently dropped; imputation is skipped for items/SAs lacking a
  usable `Created` date or when the timesheet has no `Week Start` values at
  all.

## Inputs

- **Worklist CSV export** (`--worklist`, required) — must contain:
  `Work Ref`, `Work Ref: Solution Architect`, `Status`,
  `Work Ref: Finish Date`, `Modified`, `Created` (used for imputation).
- **Timesheet Excel export** (`--timesheet`, required) — must contain (after
  its header row, which may follow a couple of preamble rows): `Resource
  Name`, `Project ID`, `Timesheet Line Class`, `Hours Reported`; ideally also
  `Week Start` (used to define the imputation period for Panel 3).
- Optional flags: `--exclude-sa` (repeatable, exact "Last, First" match;
  no exclusions by default), `--lookback-days` (default 28),
  `--timeoff-project-id` (organization-specific; no default), `--idea-hours-per-week` (default
  1.0), `--techpass-hours-per-week` (default 2.0), `--no-net-admin`,
  `--output-dir` (default shared `SA\Outputs\sa-workload-dashboard`).

## Outputs

- `sa_dashboard.html` — interactive three-panel dashboard: Worklist item
  counts by classification, Timesheet hours by classification, and an
  actual-plus-imputed time allocation view with % labels, written to
  `SA\Outputs\sa-workload-dashboard\` by default.
