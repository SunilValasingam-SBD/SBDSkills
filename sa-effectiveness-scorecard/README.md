# SA Effectiveness Scorecard

## Executive Summary

This tool answers a question leadership cares about more than raw hours or
ticket counts: how much of each Solution Architect's time is genuinely
spent on tangible, delivery-facing work that creates business value, versus
legitimate internal initiatives, versus administrative overhead. It produces
a ranked leaderboard and scorecard so the organization can recognize
architects doing real delivery work rather than simply staying busy with
tickets or admin tasks, and can have more informed conversations about
staffing, recognition, and workload balance. It builds on the same
Worklist-and-Timesheet reconciliation used for workload reporting, but adds
an "Effectiveness Score" that converts the raw numbers into a single,
comparable ranking metric per architect.

## Business Process Flow

1. A business user provides a current Worklist export (CSV) and a current
   Timesheet export (Excel) — both are required every run, and the tool
   never silently reuses a stale file from a previous run.
2. The user may optionally adjust a few assumptions (which architects to
   exclude, and how many hours per week to credit to Idea/TechPass/SA
   NPD/NPD items that have no dedicated timesheet code); defaults are
   applied automatically otherwise.
3. The tool classifies every Worklist item into one of three business
   tiers: **Tangible/Delivery** work (Projects, Ideas, TechPasses), legitimate
   **Internal/NPD** work (SA NPD and NPD initiatives, tracked separately so
   it's never mistaken for either productive delivery or pure overhead), and
   **Admin/Overhead** (admin time and training; time off is excluded
   entirely since it isn't work time).
4. Because Idea, TechPass, SA NPD, and NPD effort has no dedicated timesheet
   code, the tool estimates how much time was spent on each based on how
   long it was active, and removes that estimate from the architect's
   Admin/Overhead bucket so that bucket isn't artificially inflated.
5. Each architect's "Effectiveness Score" is calculated as the percentage of
   their total time that falls into the Tangible/Delivery tier, and
   architects are ranked from highest to lowest.
6. The deliverables — an interactive leaderboard/scorecard web page and a
   summary spreadsheet — are written to the shared
   `SA\Outputs\sa-effectiveness-scorecard` folder.
7. A reviewer opens the scorecard, sees the ranked leaderboard and a
   color-coded Effectiveness Score (green/amber/red) for each architect, and
   can click any architect's row to drill into the specific Worklist items
   that make up their tiered time allocation.

## Technical Flow

- Single script `sa_effectiveness_scorecard.py`; uses **openpyxl** to read
  the Timesheet `.xlsx`, the standard library **csv** module to read the
  Worklist CSV, and hand-rolled HTML/JS (Chart.js + chartjs-plugin-datalabels
  via CDN) for the leaderboard chart and interactive scorecard/drill-down
  table — no pandas/plotly dependency. This logic directly extends
  `sa-workload-dashboard`'s imputation approach to also cover SA NPD and NPD.
- **Input resolution**: requires two explicit inputs (`--worklist`,
  `--timesheet`); prompts interactively if either is omitted.
  `--output-dir` defaults to the shared
  `SA\Outputs\sa-effectiveness-scorecard` folder via
  `sa_paths.output_dir_for("sa-effectiveness-scorecard")`.
- **Classification**: `load_worklist_items` classifies every Worklist row
  (all items, not just a recency window, since this skill measures overall
  time allocation rather than "currently active" counts) by the leading
  alphabetic prefix of `Work Ref` via `WORKREF_PREFIX_MAP` (`PRJ`→Project,
  `ID`→Idea, `TECHPASS`→TechPass, `SA`→SA NPD, `NPD`→NPD; unmatched prefixes
  become `"Unclassified"` rather than being dropped). `load_timesheet` locates
  the header row by scanning for a `"Segment Name"` first cell (to skip
  preamble rows), then classifies rows with `Project ID` equal to
  `--timeoff-project-id` (organization-specific; no default) by `Timesheet Line Class` text
  (`classify_timeclass`), and all other rows as Project/Idea by `Project ID`
  prefix. Distinct `Week Start` values are collected to define the
  imputation period.
- **Tier roll-up (`TIER_MAP`)**: Project/Idea/TechPass → **Tangible /
  Delivery**; SA NPD/NPD → **Internal / NPD**; Admin/Other and Training &
  Development → **Admin / Overhead** (Time Off is never tiered — excluded
  entirely).
- **Imputation (`compute_imputed_hours`)**: for each Idea/TechPass/SA
  NPD/NPD item with a parseable `Created` date, the item is "active" from its
  creation date, through its `Modified` date if it reached a terminal status
  (`Complete`/`Cancelled`), otherwise through the end of the timesheet's week
  range. For every observed Timesheet `Week Start` the item is active in, it
  accrues its category's configured rate: `--idea-hours-per-week` (default
  1.0), `--techpass-hours-per-week` (default 2.0),
  `--sa-npd-hours-per-week` (default 1.0), `--npd-hours-per-week` (default
  1.0). `_parse_date` tries several date formats (`%m/%d/%Y %I:%M %p`,
  `%m/%d/%Y`, `%d-%b-%y`, `%d-%b-%Y`).
- **Per-SA tier hours (`main`)**: Tangible = actual Project timesheet hours +
  imputed Idea + imputed TechPass. Internal/NPD = imputed SA NPD + imputed
  NPD. Admin/Overhead = Admin/Other (netted of all four imputed totals,
  floored at 0, unless `--no-net-admin`) + Training & Development.
- **Effectiveness Score**: `tangible_pct = tangible_hours / total_hours * 100`
  for each SA, where `total_hours` is the sum of all three tiers; this is the
  headline ranking metric and also drives the score-pill color (green ≥50%,
  amber 25–49%, red <25%).
- **Ranking and rendering**: `scorecard_rows` are sorted by `tangible_pct`
  descending; the leaderboard chart datasets (`tier_datasets`) and
  drill-down table (keyed per SA, each item tagged with its tier) follow the
  same ranked SA order. `build_html` embeds all data as JSON for a
  client-side, no-server interactive page (click a scorecard row to view
  that SA's underlying Worklist items and tiers).
- **Outputs**: `sa-effectiveness-scorecard.html` (interactive leaderboard +
  scorecard + drill-down) via `Path.write_text`, and
  `sa-effectiveness-summary.csv` (one row per SA: rank, total hours, hours
  and % per tier, and Effectiveness Score) via the standard library `csv`
  writer. Both default to `SA\Outputs\sa-effectiveness-scorecard\`.
- **Edge cases handled**: SAs present in only one source still appear
  (zero-filled on the missing side); unclassified Worklist prefixes remain
  visible in the drill-down but are excluded from tier scoring (a
  data-quality signal rather than a silent drop or guess); imputation is
  skipped for items without a usable `Created` date or when no Timesheet
  `Week Start` values exist.

## Inputs

- **Worklist CSV export** (`--worklist`, required) — must contain:
  `Work Ref`, `Work Ref: Solution Architect`, `Status`, `Created`,
  `Modified`.
- **Timesheet Excel export** (`--timesheet`, required) — must contain
  (after its header row): `Resource Name`, `Project ID`, `Timesheet Line
  Class`, `Hours Reported`; ideally also `Week Start` (used to define the
  imputation period).
- Optional flags: `--exclude-sa` (repeatable, exact "Last, First" match;
  no exclusions by default), `--timeoff-project-id` (organization-specific; no default),
  `--idea-hours-per-week` (default 1.0), `--techpass-hours-per-week`
  (default 2.0), `--sa-npd-hours-per-week` (default 1.0),
  `--npd-hours-per-week` (default 1.0), `--no-net-admin`, `--output-dir`
  (default shared `SA\Outputs\sa-effectiveness-scorecard`).

## Outputs

- `sa-effectiveness-scorecard.html` — interactive leaderboard chart, ranked
  scorecard table with color-coded Effectiveness Score pills, and
  click-to-drill-down Worklist item detail per Solution Architect.
- `sa-effectiveness-summary.csv` — one row per SA: rank, total hours, hours
  and % for each of the three tiers, and Effectiveness Score, written to
  `SA\Outputs\sa-effectiveness-scorecard\` by default.
