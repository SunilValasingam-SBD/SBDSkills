# SA Engagement Analysis

## Executive Summary

Leadership needs to know whether Solution Architects are being brought into initiatives early
— at the idea stage, where they can influence direction and reduce rework — or only once an
initiative is already a funded project. This skill turns a raw workload export into a clear,
interactive picture of engagement patterns across the SA team, broken out by business
portfolio and by individual architect. It highlights how much engagement is genuinely
"end-to-end" (idea through project) versus "project-only," giving leadership an evidence-based
way to assess whether architects are engaged soon enough, where early engagement is strong,
and where it is lagging, without anyone manually cross-referencing ideas and projects by hand.

## Business Process Flow

1. A business user provides a CSV export containing all idea and project initiatives, along
   with which Solution Architect and business portfolio each is tied to, and when each was
   created. If no file is supplied, the tool automatically uses the most recently added CSV
   in the shared inputs folder, or prompts the user for a path if none is found.
2. The user specifies (or accepts the default of) the calendar year to analyze.
3. The tool keeps only genuine idea and project records created in that year, discarding
   anything else, and classifies each one as either an "Idea" or a "Project."
4. For every project, the tool determines whether it was linked back to an idea the architect
   was already working — meaning the architect engaged "end-to-end" — or whether it appeared
   directly as a project with no prior idea-stage involvement ("project-only").
5. The tool rolls these classifications up two ways: by business portfolio and by individual
   architect, calculating total engagement, idea counts, project counts, end-to-end project
   counts, and the percentage of work that was end-to-end versus project-only.
6. The tool builds an interactive dashboard (viewable in any web browser) that lets a
   reviewer switch between the portfolio view and the SA view, see a month-by-month trend of
   engaged initiatives, and drill into end-to-end percentages against a team average.
7. The dashboard and two supporting summary spreadsheets (one row per portfolio/SA, and one
   row per underlying record) are saved to the shared outputs folder for this skill, ready to
   be shared with leadership or used in a governance/performance review. The original source
   CSV is never modified.

## Technical Flow

- **Main script**: `sa_engagement_analysis.py`, Python 3.10+, using `pandas` for data loading/
  transformation and `plotly.graph_objects` for the interactive HTML dashboard; `argparse` and
  `pathlib` for CLI plumbing.
- **Input resolution**: `--input` is optional. If omitted, `resolve_input_path` calls
  `latest_input("*.csv")` from the shared `_shared\sa_paths.py` helper to find the newest
  `.csv` in `SA\Inputs`; if none is found, it falls back to an interactive prompt for a path.
  The file must have a `.csv` extension and must exist, or the script exits with an error.
- **Output resolution**: `parse_args` calls `output_dir_for("sa-engagement-analysis")` from
  `_shared\sa_paths.py` to create/return `SA\Outputs\sa-engagement-analysis\`, used as the
  default directory for all three output files (`--output`, `--summary-output`,
  `--detail-output` can override individually).
- **Column detection**: `REQUIRED_COLUMNS` = `Work Ref`, `Idea ID`, `Program Owning
  Portfolio` must be present or the script exits. The creation-date column is auto-detected
  via `select_column` against `DATE_CANDIDATES` (`Created`, `Created Date`,
  `Initiative Created`, `Created On`, `Date Created`), and the SA column against
  `SA_CANDIDATES` (`Work Ref: Solution Architect`, `Solution Architect`, `SA`,
  `Assigned to`) — both case-insensitive exact matches; ambiguous (multiple) or missing
  matches raise an actionable error unless `--created-column`/`--sa-column` is passed
  explicitly.
- **Filtering and classification** (`load_and_classify`): rows are kept only if the detected
  creation date falls in the requested `--year` and `Work Ref` (uppercased) starts with `ID`
  or `PRJ`. `Engagement Type` is `Project` if `Work Ref` starts with `PRJ`, else `Idea`.
  `E2E Engagement` is true only for Project rows with a non-blank `Idea ID` (blank includes
  empty string, whitespace, `nan`, `none`, `null` via `is_blank`); `Project-only Engagement`
  is Project rows that are not E2E.
- **Portfolio normalization**: `Program Owning Portfolio` is cleaned (blank → `Unassigned`),
  trailing slashes trimmed, and only the last `/`-delimited segment kept; a hardcoded
  `PORTFOLIO_GROUP_MAP` then consolidates related portfolio names (e.g. `CIO Office`,
  `HR, Legal, & Compliance`, and `Mergers, Acquisitions & Divestitures` all roll up into
  `HR & Legal Compliance`; `Commercial & Regions` rolls into `Customer Engagement`) for
  cleaner charting. The SA column is similarly blank-filled to `Unassigned`.
- **Aggregation** (`summarize`): groups by `Portfolio` or `Solution Architect`, computing
  distinct `Work Ref` counts for total engagement, ideas, projects, and E2E projects; derives
  `Project-only Projects` (projects minus E2E, floored at 0), `E2E % of Total`, and
  `Engaged % of Total` (= (Ideas + E2E Projects) / Total, i.e. everything except
  project-only rows).
- **Dashboard construction** (`build_dashboard`): builds three Plotly figures combined into one
  self-contained HTML file —
  1. an "Engagement Overview" grouped bar chart (Total Engagement / Ideas / Projects / E2E
     Projects) with an overlaid "Engaged % of Total" line and a team-average dashed line,
     toggleable between Portfolio and SA views via a dropdown (`build_engagement_figure`);
  2. an "Engaged Initiatives by Month" stacked bar chart per category (Portfolio or SA),
     built from a month-ordered pivot table of distinct engaged `Work Ref` counts
     (`build_monthly_figure`), with a dotted total-team-engaged trend line;
  3. an "E2E Engagement Excluding Ideas" chart restricted to Project rows only, showing
     Projects / E2E Projects / Project-only Projects and `E2E % of Total` = E2E / Projects.
  The three figures are concatenated into one HTML document (`<html>...</html>`) with Plotly
  JS embedded once (`include_plotlyjs=True` on the first figure only) and written to the
  resolved `--output` path.
- **CSV outputs**: `summarize` results for both dimensions are concatenated and written to
  `--summary-output`; the classified detail frame (`Work Ref`, `Idea ID`, `Created Date`,
  `Portfolio`, `Solution Architect`, `Engagement Type`, `E2E Engagement`,
  `Project-only Engagement`) is written to `--detail-output`. Both use UTF-8 with BOM
  (`utf-8-sig`) for Excel compatibility.
- **Edge cases handled**: no parseable dates in the detected date column raises an error; the
  source CSV is read as all-string dtype (`dtype=str`) to avoid pandas coercing IDs/codes, and
  is never written back to.

## Inputs

- **Source CSV** (required; `--input`, or auto-detected as the newest `.csv` in `SA\Inputs`,
  or prompted for interactively), expected columns:
  - `Work Ref` — initiative reference; must start with `ID` (idea) or `PRJ` (project) to be
    included.
  - `Idea ID` — populated on a project row if it originated from a tracked idea (drives
    end-to-end classification).
  - `Program Owning Portfolio` — business portfolio grouping.
  - A creation-date column, auto-detected from `Created`, `Created Date`,
    `Initiative Created`, `Created On`, or `Date Created` (override with `--created-column`).
  - A Solution Architect column, auto-detected from `Work Ref: Solution Architect`,
    `Solution Architect`, `SA`, or `Assigned to` (override with `--sa-column`).
- **Optional**: `--year` (defaults to the current calendar year) to analyze a different
  reporting period.

## Outputs

- `SA\Outputs\sa-engagement-analysis\sa_engagement_dashboard.html` — interactive dashboard
  with Portfolio/SA view toggles: engagement overview, monthly engagement trend, and E2E vs.
  project-only engagement.
- `SA\Outputs\sa-engagement-analysis\sa_engagement_summary.csv` — one row per Portfolio and
  one row per Solution Architect with total engagement, idea/project/E2E counts, and
  percentages.
- `SA\Outputs\sa-engagement-analysis\sa_engagement_detail.csv` — normalized, record-level
  detail of every included initiative with its classification fields.
