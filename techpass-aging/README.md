# TechPass Aging Report

## Executive Summary

Unresolved TechPass tickets that sit untouched for weeks are a visibility risk — without a consistent way to see how old a ticket is and who owns it, aging backlogs can hide until they become escalations. This skill turns a routine weekly TechPass export into an instant aging report: how many tickets each Solution Architect is carrying, and how long each one has been open, grouped into easy-to-read weekly buckets. It replaces a manual pivot-table exercise with a one-command refresh, and produces outputs in three formats — an interactive chart for exploration, a spreadsheet for further analysis, and a ready-to-present PowerPoint slide — so the same data can support a team working session, a leadership readout, or a historical trend comparison, with no extra formatting effort each week.

## Business Process Flow

1. A team member exports the current TechPass ticket list to an Excel workbook (the standard weekly export already in use).
2. The workbook is dropped into the shared `SA\Inputs` folder (or an explicit path is supplied), and the skill is run via Copilot CLI (`/techpass-aging`) or directly with Python, optionally specifying a reporting ("as-of") date for a historical snapshot instead of today.
3. The tool checks that the workbook has the expected fields and valid creation dates for every ticket; all ticket statuses are included, regardless of state.
4. For every ticket, the tool calculates its age in days from its creation date through the as-of date, and places it into one of six aging buckets: Week 1 through Week 5, or Above 5 Weeks for anything older.
5. The tool groups tickets by the Solution Architect they're assigned to and counts how many fall into each aging bucket, producing a team-wide view of who is carrying the oldest backlog.
6. Three deliverables are generated automatically and written to the shared `SA\Outputs\techpass-aging` folder: an interactive chart showing each Solution Architect's ticket counts across aging buckets (with hover detail per ticket), a summary spreadsheet with one row per Solution Architect and a column per aging bucket, and a single PowerPoint slide combining the same visual and table for easy sharing in a meeting or status deck.
7. The tool reports back how many tickets were included and where the output files were written, so the requester can confirm the run succeeded before sharing results.

## Technical Flow

- **Entry point:** `techpass_aging_chart.py`, run via `py .\techpass_aging_chart.py --input "<Excel path>" [--as-of YYYY-MM-DD]` (directly, or via the `/techpass-aging` Copilot CLI skill). Requires Python 3.10+ and the packages listed in `requirements.txt`: `openpyxl` (Excel reading), `Pillow`/`PIL` (image drawing for the PowerPoint slide), `python-pptx` (PowerPoint generation), and `plotly` (interactive HTML chart, via `plotly.graph_objects` and `plotly.subplots.make_subplots`).
- **Shared path resolution (`parse_args`):** Inserts `..\_shared` onto `sys.path` and imports `output_dir_for` from `sa_paths.py` to resolve/create `SA\Outputs\techpass-aging` as the default location for `--output` (HTML) and `--pptx-output` (PowerPoint); the CSV summary is always written beside the HTML output (same folder, fixed filename). If `--input` is omitted, `latest_input("*.xlsx")` is called to pick the most recently modified `.xlsx` file in the shared `SA\Inputs` folder, raising `SystemExit` if none is found.
- **Workbook loading (`load_tickets`):** Uses `openpyxl.load_workbook(..., read_only=False, data_only=True)` deliberately *not* in optimized read-only mode, because some exported workbooks declare incorrect row/column spans that cause the read-only parser to truncate data. Reads the first worksheet's header row, validates that `Number`, `State`, `Assigned to`, and `Created` are all present (raising `ValueError` listing any missing columns), then iterates every data row into a dict keyed by header.
- **Aging calculation:** For each ticket, `Created` is coerced to a `date` (handling `datetime` values) and the script raises `ValueError` if it isn't a valid date. `Age (days) = max(0, as_of - created)`, `Age (weeks) = age_days // 7`, and `Aging bucket = AGING_BUCKETS[min(age_days // 7, 5)]` where `AGING_BUCKETS = ("Week 1", "Week 2", "Week 3", "Week 4", "Week 5", "Above 5 Weeks")` — so 0–6 days is Week 1, 7–13 is Week 2, ... and 35+ days always collapses into the final "Above 5 Weeks" bucket via the `min(..., 5)` cap. `Assigned to` becomes `Solution Architect` (defaulting to `"Unassigned"` if blank) and `State` becomes `Status` (defaulting to `"Unknown"`); blank/header-echo rows (where `Number` is missing or literally the string `"Number"`) are skipped.
- **Summary CSV (`write_summary`):** Uses Python's built-in `csv` module to write one row per Solution Architect (sorted alphabetically) with a count per aging bucket (via `collections.Counter` keyed on `(architect, bucket)` pairs) plus a row total.
- **Interactive chart (`create_chart`):** Builds a two-row Plotly subplot figure: row 1 is a scatter plot with Solution Architects on the x-axis (categorical, sorted) and aging buckets on the y-axis (categorical, in bucket order), where each non-zero architect/bucket combination is plotted as a colored circular marker sized to fit its count label, with a hover template showing architect, bucket, and ticket count; row 2 is a `go.Table` rendering the same counts as a grid with a Total column. The figure is written to HTML via `figure.write_html(output_path, include_plotlyjs=True, full_html=True)`, producing a fully self-contained, offline-viewable file.
- **PowerPoint slide (`create_powerpoint`):** Does not use Plotly for the slide; instead it renders the entire visual by hand with `PIL.Image`/`ImageDraw`/`ImageFont` (using the Windows-installed Segoe UI font files) onto an in-memory 1800x1000 canvas — drawing the scatter-style chart as circles with count labels and the summary table as a manually positioned grid — then saves that canvas as a temporary PNG, inserts it as a full-slide picture into a blank `python-pptx` `Presentation` (16:9, 13.333"x7.5"), saves the `.pptx`, and deletes the temporary PNG.
- **Outputs:** All three files (`techpass_aging_by_architect.html`, `techpass_aging_summary.csv`, `techpass_aging_summary.pptx`) default to the resolved `SA\Outputs\techpass-aging` folder; all can be overridden with `--output` and `--pptx-output` (the CSV path is always derived from `--output`'s folder via `.with_name(...)`).
- **Edge cases handled:** negative ages are floored to zero (`max(0, ...)`) so an as-of date mistakenly set before a ticket's creation date doesn't produce a nonsensical negative age; any ticket with an unparseable `Created` value stops the run with a clear error rather than silently miscounting; the script never writes back to the source workbook.

## Inputs

- An Excel workbook export of TechPass tickets (explicit `--input` path, or auto-detected as the newest `*.xlsx` file in the shared `SA\Inputs` folder). The first worksheet must have headers in row 1, including:
  - `Number` — ticket identifier.
  - `State` — ticket status (all statuses are included in the report).
  - `Assigned to` — used as the Solution Architect label (defaults to "Unassigned" if blank).
  - `Created` — ticket creation date, used as the aging clock start.
- Optional `--as-of YYYY-MM-DD` (defaults to today) for a historical/point-in-time snapshot.
- Optional `--output` and `--pptx-output` paths (default: `SA\Outputs\techpass-aging\techpass_aging_by_architect.html` and `...\techpass_aging_summary.pptx`).

## Outputs

All written to `SA\Outputs\techpass-aging\` by default (or custom `--output`/`--pptx-output` paths):

- `techpass_aging_by_architect.html` — interactive chart (scatter plot of Solution Architect x aging bucket with ticket counts, plus a summary table) with per-ticket hover details.
- `techpass_aging_summary.csv` — one row per Solution Architect with ticket counts for Week 1–Week 5, Above 5 Weeks, and a total.
- `techpass_aging_summary.pptx` — a single PowerPoint slide combining the same chart visual and summary table for presentations.
