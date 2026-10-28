# Application Integration Mapper

## Executive Summary

In a divestiture, understanding every system dependency tied to the applications in scope is
critical to planning a clean and risk-free separation. This capability takes the list of
applications involved in the divestiture and produces a complete, application-by-application
view of their integrations — showing what each system sends data to and receives data from
across the environment. This gives the divestiture team an immediate, reliable basis for
assessing separation impact, identifying shared or dependent systems that complicate
carve-out, and informing the broader separation plan — without manual, error-prone discovery
across multiple teams and sources.

## Business Process Flow

1. A business user (or another skill) provides a list of applications in an Excel workbook —
   typically a sheet named `ExcelApplications` with an "Application Name" column, and ideally
   an "App ID" column for more precise matching.
2. If no file is explicitly provided, the tool automatically picks up the most recently added
   Excel file from the shared `SA\Inputs` drop folder.
3. The tool connects to Ardoq's live "Integration Catalog" (the enterprise's authoritative,
   always-current record of application integrations) rather than relying on a static export,
   ensuring the results reflect what's documented in Ardoq today.
4. For each application in the list, the tool scans the entire integration catalog and
   collects every integration record where that application is either the sender
   (Initiator) or the receiver (Terminator) of data.
5. The tool builds a new Excel workbook: a "Summary" tab listing every application and how
   many integrations it has, plus one dedicated tab per application containing the full
   detail of its integrations.
6. Applications with no documented integrations are left out of the workbook by default
   (since an empty tab adds no value), unless the user asks to see them explicitly.
7. The finished workbook is saved to the shared `SA\Outputs\app-integration-mapper\` folder,
   ready to be reviewed, shared, or attached to an architecture assessment or SAD. The
   original source files (the application list and Ardoq data) are never modified.

## Technical Flow

- **Orchestration**: the SKILL.md instructs the CLI to first confirm the `ardoq` MCP server's
  tools are loaded in the session, then call `ardoq-get_integration_catalog` (falling back to
  `ardoq-run_report` against report id `16764d6326d36ca5af13d6b7`) and write the returned
  `{"_meta": {...}, "values": [...]}` JSON payload to a temporary file, since the catalog is
  typically thousands of rows.
- **Main script**: `map_app_integrations.py`, written in Python 3.8+, using `openpyxl` for all
  Excel reading/writing (no pandas dependency) and the standard library `argparse`, `json`,
  `re`, and `pathlib`.
- **Input resolution**: `--input` is optional. If omitted, the script calls
  `latest_input("*.xlsx")` from the shared `_shared\sa_paths.py` helper to pick the most
  recently modified `.xlsx` file in `SA\Inputs`. If neither is available, the script exits
  with an error.
- **Output resolution**: `--output` is optional. If omitted, the script calls
  `output_dir_for("app-integration-mapper")` from `_shared\sa_paths.py`, which creates (if
  needed) and returns `SA\Outputs\app-integration-mapper\`, and writes
  `<input-basename>_Integrations.xlsx` there.
- **Catalog loading** (`load_catalog_json`): accepts either the raw Ardoq report object
  (extracting `headers` from `_meta.columns` and rows from `values`) or a bare list of row
  dicts; falls back to the union of keys across all rows if no explicit header list is given.
  A legacy mode (`--catalog-sheet`, default `"Integration Catalog"`) can instead read a static
  Excel tab from the input workbook when `--catalog-json` isn't supplied.
- **Matching logic** (`values_match`): case-insensitive string comparison; catalog cells for
  Initiator/Terminator may be lists of names (flattened via `flatten_cell` for output), so a
  match is true if the application name equals any element of the list. If the applications
  sheet has an App ID column, a row is also matched when the App ID equals the catalog's
  Initiator App ID or Terminator App ID — this is more reliable than name matching because
  Ardoq can have multiple app records with similar names but distinct IDs.
- **Sheet construction**: one sheet per application with at least one match (`sanitize_sheet_name`
  enforces Excel's 31-character limit, strips illegal characters `[ ] : * ? / \`, and appends a
  numeric suffix on collisions). A bold header row is written from the catalog's header list,
  followed by each matched, flattened row. Column widths are auto-sized (10–60 chars). A
  "Summary" sheet (first tab) lists each application's name, integration count, and sheet name.
- **Zero-match handling**: by default, applications with no matches are skipped entirely (no
  sheet, no summary row); `--include-empty` keeps them with a header-only sheet and a summary
  row showing 0.
- **Output generation**: the workbook is assembled in-memory with `openpyxl.Workbook()` and
  saved via `wb.save(out)` to the resolved output path under
  `SA\Outputs\app-integration-mapper\`. The source workbook is opened with `data_only=True`
  and never written back to.

## Inputs

- **Applications workbook** (required, Excel `.xlsx`): defaults to the newest `.xlsx` file in
  `SA\Inputs` if `--input` is not given.
  - Sheet: `ExcelApplications` (configurable via `--apps-sheet`).
  - `Application Name` column (configurable via `--app-name-column`) — required.
  - `App ID` column (configurable via `--app-id-column`) — optional, improves match accuracy.
- **Integration catalog** (preferred: JSON): `--catalog-json`, a JSON file containing the
  Ardoq Integration Catalog report payload (`{"_meta": {...}, "values": [...]}` or a bare
  list), fetched live via the `ardoq` MCP server's `get_integration_catalog`/`run_report`
  tools. Relevant keys: `Initiator App`, `Initiator App ID`, `Terminator`,
  `Terminator App ID` (all configurable via `--initiator-column`, `--initiator-id-column`,
  `--terminator-column`, `--terminator-id-column`).
- **Integration catalog** (legacy fallback): a sheet in the same workbook, default name
  `Integration Catalog` (configurable via `--catalog-sheet`), used only when `--catalog-json`
  is omitted.
- Optional flag `--include-empty` to retain applications with zero matched integrations.

## Outputs

- `SA\Outputs\app-integration-mapper\<input-basename>_Integrations.xlsx` (or an explicit
  `--output` path) — a new workbook containing:
  - **Summary** tab — one row per application with its integration count and sheet name.
  - One tab per application (sanitized, ≤31-character sheet name) — every catalog row where
    that application is the Initiator or Terminator.
