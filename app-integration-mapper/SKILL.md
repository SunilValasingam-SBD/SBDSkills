---
name: app-integration-mapper
description: 'Given an Excel workbook with an applications list, generate a workbook with one tab per application listing every documented integration where that application is the Initiator or the Terminator. The integration catalog is pulled live from the local `ardoq` MCP server''s Integration Catalog report (not from an Excel tab). Use when the user has a list of applications and wants to know all integrations per application sourced from Ardoq.'
argument-hint: '[Excel path]'
user-invocable: true
disable-model-invocation: false
---

# Application Integration Mapper

Map each application in an application list to every documented integration (as Initiator or
Terminator) from Ardoq's "Integration Catalog" report, and produce a new Excel workbook with
one tab per application plus a Summary tab.

## Use When

- The user has an Excel workbook with a sheet listing applications (e.g. `ExcelApplications`),
  one application per row (an `Application Name` column, and ideally an `App ID` column).
- The user wants, for each application, the full list of integrations it participates in,
  sourced from the live Ardoq "Integration Catalog" report (via the local `ardoq` MCP server),
  not a static Excel catalog tab.
- The user wants applications with no documented integrations excluded (default) or
  included as empty sheets.

## Required Inputs

- Applications workbook: a sheet with a column for the application name (default header
  `Application Name`). An optional `App ID` column improves match accuracy (matches both
  by name and by Ardoq App ID), since some apps in Ardoq are duplicated/renamed and only
  reliably distinguishable by App ID.
- Ardoq MCP server (`ardoq`, configured in `~/.copilot/mcp-config.json`, implemented at
  `C:\Users\<user>\ardoq-mcp\server.py`) must be reachable to fetch the integration catalog.

Column and sheet names are configurable (see Procedure) because different exports use
different headers — only the app-name / app-id columns matter for the match.

## Procedure

1. **Identify the Excel input path.** If no path is provided, ask the user for it; do not
   assume a filename.

2. **Ensure the `ardoq` MCP server's tools are loaded in this session.** Call the tool
   search tool for a pattern like `ardoq` (or `get_integration_catalog`). Tool search
   failing to find it means the server isn't loaded in the current session (MCP servers
   only load at session startup):
   - Run `copilot mcp list` to confirm the `ardoq` server is configured. If it's missing
     entirely, tell the user it needs to be set up first (see the `ardoq-mcp` project the
     user previously had built) and stop.
   - If it's configured but just not loaded in this session, tell the user: "The ardoq MCP
     tools aren't loaded in this session. Please run `/restart` (or start a new session) so
     the ardoq MCP server loads, then re-run this skill." Stop here and wait.
   - If tool search finds `ardoq-get_integration_catalog` (or similar `ardoq-*` tools),
     proceed.

3. **Fetch the integration catalog from Ardoq and save it to a temp JSON file:**
   Call the `ardoq-get_integration_catalog` tool (defaults to `keyFormat=key`, follows
   pagination automatically). It returns an object like
   `{"_meta": {"columns": [...]}, "values": [...]}`. Write that exact JSON object to a
   temp file, e.g. `<session workspace>\integration_catalog.json`, since the row count is
   usually in the thousands and too large to pass inline.

   If `ardoq-get_integration_catalog` isn't available for some reason, fall back to
   `ardoq-run_report` with the Integration Catalog report id `16764d6326d36ca5af13d6b7`.

4. **Run the bundled generator** pointing at the applications workbook and the catalog
   JSON file:

   ```powershell
   py .\map_app_integrations.py --input "<Excel path>" --catalog-json "<path to integration_catalog.json>"
   ```

   Useful optional flags:

   ```powershell
   py .\map_app_integrations.py `
     --input "<Excel path>" `
     --catalog-json "<path to integration_catalog.json>" `
     --output "<output Excel path>" `
     --apps-sheet "ExcelApplications" `
     --app-name-column "Application Name" `
     --app-id-column "App ID" `
     --initiator-column "Initiator App" `
     --terminator-column "Terminator" `
     --initiator-id-column "Initiator App ID" `
     --terminator-id-column "Terminator App ID" `
     --include-empty
   ```

   - Omit `--output` to default to `<input>_Integrations.xlsx` next to the source file.
   - By default, applications with zero matching integrations are skipped entirely (no
     sheet, no summary row). Pass `--include-empty` to keep them as empty sheets with a
     summary row showing a count of 0.
   - Legacy mode: omitting `--catalog-json` falls back to reading a static `Integration
     Catalog` sheet from the input workbook instead (original Excel-only behavior), via
     `--catalog-sheet`. Prefer the MCP JSON path whenever Ardoq is reachable, since it's
     always current.

5. **Report** the output file path, total applications considered, and how many had
   documented integrations (sheets created).

6. **Do not modify the source workbook**; only the new output workbook is written. Delete
   the intermediate `integration_catalog.json` temp file once the run completes
   successfully, unless the user wants to keep it.

## Matching Rules

- A catalog row is attributed to an application if the application name exactly matches
  (case-insensitive) a value in the Initiator App or Terminator column on that row — these
  columns from the MCP JSON catalog may be lists of names, which are matched by membership.
- If the applications sheet has an App ID column, a row is also attributed to the
  application if its App ID matches the catalog's Initiator App ID or Terminator App ID —
  this is more reliable than name matching since Ardoq sometimes has multiple
  similarly-named app records (e.g. "Altium", "Altium - GTS", "Altium - SEF" are distinct
  App IDs with distinct integration records).
- Sheet names are sanitized to Excel's 31-character limit and illegal characters
  (`[ ] : * ? / \`) are replaced with `-`; duplicate resulting names get a numeric suffix.
- List-valued cells (e.g. `Initiator App`) are flattened to a comma-separated string in the
  output sheet.

## Notes

- Requires Python 3.8+ and `openpyxl` (`py -m pip install openpyxl`).
- If an application has zero integrations, that's expected and reflects the data (some
  apps genuinely have none documented) — don't assume it's a mapping bug. Flag large,
  unexpected all-zero results to the user rather than silently fuzzy-matching names.
- The Ardoq token used by the MCP server may lack read access to some workspaces (e.g.
  infrastructure/hosting); this does not affect the Integration Catalog report itself,
  which is a precomputed graph-search report independent of per-workspace permissions.
