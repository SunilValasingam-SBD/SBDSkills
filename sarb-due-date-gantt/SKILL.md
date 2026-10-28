# SARB Due Date Heatmap

Generate a single interactive heatmap (HTML) showing SARB Due Date workload by **Solution Architect x Month**, for Projects/Ideas that have not yet been reviewed/approved — with per-cell counts, overdue items flagged, and marginal totals per month and per SA.

## Use When

- The user has the same Work List export used by the other SA skills (it may contain a mix of Projects, Ideas, TechPass requests, SA NPD, and NPD rows) and wants to see, visually, how many **Projects and Ideas only** each Solution Architect has coming due each month — TechPass/NPD/SA NPD rows are automatically excluded.
- The user wants overdue items (due date before today) clearly flagged without the clutter of one bar per item.
- The user wants month-level granularity (not individual day-by-day bars) scoped to a fixed window: Q4 2026 through the end of 2027.

## Required Workbook/CSV Fields

The first worksheet (if Excel) or the CSV must contain headers in its first row. Required columns:

- `Work Ref`
- `Work Ref: Name`
- `Work Ref: Solution Architect`
- `SARB Due Date`

Optional but used when present:

- `SARB Rating` — if this column exists, rows that already have a rating are excluded by default (they're already reviewed, so they no longer have a "pending" due date). Pass `--include-approved` to keep them anyway.

This is the same Work List export used by the other SA skills, so no separate export is needed. The file may include every classification (Projects, Ideas, TechPass requests, SA NPD, NPD); this skill automatically keeps only **Projects** (`Work Ref` starting with `PRJ`) and **Ideas** (`Work Ref` starting with `ID`) and silently skips everything else.

## Procedure

1. Identify the CSV/Excel input path. If no path is provided, look for the newest `.csv`/`.xlsx`/`.xls` file in the shared `SA\Inputs` folder automatically.
2. Confirm the file has the required fields (`Work Ref`, `Work Ref: Name`, `Work Ref: Solution Architect`, `SARB Due Date`).
3. Run the bundled generator:

   ```powershell
   py .\sarb_due_date_gantt.py --input "<CSV or Excel path>"
   ```

   Omit `--input` to auto-detect the newest file in `SA\Inputs`. Omit `--as-of` to use today's date. Output is written to the shared `SA\Outputs\sarb-due-date-gantt\` folder:

   - `SARB_Due_Dates_Heatmap.html` — the Solution Architect x Month heatmap with margin totals.

4. Report the counts of overdue vs. upcoming items and the output file path.

## Options

- `--as-of YYYY-MM-DD` — use a fixed reference date instead of today (e.g. for a historical snapshot).
- `--include-approved` — also chart rows that already have a `SARB Rating` set (normally excluded).
- `--output-dir <path>` — override the default shared output folder.

## Chart Design

- Only **Projects** and **Ideas** are considered (`Work Ref` prefixed `PRJ`/`ID`); TechPass, SA NPD, and NPD rows in the same Work List export are excluded automatically.
- Only items with a `SARB Due Date` in **August 2026 through the end of 2027 (Dec 31, 2027)** are included; everything else is dropped.
- **Main grid**: one row per Solution Architect (busiest at top), one column per month (`Aug 2026` ... `Dec 2027`). Darker blue = more items due that SA/month; the bold number shown is simply the count, automatically rendered in white on the darkest cells (and dark gray elsewhere) so it stays legible everywhere. There is no separate color-scale legend/side chart or per-month totals bar — the grid is the whole story.
- A single **"Total: N"** label sits in the top-left corner, above the Solution Architect column, showing the grand total of in-scope items.
- 🟧 **Overdue flag**: any cell containing at least one overdue item gets a bold orange outline — no extra side chart needed, the count and the outline tell the full story.
- **Hover tooltip**: a quick preview on mouseover lists every individual item in that SA/month cell.
- **Click for details**: clicking any cell opens a styled details panel below the chart with the full item list for that cell — reference ID, name, status, and exact due date — so nothing is hidden behind a hover-only tooltip.
- A dashed line marks **today's exact position** — proportional to how far the current day is through its month, not just the start of the month (e.g. late October sits near the right edge of the Oct 2026 column, not the left edge).
- Month labels are angled (-45°) with a larger, clean font so all 17 months stay readable without overlapping.

The bundled generator requires Python 3.10+ and the `pandas`, `plotly`, and `openpyxl` (for `.xlsx`/`.xls` input) packages. Install with `py -m pip install pandas plotly openpyxl`.
