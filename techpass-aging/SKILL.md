---
name: techpass-aging
description: 'Generate reusable TechPass or ticket aging reports from an Excel workbook. Use when the user wants aging by Created date, counts by aging week, Solution Architect breakdowns, an interactive chart, a CSV summary table, or a one-slide PowerPoint.'
argument-hint: '[Excel path] [as-of date YYYY-MM-DD]'
user-invocable: true
disable-model-invocation: false
---

# TechPass Aging Report

Generate an interactive ticket-aging chart, a summary CSV, and a one-slide PowerPoint from an Excel workbook.

## Use When

- A weekly Excel export contains the same ticket fields but a new set of rows.
- The user wants aging calculated from `Created` through today or a reporting date.
- The user wants Solution Architects on the X axis, aging buckets on the Y axis, and count labels at each point.
- The user wants a table with one Solution Architect per row and counts for Week 1, Week 2, Week 3, Week 4, Week 5, and Above 5 Weeks.

## Required Workbook Fields

The first worksheet must contain headers in its first row. Required columns are:

- `Number`
- `State`
- `Assigned to`
- `Created`

All ticket statuses are included. `Assigned to` is used as the Solution Architect label.

## Procedure

1. Identify the Excel input path. If no path is provided, prompt the user to enter or select the workbook location; do not assume a filename such as `TP.xlsx`.
2. Confirm that the workbook has the required fields and that `Created` values are valid dates.
3. Run the bundled generator:

   ```powershell
   py .\techpass_aging_chart.py --input "<Excel path>" --as-of YYYY-MM-DD
   ```

   Omit `--as-of` to use the current date. The default output files are created beside the script:

   - `techpass_aging_by_architect.html`
   - `techpass_aging_summary.csv`
   - `techpass_aging_summary.pptx`

4. Report the number of tickets included and the output file paths.
5. If the input has a different filename, pass it with `--input`; do not modify the source workbook.

The bundled generator requires Python 3.10 or newer and the `openpyxl` and `python-pptx` packages. Install them with `py -m pip install openpyxl python-pptx`.

## Aging Rules

- `Age (days) = max(0, as-of date - Created date)`.
- `Age (weeks) = floor(Age (days) / 7)`.
- Aging buckets are `Week 1`, `Week 2`, `Week 3`, `Week 4`, `Week 5`, and `Above 5 Weeks`.
- A ticket with 0 through 6 days of age is placed in `Week 1`; 7 through 13 days in `Week 2`; and so on. Tickets aged 35 days or more are placed in `Above 5 Weeks`.

## Historical Reports

Use a fixed reporting date for a prior weekly snapshot:

```powershell
py .\techpass_aging_chart.py --input "<Excel path>" --as-of 2026-08-21
```

The HTML chart uses the six aging buckets as colored rows and places each nonzero architect/bucket intersection on the chart with its ticket count as a label. Hover details include ticket number, created date, status, and exact age.
