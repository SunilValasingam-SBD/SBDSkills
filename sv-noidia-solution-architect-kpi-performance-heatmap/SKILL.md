---
name: sv-noidia-solution-architect-kpi-performance-heatmap
description: 'Generate an interactive WIP KPI performance heatmap per Solution Architect. Measure TechPass, scenario publication, and SARB approval across PRJ projects and ID ideas; measure SAD approval only across PRJ projects.'
argument-hint: 'Provide the Work Queue CSV/Excel export when prompted'
user-invocable: true
disable-model-invocation: false
---

# SV-Noidia-Solution Architect KPI Performance Heatmap

Generate a four-KPI interactive heatmap with click-to-drill-down using a Work Queue CSV or Excel export.

## KPI rules and denominator

- Include WIP records by default. Only `Work Ref` values beginning with `PRJ` (projects) or `ID` (ideas) are in scope.
- **TechPass Created** — `TechPass` is populated; denominator includes WIP PRJ and ID work items.
- **Scenario Published** — `Publication Status` is populated and is not `Not Published`; denominator includes WIP PRJ and ID work items.
- **SARB Approved** — `SARB Approval Date` is populated; denominator includes WIP PRJ and ID work items.
- **SAD Approved** — `SAD Approved` is populated; denominator includes only WIP PRJ projects. ID ideas are excluded from both the numerator and denominator because SAD approval is not applicable to them.
- If an architect has no WIP PRJ projects, show SAD Approved as `N/A` (`0/0`), not 0%.

## Filters

- The default status is `WIP`. Add `--status` only when the user requests a different status.
- Exclude `Unassigned` architect records by default. Add other architect exclusions only when the user requests them.
- Keep the source file unchanged.

## Procedure

1. Use the supplied Work Queue CSV or Excel export. A SharePoint page URL is not a data export; if the export is not attached or accessible as a local file, ask the user to upload it.
2. Verify these required columns exist: `Work Ref`, `Work Ref: Name`, `Work Ref: Solution Architect`, `Status`, `TechPass`, `Publication Status`, `SARB Approval Date`, `SAD Approved`.
3. Run the bundled generator, providing explicit output paths beside the source file when possible:

   ```powershell
   python .\sa_kpi_heatmap.py --input "<source file path>" --exclude-sa Unassigned --status WIP --output "<source directory>\sv_noidia_sa_kpi_performance_heatmap.html" --summary-output "<source directory>\sv_noidia_sa_kpi_summary.csv"
   ```

4. Open the HTML heatmap and click a cell to inspect the eligible work items that met or missed the KPI. For SAD Approved, the drill-down must contain PRJ work references only.
5. Review the summary CSV. It reports total WIP work items, SAD-eligible PRJ project count, and each KPI's met/denominator and percentage.

## Output

- `sv_noidia_sa_kpi_performance_heatmap.html` — interactive KPI heatmap with project-level drill-down.
- `sv_noidia_sa_kpi_summary.csv` — per-architect KPI percentages and counts.

The script requires Python 3.10+ and the packages in `requirements.txt` (`pandas`, `plotly`, `openpyxl`).
