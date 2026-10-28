# SV-Noidia-Solution Architect KPI Performance Heatmap

Generate an interactive heatmap for Solution Architect governance KPIs from a Work Queue CSV or Excel export.

## KPI scope

The report includes only WIP work references beginning with `PRJ` or `ID`.
TechPass Created, Scenario Published, and SARB Approved use all included PRJ
projects and ID ideas as their denominators. SAD Approved uses **PRJ projects
only**; ID ideas do not count as either met or missing. An architect without
any WIP PRJ projects is shown as `N/A` (`0/0`) for SAD Approved.

## Run

```powershell
python .\sa_kpi_heatmap.py --input "<path to Work Queue CSV or Excel export>" --exclude-sa Unassigned --status WIP
```

By default, the HTML and CSV files are written under the shared
`SA\Outputs\sv-noidia-solution-architect-kpi-performance-heatmap` folder.
Use `--output` and `--summary-output` to save them beside the source export:

```powershell
python .\sa_kpi_heatmap.py --input "<source path>" --exclude-sa Unassigned --status WIP --output "<source directory>\sv_noidia_sa_kpi_performance_heatmap.html" --summary-output "<source directory>\sv_noidia_sa_kpi_summary.csv"
```

The summary distinguishes total WIP work items from SAD-eligible PRJ projects.
Heatmap drill-downs show eligible work items for each KPI; the SAD Approved
drill-down contains PRJ projects only.

## Required columns

- `Work Ref`
- `Work Ref: Name`
- `Work Ref: Solution Architect`
- `Status`
- `TechPass`
- `Publication Status`
- `SARB Approval Date`
- `SAD Approved`

## Dependencies

Python 3.10+ and the packages listed in `requirements.txt`:

```powershell
python -m pip install -r requirements.txt
```
