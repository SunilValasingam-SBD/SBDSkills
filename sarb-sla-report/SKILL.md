---
name: sarb-sla-report
description: 'Analyze Solution Architecture Review Board SLA performance from a CSV or SharePoint export. Use for SARB SLA reports, Solution Architect performance, met/not-met project classification, Tuesday due-date rules, charts, and grouped exception details.'
argument-hint: 'Provide the refreshed CSV path'
user-invocable: true
disable-model-invocation: false
---

# SARB SLA Report

Analyze a supplied Work Queue extract and produce project-level and Solution Architect-level SLA results.

## Rules

- Include records whose `Created` date is in calendar year 2026.
- Include `Work Ref` values beginning with `PRJ`.
- Also include `Work Ref` values beginning with `ID` (Ideas), but only when that Idea has not been converted to a Project — i.e., its `Work Ref` does not appear as the `Idea ID` on any `PRJ` row. If a `PRJ` row references an `Idea ID`, that Idea has been converted to a Project; count only the `PRJ` row, not the original Idea row, so the same initiative is never double-counted.
- Exclude records whose `Status` is `Cancelled`.
- Assignment date is the `Created` date.
- Calculate the SLA date as assignment date plus 14 calendar days for Projects (`Work Ref` starts with `PRJ`), or assignment date plus 28 calendar days for Ideas (`Work Ref` starts with `ID`).
- If that date is not Tuesday, move it forward to the next Tuesday.
- Approval on or before the expected date is `Met`.
- Missing approval after the expected date is `Not met`.
- Missing approval before or on the expected date is `Not yet due`.
- Use the current date for open-item classification unless an explicit as-of date is supplied.
- SLA percentage is `Met / (Met + Not met)`; exclude `Not yet due` from the denominator.

## Procedure

1. Provide the refreshed CSV path when running the script. The path may be absolute or relative to the current directory.
2. Run [run-sarb-sla.ps1](./scripts/run-sarb-sla.ps1).
3. Review `analyzed-2026-prj.csv` for project-level statuses and `summary-2026-prj.csv` for SA metrics.
4. Review `analyzed-2026-prj.csv` for project-level statuses and `summary-2026-prj.csv` for Solution Architect metrics.
Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-sarb-sla.ps1 -Source .\source.csv -AsOfDate 2026-08-24
```

The script creates the CSV outputs and `SARB-SLA-report.html`, including a readable horizontal Solution Architect chart, in the output folder. It does not send email or access authenticated SharePoint content; export the SharePoint list to CSV first.
