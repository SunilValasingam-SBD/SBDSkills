---
name: sa-project-status-review
description: 'Analyze one Solution Architect’s project records for missing SAD approvals and scenario publication status using a Work Queue CSV export.'
argument-hint: '[architect name, source CSV, optional reporting year]'
user-invocable: true
disable-model-invocation: false
---

# Solution Architect Project Status Review

Review project-level SAD approval and scenario publication status for a named Solution Architect.

## Procedure

1. Provide the Work Queue CSV path and the Solution Architect name exactly as it appears in the `Work Ref: Solution Architect` column.
2. Optionally provide a reporting year; if omitted, the current calendar year is used.
3. Run [run-sa-project-status.ps1](./scripts/run-sa-project-status.ps1):

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\run-sa-project-status.ps1 `
       -ArchitectName "Surname, Given" -Source .\source.csv -Year 2026
   ```

4. Review `sa-project-status.csv` for each included project, its SAD approval and publication statuses, and the derived status flags.

## Scope and outputs

- Includes `PRJ` project records created in the selected year and excludes records whose status is `Cancelled`.
- Matches the assigned architect name exactly.
- Marks a blank `SAD Approved` value as `Missing SAD Approval`; otherwise, it is `Present`.
- Treats a publication status beginning with `Published` as `Scenario published`; all other values are reported as `Not scenario published`.
- Writes `sa-project-status.csv` to the supplied output folder or the shared `SA\Outputs\sa-project-status-review` folder.
- The tool reads a user-supplied export locally. It does not authenticate to or retrieve data from SharePoint.
