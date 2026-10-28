# Solution Architect Project Status Review

This skill reviews project records for a selected Solution Architect and highlights missing SAD approvals and unpublished scenarios.

## Input

A CSV export of the Solution Architecture Work Queue with these columns:

- `Created`
- `Work Ref`
- `Status`
- `Work Ref: Solution Architect`
- `SAD Approved`
- `Publication Status`

## Run

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-sa-project-status.ps1 `
    -ArchitectName "Surname, Given" `
    -Source .\work-queue.csv `
    -Year 2026 `
    -OutputDirectory .\output
```

`-Source` and `-OutputDirectory` are optional. If no source is supplied, the newest CSV in the shared `SA\Inputs` folder is used. If no output directory is supplied, the report is written under `SA\Outputs\sa-project-status-review`. `-Year` defaults to the current year.

## Output

`sa-project-status.csv` contains the matching project rows, source approval/publication values, and derived SAD approval and scenario publication status.

The script processes the supplied export locally. It does not connect to SharePoint or transmit the data.
