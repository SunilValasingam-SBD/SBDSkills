---
name: divestiture-separation-plan
description: 'Build a Solution Architecture perspective on an application divestiture: match an IT Application Inventory (with SBD Proposed Disposition) against an Application Integration catalog, classify what actually happens to each application, flag disposition/sharing conflicts, call out every application that actually relocates/transfers ownership, and produce a Separation Plan Excel workbook plus a self-contained PowerPoint (overview + a fully-transferring-applications slide + one detail slide per application with any in-inventory integration partners). Use when the user has a divestiture/carve-out application inventory and an integration export and wants a solution-architecture separation plan, without needing to open either source file.'
argument-hint: 'Provide the IT Application Inventory workbook and the Integration catalog workbook'
user-invocable: true
disable-model-invocation: false
---

# Divestiture Separation Plan

Produces a Solution Architecture-ready separation plan for an application
divestiture/carve-out from two source workbooks: an **IT Application
Inventory** (one row per application, with a disposition column) and an
**Application Integration catalog** (one Summary sheet + one sheet per
application listing its Initiator/Terminator interfaces).

The goal is a single pass that tells a Solution Architect, for every
in-scope application: what is actually happening to it, what to do about
its integrations, and which cases need business clarification — with no
follow-up trip back into either source spreadsheet required.

## Required inputs

- **Inventory workbook**: a sheet with at least `Product Name` and `AppID`
  columns, plus a disposition column (e.g. `SBD Proposed Disposition`), a
  disposition-method column (e.g. `SBD Proposed Disposition Method`), and a
  Shared/Dedicated column. Column positions and exact header wording may
  vary between exports — the script detects the header row and resolves
  each field by fuzzy header-text matching, not fixed column numbers.
- **Integrations workbook**: a `Summary` sheet (`Application Name`,
  `Integration Count`, `Sheet Name`) plus one sheet per application with
  `Data Object`, `Display Text`, `Initiator App ID`, `Initiator App`,
  `Interfaces`, `Terminator App ID`, `Terminator` columns.
- Applications are matched between the two workbooks **by AppID**, not by
  name — naming conventions commonly differ between the two exports for the
  same application.

## Key domain logic (do not re-derive from scratch each time)

- **"Carve Out from SBD Environment" does not mean the application
  leaves SBD.** When paired with Disposition Method = `Data Extract` and
  Shared/Dedicated = `Shared`, the application stays with SBD as shared
  infrastructure; only a data extract of the divesting business's data is
  delivered to the buyer. Only applications that are actually
  **Direct Transfer** or **Dedicated Lift & Shift** physically relocate —
  those are the only ones with real integration rebuild/retire risk.
- **Conflict flag**: a `Lift & Shift` application marked `Shared` is a
  logical contradiction (a shared app can't physically relocate without
  affecting other SBD consumers still using it) — report these explicitly
  as open conflicts needing business clarification, never guess an answer.
- **Integration scope is in-inventory-only.** Only report an integration
  where **both** the home application and its partner are present in the
  inventory (matched by AppID). Do not infer anything about counterparts
  that aren't in the inventory (no "this must be an SBD-retained system"
  guesses) — if the partner isn't in the inventory, the integration is
  simply out of scope and silently excluded.
- Applications with no disposition set are reported as **Open Items**,
  not guessed at or dropped silently.
- **Applications that actually relocate or transfer ownership** (Direct
  Transfer, or Lift & Shift + Dedicated) are always surfaced as their own
  named category — a dedicated Excel tab (`Fully Transferring
  Applications`) and a dedicated PPT slide — regardless of whether the
  integration catalog has any data for them. If an application in this
  category has zero in-inventory integration partners, say so explicitly
  (e.g. "no sheet found for this AppID in the integration catalog —
  confirm with the application owner") rather than omitting it; these are
  the highest-risk applications in the plan and must never be silently
  dropped just because they don't also appear in the integration-partner
  list.

## Procedure

1. Get both source workbook paths from the user (inventory + integrations).
   Do not assume filenames — ask if not provided.
2. Run the script:

   ```powershell
   py .\divestiture_separation_plan.py --inventory "<inventory workbook path>" --integrations "<integrations workbook path>" --output-dir "<output folder>"
   ```

   Optional flags:
   - `--excel-output "<filename>.xlsx"` — override the Excel output filename.
   - `--ppt-output "<filename>.pptx"` — override the PPT output filename.
   - Omit `--output-dir` to write next to the script.

3. Review the console summary: in-scope application count, count of
   applications that actually relocate/transfer ownership, count of
   applications with in-inventory integration partners, open-item count,
   and any disposition/sharing conflicts.
4. Open the Excel workbook (`Summary`, `Glossary`, `Separation Plan`,
   `Fully Transferring Applications`, `Open Items`, `Integration Detail`
   tabs) for the full data set.
5. Open the PowerPoint: slide 1 overview/key observations, slide 2 every
   application that actually relocates/transfers ownership (with its
   integration status called out explicitly, even when zero), slide 3 a
   summary table of every application with an in-inventory integration
   partner, then one fully self-contained detail slide per such
   application — meta line, SA guidance, and a complete partner table
   (no "see Excel workbook" deferral and no row truncation; every partner
   is listed directly on the slide).

## Table formatting rule (PPT)

python-pptx stretches a table's rows evenly across whatever height is
passed to `add_table(...)`, which makes a 2-row table absurdly tall. The
script fixes this by explicitly setting `table.rows[i].height` after
creation, scaling between ~0.27in (many rows) and ~0.5in (few rows) based
on row count, with a smaller font (9.5pt) for dense tables. Keep this
fix if the script is modified — do not let row heights reset to
python-pptx's default auto-stretch behavior.

## Interpreting results

- Most "Carve Out" + "Data Extract" + "Shared" applications need no
  integration rebuild work — only a confirmation that the data extract
  feed itself is complete and reliable for the buyer (a data lineage
  check, not an integration migration).
- Applications flagged as a Lift & Shift / Shared conflict need a business
  decision before architecture work can proceed (e.g. split into a
  dedicated instance, or reclassify the disposition).
- Open Items (no disposition) block separation planning entirely for that
  application and should be escalated for a disposition decision first.
- Applications in the "Fully Transferring Applications" category are the
  highest-risk group — they physically move. If one shows zero
  in-inventory integration partners, that's usually a gap in the source
  data (missing AppID, or no sheet in the integration catalog), not proof
  the application has no integrations — confirm directly with the
  application owner before assuming it's integration-free.
