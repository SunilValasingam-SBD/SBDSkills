# Divestiture Separation Plan

## Executive Summary

In a divestiture, understanding every system dependency tied to the applications in scope is critical to planning a clean and risk-free separation. This capability takes the application inventory and disposition plan for the applications involved in the divestiture and produces a complete separation plan from a Solution Architecture perspective — showing, for each application, what disposition applies (e.g. transfer, retain, retire, carve-out) and what integrations must be addressed as a result, especially shared/dependent systems that complicate a clean carve-out. This gives the divestiture team an immediate, reliable basis for separation planning, stakeholder communication, and risk assessment — without manual, error-prone discovery across multiple teams and sources.

## Business Process Flow

1. The business provides two source files it already has on hand: the **IT Application Inventory** (every in-scope application, with its agreed disposition — e.g. Carve Out, Direct Transfer, Lift & Shift — and whether it's Shared or Dedicated), and the **Application Integration catalog** (every known system-to-system interface, by application).
2. The tool matches the two files application-by-application (by internal application ID, not by name, since naming often differs between exports) and evaluates each application's disposition against its integration footprint.
3. It automatically applies Solution Architecture judgment: it recognizes that a "Carve Out" tied to a data-extract/shared setup means the application and its integrations stay in place (only a one-time data extract goes to the buyer), while only applications that are genuinely transferring or being lifted-and-shifted onto dedicated infrastructure carry real integration cut-over risk.
4. It flags any logical conflicts (e.g. an application marked for physical relocation that is still shared with other retained applications) and any application missing a disposition decision, so these can be escalated for a business decision rather than guessed at.
5. It produces two deliverables: a detailed **Excel separation plan workbook** for full data review, and a **PowerPoint executive summary** that calls out every application that is actually relocating/transferring, with its full integration partner list spelled out on-slide — so stakeholders don't need to dig through the Excel file for the headline risks.
6. The user reviews the console summary counts (in-scope applications, applications truly relocating, open items, conflicts), then shares the Excel workbook with the detailed team and the PowerPoint with leadership/the divestiture steering committee.

## Technical Flow

- The user supplies the two source workbook paths; the model does not guess filenames since two ambiguous candidate files may exist in the shared Inputs folder.
- The user (via the model) runs `divestiture_separation_plan.py --inventory "<path>" --integrations "<path>" [--output-dir "<path>"] [--excel-output "<name>.xlsx"] [--ppt-output "<name>.pptx"]`. Both `--inventory` and `--integrations` are required, explicit arguments — there is no auto-detection of the newest file in the shared Inputs folder for this skill, precisely because two plausible workbooks can coexist there.
- `--output-dir` defaults to the shared `SA\Outputs\divestiture-separation-plan\` folder, resolved at argument-parse time via `_shared\sa_paths.py`'s `output_dir_for("divestiture-separation-plan")` helper (which creates the folder if it doesn't exist). An explicit `--output-dir` overrides this.
- **Inventory parsing**: the script detects the header row and resolves each needed column (Product Name, AppID, disposition, disposition method, Shared/Dedicated, etc.) by fuzzy, rule-based header-text matching (`FIELD_RULES`), not fixed column positions, so differently-laid-out exports still parse correctly.
- **Integration catalog parsing**: reads the `Summary` sheet (application → sheet name/interface count) plus each application's own sheet (`Data Object`, `Initiator App ID`/`Initiator App`, `Terminator App ID`/`Terminator`, `Interfaces`, etc.).
- **Matching**: applications are joined between the two workbooks strictly by AppID (not name).
- **Classification logic**: for each in-scope application the script derives SA guidance from the combination of disposition + disposition method + Shared/Dedicated — e.g. Carve Out + Data Extract + Shared = stays with SBD, no integration rebuild; Direct Transfer or (Lift & Shift + Dedicated) = physically relocates, flagged as a "Fully Transferring Application"; Lift & Shift + Shared = a logical conflict, flagged for business clarification; no disposition set = an Open Item.
- **Integration scope rule**: an integration is only reported when both the home application and its partner are present in the inventory (matched by AppID); partners outside the inventory are silently excluded rather than guessed at.
- **Excel output** (`openpyxl`): writes `Summary`, `Glossary`, `Separation Plan`, `Fully Transferring Applications`, `Open Items`, and `Integration Detail` tabs with styled headers/fills/borders.
- **PowerPoint output** (`python-pptx`): builds an overview slide, a slide listing every fully-transferring application (with its integration status called out explicitly, even when zero partners exist), a summary table slide, and one fully self-contained detail slide per application that has an in-inventory integration partner — every partner row is listed directly on the slide (no "see Excel" deferral). Table row heights are explicitly set per row count/density to avoid `python-pptx`'s default even-stretch row-height behavior.
- Final output file paths: `args.output_dir / <excel filename>` and `args.output_dir / <ppt filename>` unless an absolute path is given for either.

## Inputs

- **`--inventory`** (required) — IT Application Inventory workbook: `Product Name`, `AppID`, a disposition column (e.g. `SBD Proposed Disposition`), a disposition-method column, and a Shared/Dedicated column.
- **`--integrations`** (required) — Application Integration catalog workbook: a `Summary` sheet plus one sheet per application with `Data Object`, `Initiator App ID`/`Initiator App`, `Interfaces`, `Terminator App ID`/`Terminator` columns.
- **`--output-dir`** (optional) — destination folder; defaults to the shared `SA\Outputs\divestiture-separation-plan\` folder.
- **`--excel-output`** (optional) — Excel filename override; default `Divestiture-SA-Separation-Plan.xlsx`.
- **`--ppt-output`** (optional) — PowerPoint filename override; default `Divestiture-SA-Separation-Plan-Executive-Summary.pptx`.

## Outputs

- **`Divestiture-SA-Separation-Plan.xlsx`** — full separation plan workbook with `Summary`, `Glossary`, `Separation Plan`, `Fully Transferring Applications`, `Open Items`, and `Integration Detail` tabs covering every in-scope application.
- **`Divestiture-SA-Separation-Plan-Executive-Summary.pptx`** — executive-ready slide deck: overview/key observations, fully-transferring-applications slide, integration-partner summary table, and one detail slide per application with an in-inventory integration partner.
