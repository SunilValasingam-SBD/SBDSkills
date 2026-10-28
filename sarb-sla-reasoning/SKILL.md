---
name: sarb-sla-reasoning
description: 'Explain SARB SLA misses, not just measure them: classify logged reasons as business-caused vs ambiguous vs undocumented, compute Gross vs Controllable SLA % per Solution Architect and per Program Owning Portfolio, and produce an interactive click-to-drill-down HTML dashboard. Use when asked why SLA was missed, whether reasons are reasonable, whether SLA policy should change per portfolio, or for a fairer SA performance measure.'
argument-hint: 'Provide the refreshed Work Queue CSV path'
user-invocable: true
disable-model-invocation: false
---

# SARB SLA Reasoning Analysis

Goes one level deeper than the `sarb-sla-report` skill. That skill tells you
*whether* SLA was met; this skill explains *why* it was missed, whether that
reason is reasonable, and whether a Solution Architect's or portfolio's low
SLA % is actually about things outside their control.

## Scope rules (same base filters as sarb-sla-report)

- Only `Work Ref` values starting with `PRJ`, or standalone `ID` (Idea)
  values that have not been converted to a Project (i.e. do not appear as
  the `Idea ID` on any `PRJ` row).
- **Explicitly excludes all `TechPASS`, `NPD`, and `SA`-prefixed `Work Ref`
  rows** — these are not Projects or Ideas and must never be counted.
- Only records whose `Created` date falls in the target year (default:
  current calendar year — override with `--year` if analyzing a prior year).
- Excludes `Status` = `Cancelled`.
- SLA date = `Created` + 14 calendar days for Projects (`PRJ`), or `Created` + 28
  calendar days for Ideas (`ID`), rolled forward to the next Tuesday if not
  already a Tuesday. Met/Not met/Not yet due classification matches
  `sarb-sla-report` exactly.

## Reason classification

Every "Not met" item's `SARB Reason` is classified into one of three
categories (edit the `REASON_CATEGORIES` dict in
[sarb_sla_reasoning.py](./sarb_sla_reasoning.py) as new reason text appears
in the Work Queue):

- **Business-caused** (`Pending BRD`, `Scope not defined`, `Project under NDA`,
  `Future start date`, `SA change`, `Bring back to SARB`) — a genuine
  upstream dependency outside the SA's control, including an SA
  reassignment (the new SA needs time to ramp up) and SARB directing the
  item to come back for a later presentation after the SARB team's own
  feedback required more architectural/business work. Excluded from the
  **Controllable SLA %** denominator/numerator, the same way "Not yet due"
  is excluded from Gross SLA %.
- **Ambiguous** (`Initial review`, or any reason text not yet in the
  lookup) — still counted against Controllable SLA (treated conservatively
  as controllable) but flagged separately for review, since these reasons
  don't clearly explain the delay.
- **Undocumented** (no reason logged at all) — still counted against
  Controllable SLA, and reported as its own metric, since it's a
  documentation/governance gap, not evidence of actual performance either
  way.

## Procedure

1. Get the refreshed Work Queue CSV path (SharePoint export) from the user.
2. Run the script:

   ```powershell
   py .\sarb_sla_reasoning.py --input "<path to Work Queue CSV>"
   ```

   Optional flags:
   - `--year 2026` — override the Created-year filter (default: current year).
   - `--as-of 2026-09-15` — analyze as of a historical date instead of today.
   - `--output-dir "<folder>"` — write outputs elsewhere (default: this skill folder).

3. Open `sarb-sla-reasoning.html` — four clickable charts plus four
   clickable summary cards:
   - **SLA % by Solution Architect** (Gross vs Controllable, grouped bars)
   - **SLA % by Program Owning Portfolio** (Gross vs Controllable, grouped bars)
   - **Missed-SLA reasons** (count per reason, color-coded by category)
   - **SLA % by Solution Architect — Projects only** (same Gross vs
     Controllable grouped bars, but excludes standalone Ideas so the
     28-day Idea SLA window never skews a Solution Architect's Project
     performance) — last chart on the dashboard
   - **Not-met items (all)**, **...business-caused**, **...ambiguous**, and
     **...no reason logged** cards at the top — click any of them for a
     single entry point into every not-met project in that scope, without
     needing to pick a specific SA/portfolio/reason bar first.

   Click any bar, or any of those four cards, to drill into the exact
   not-met projects behind it — Project ID, name, SA, portfolio, logged
   reason, and category. Drill-downs from the **cards** (whole-team or
   category-wide views) are additionally grouped into one mini-table per
   Solution Architect, each with its own "N not-met project(s)" header,
   sorted by most not-met projects first — since those views span multiple
   SAs at once. Drill-downs from a single SA/portfolio/reason bar stay as
   one flat table, since they already cover just that one SA/portfolio/reason.
4. Review the CSV outputs for further analysis or reporting:
   - `sa-sla-reasoning-summary.csv`
   - `portfolio-sla-reasoning-summary.csv`
   - `reason-sla-reasoning-summary.csv`
   - `project-only-sla-reasoning-summary.csv` (same per-SA summary, Projects only)
   - `not-met-detail.csv` (every not-met project with its reason/category)

## Interpreting results

- A Solution Architect or portfolio with a big gap between Gross and
  Controllable SLA % is mostly blocked by business dependencies (missing
  BRD/scope, NDA holds) — a process/SLA-policy question, not a performance
  one.
- A high "Undocumented %" means the real cause is unknown — fix the
  documentation discipline before drawing conclusions about that SA/portfolio.
- A portfolio where misses are dominated by one business-caused reason (e.g.
  `Project under NDA` for M&A) is a candidate for its own SLA track rather
  than being held to the standard 14-day/Tuesday rule.
