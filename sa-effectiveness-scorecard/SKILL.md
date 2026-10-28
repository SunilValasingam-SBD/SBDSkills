---
name: sa-effectiveness-scorecard
description: 'Generate an interactive "SA Effectiveness Scorecard" per Solution Architect, combining a Worklist CSV export with a Timesheet Excel export. Ranks SAs by the % of their time spent on tangible delivery work (Project/Idea/TechPass) vs. internal/NPD work (SA NPD/NPD, tracked separately) vs. admin/overhead (Admin-Other/Training), with click-to-drill-down into underlying Worklist items. Use when asked about SA effectiveness, productivity, brand value, or "how much real work" each SA does vs admin/overhead.'
argument-hint: 'Provide the Worklist CSV and Timesheet Excel source paths when prompted'
user-invocable: true
disable-model-invocation: false
---

# SA Effectiveness Scorecard

Answers a different question than `sa-workload-dashboard` or `sarb-sla-reasoning`:
not "is SLA being met" or "how many hours per category", but **"how much of
each SA's time is genuinely tangible, delivery-facing work vs. admin/overhead,
and who should be recognized for doing real work rather than just staying
busy with tickets/admin?"**

## The three tiers

Every classification is rolled up into exactly one of three tiers, each with
a distinct color and a clickable legend tag so there's no ambiguity reading
the chart:

- 🟢 **Tangible / Delivery** — Project (actual timesheet hours) + Idea
  (imputed) + TechPass (imputed). Real deliverables with business value —
  this is the "brand value" / productive-work tier.
- 🔵 **Internal / NPD** — SA NPD (imputed) + NPD (imputed). Legitimate SA
  work that doesn't go through SARB/SLA tracking — kept as its **own** tier
  on purpose, so it's never silently folded into either "productive" or
  "overhead" and misread either way.
- 🟠 **Admin / Overhead** — Admin/Other (netted of all imputed hours below,
  to avoid double counting) + Training & Development. Time Off is excluded
  entirely — it isn't work time at all.

**Effectiveness Score** (the headline ranking metric) = Tangible hours /
(Tangible + Internal/NPD + Admin) hours, as a %. SAs are ranked
highest-to-lowest in both the leaderboard chart and the scorecard table.

## Why imputation is needed

Idea, TechPass, SA NPD, and NPD Worklist items may have no dedicated Project ID
in a timesheet system — that time can be logged under an administrative
project alongside real admin work and time off, so it may be invisible unless
imputed from the Worklist. This reuses and extends the
imputation approach from `sa-workload-dashboard` (which only imputes
Idea/TechPass) to also impute **SA NPD and NPD** hours:

- Idea: `--idea-hours-per-week` hours/week while active (default 1).
- TechPass: `--techpass-hours-per-week` hours/week while active (default 2).
- SA NPD: `--sa-npd-hours-per-week` hours/week while active (default 1).
- NPD: `--npd-hours-per-week` hours/week while active (default 1).
- "Weeks" = the distinct `Week Start` values present in the Timesheet export.
- An item is "active" in a week if its `Created` date is on/before that
  week's start, and — if it reached a terminal status (`Complete`/
  `Cancelled`) — its `Modified` date is on/after that week's start.
- All four imputed totals are subtracted from that SA's Admin/Other hours
  (floored at 0) to avoid double counting.

## Required Inputs (always ask for both, every run)

1. **Worklist CSV export** — must contain: `Work Ref`, `Work Ref: Solution
   Architect`, `Status`, `Created`, `Modified`.
2. **Timesheet Excel export** — must contain (after its header row):
   `Resource Name`, `Project ID`, `Timesheet Line Class`, `Hours Reported`,
   ideally also `Week Start`.

Do not reuse a prior run's paths without confirming they're still current.

## Procedure

1. Ask the user for the Worklist CSV path and the Timesheet Excel path.
2. Run the bundled generator:

   ```powershell
   py .\sa_effectiveness_scorecard.py --worklist "<worklist csv path>" --timesheet "<timesheet xlsx path>"
   ```

   Override defaults only if the user explicitly asks:
   - `--exclude-sa "<Last, First>"` — exclude a specific SA (repeatable;
     omitted means no exclusions).
   - `--idea-hours-per-week`, `--techpass-hours-per-week`,
     `--sa-npd-hours-per-week`, `--npd-hours-per-week` — imputed rates.
   - `--timeoff-project-id` — organization-specific Project ID representing
     time off/admin rows (no default).
   - `--no-net-admin` — don't subtract imputed hours from Admin/Other.
   - `--output-dir "<path>"` — output location (default: this skill's folder).

3. Open `sa-effectiveness-scorecard.html`:
   - **Leaderboard** — ranked 100%-ish stacked horizontal bar per SA
     (Tangible/Internal/Admin), sorted by Tangible % descending.
   - **Scorecard table** — one row per SA with hours + % for each tier and
     an Effectiveness Score pill (green ≥50%, amber 25–49%, red <25%).
     Click any row to drill into that SA's underlying Worklist items,
     each tagged with its tier.
4. Review `sa-effectiveness-summary.csv` for further reporting.

## Rules

- Source files are never modified.
- SAs appearing in only one source still appear in the output (zero-filled
  on the side they're missing from).
- Unclassified/unrecognized `Work Ref` prefixes are kept visible in the
  drill-down but are not scored into any tier (data-quality signal, not
  silently dropped or guessed).
