"""SARB SLA Reasoning Analysis.

Goes one level deeper than a plain Met/Not-met SLA report: it classifies the
*reason* logged against each missed-SLA item, separates reasons that are
genuinely outside the Solution Architect's control from ones that are not a
real explanation at all (or missing entirely), and reports two SLA numbers
per Solution Architect / Portfolio:

  - Gross SLA %         same definition as the base sarb-sla-report skill
                         (Met / (Met + Not met), all Not-met items counted)
  - Controllable SLA %  Met / (Met + Not-met-for-a-controllable-reason)
                         i.e. excludes Not-met items whose logged reason is
                         a business/process dependency outside the SA's
                         control (see REASON_CATEGORIES below)

The output is a single interactive HTML dashboard (bar charts for SA and
Portfolio performance, plus a reason-category breakdown) where every bar is
clickable and shows the contributing project list - including the exact
logged reason - as a drill-down table. This lets anyone asking "why is SA X
at 30%?" or "why is this portfolio low?" get the underlying project list and
reasons immediately, without re-running anything.

Reason classification is intentionally a simple, editable lookup
(REASON_CATEGORIES below) rather than a hardcoded list of SA/portfolio
exceptions, so new reason text introduced in the Work Queue can be
re-classified by editing one dictionary.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go

# ---------------------------------------------------------------------------
# Reason classification
# ---------------------------------------------------------------------------
# Category meanings:
#   business_caused   - A genuine upstream/process dependency outside the SA's
#                        control (missing BRD, undefined scope, confidentiality
#                        hold, a project that has not actually started yet, an
#                        SA reassignment that forces the new SA to spend time
#                        getting up to speed, or SARB asking for the item to
#                        come back after a prior presentation because the SARB
#                        team's own feedback required more architectural/
#                        business work before it could be finalized).
#                        Excluded from the Controllable SLA denominator/numerator
#                        the same way "Not yet due" is excluded from Gross SLA.
#   ambiguous         - A logged reason that may or may not reflect an SA-caused
#                        delay (e.g. rework loops, vague stage labels). Counted
#                        as controllable (conservative) but flagged separately
#                        so it can be reviewed/re-coded rather than ignored.
#   undocumented      - No reason was logged at all. Always counted as
#                        controllable (most conservative), and surfaced as its
#                        own metric because it is a data-quality/governance gap,
#                        not evidence either way about performance.
# Unrecognized reason text (not in this dict) falls back to "ambiguous" so it
# is visible in the dashboard rather than silently mis-classified.
REASON_CATEGORIES: dict[str, str] = {
    "pending brd": "business_caused",
    "scope not defined": "business_caused",
    "project under nda": "business_caused",
    "future start date": "business_caused",
    "sa change": "business_caused",
    "bring back to sarb": "business_caused",
    "initial review": "ambiguous",
}

CATEGORY_LABELS = {
    "business_caused": "Business-caused (outside SA control)",
    "ambiguous": "Ambiguous / needs clearer logging",
    "undocumented": "No reason logged",
}


def classify_reason(reason: str) -> str:
    if not reason or not str(reason).strip():
        return "undocumented"
    key = str(reason).strip().lower()
    return REASON_CATEGORIES.get(key, "ambiguous")


# ---------------------------------------------------------------------------
# Core SLA calculation (same rules as the sarb-sla-report skill)
# ---------------------------------------------------------------------------
def expected_sarb_date(created: datetime, work_ref: str = "") -> datetime:
    # Projects (PRJ) get a 14-calendar-day SLA window; Ideas (ID) get 28 calendar days.
    days = 28 if str(work_ref).strip().upper().startswith("ID") else 14
    date = created + timedelta(days=days)
    while date.weekday() != 1:  # Monday=0 ... Tuesday=1
        date += timedelta(days=1)
    return date


def is_blank(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    text = str(value).strip()
    return text == "" or text.lower() in ("nan", "none", "nat")


def load_source(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in (".xlsx", ".xls"):
        frame = pd.read_excel(path, dtype=str)
    else:
        frame = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
    frame.columns = [str(c).strip() for c in frame.columns]
    return frame


def build_analysis(frame: pd.DataFrame, year: int, as_of: datetime) -> pd.DataFrame:
    required = ["Created", "Work Ref", "Idea ID", "Status", "SARB Approval Date",
                "Work Ref: Solution Architect"]
    missing = [c for c in required if c not in frame.columns]
    if missing:
        raise ValueError(f"Missing required column(s): {', '.join(missing)}")

    df = frame.copy()
    df["Created_dt"] = pd.to_datetime(df["Created"], errors="coerce")
    work_ref_upper = df["Work Ref"].fillna("").astype(str).str.strip().str.upper()

    # Ideas already converted to a Project are referenced via Idea ID on a PRJ row.
    converted_idea_ids = set(
        df.loc[work_ref_upper.str.startswith("PRJ"), "Idea ID"]
        .dropna().astype(str).str.strip()
    )
    converted_idea_ids.discard("")

    is_prj = work_ref_upper.str.startswith("PRJ")
    is_standalone_idea = work_ref_upper.str.startswith("ID") & ~df["Work Ref"].astype(str).str.strip().isin(converted_idea_ids)
    # Explicitly ignore anything that is not a Project or a standalone Idea
    # (e.g. TechPASS / NPD / SA work-ref types), and only this reporting year.
    in_scope = (is_prj | is_standalone_idea) & (df["Created_dt"].dt.year == year)
    not_cancelled = df["Status"].fillna("").astype(str).str.strip().str.lower() != "cancelled"

    df = df[in_scope & not_cancelled].reset_index(drop=True)

    df["Expected SARB Date"] = [
        expected_sarb_date(created, work_ref)
        for created, work_ref in zip(df["Created_dt"], df["Work Ref"])
    ]
    approval_dt = pd.to_datetime(df["SARB Approval Date"], errors="coerce")

    def classify_row(row, approval):
        expected = row["Expected SARB Date"]
        if pd.notna(approval):
            return "Met" if approval.date() <= expected.date() else "Not met"
        if as_of.date() > expected.date():
            return "Not met"
        return "Not yet due"

    df["SLA Status"] = [classify_row(row, approval_dt.iloc[i]) for i, row in df.iterrows()]

    reason_col = "SARB Reason" if "SARB Reason" in df.columns else None
    df["Reason"] = df[reason_col].fillna("").astype(str).str.strip() if reason_col else ""
    df.loc[df["Reason"] == "", "Reason"] = ""
    df["Reason Category"] = df["Reason"].apply(classify_reason)
    # Only "Not met" rows get a meaningful reason category for SLA purposes.
    df.loc[df["SLA Status"] != "Not met", "Reason Category"] = ""
    df["Controllable Not Met"] = (df["SLA Status"] == "Not met") & (df["Reason Category"] != "business_caused")

    df["Solution Architect"] = df["Work Ref: Solution Architect"].fillna("").astype(str).str.strip()
    df.loc[df["Solution Architect"] == "", "Solution Architect"] = "(Unassigned)"

    portfolio_col = "Work Ref: Program Owning Portfolio"
    if portfolio_col in df.columns:
        df["Portfolio"] = df[portfolio_col].fillna("").astype(str).str.strip()
        df["Portfolio"] = df["Portfolio"].str.replace(r"^/Active SBD(/SBD/SBD)?/?", "", regex=True)
        df.loc[df["Portfolio"] == "", "Portfolio"] = "(No portfolio)"
    else:
        df["Portfolio"] = "(No portfolio)"

    df["Project ID"] = df["Work Ref"].fillna("").astype(str).str.strip()
    name_col = "Work Ref: Name"
    df["Project Name"] = df[name_col].fillna("").astype(str).str.strip() if name_col in df.columns else ""
    return df


# ---------------------------------------------------------------------------
# Summaries
# ---------------------------------------------------------------------------
def build_group_summary(df: pd.DataFrame, group_col: str) -> pd.DataFrame:
    rows = []
    for key, group in df.groupby(group_col):
        met = int((group["SLA Status"] == "Met").sum())
        not_met_gross = int((group["SLA Status"] == "Not met").sum())
        not_met_ctrl = int(group["Controllable Not Met"].sum())
        not_yet_due = int((group["SLA Status"] == "Not yet due").sum())
        gross_pct = round(100 * met / (met + not_met_gross), 1) if (met + not_met_gross) else None
        ctrl_pct = round(100 * met / (met + not_met_ctrl), 1) if (met + not_met_ctrl) else None
        undocumented = int(((group["SLA Status"] == "Not met") & (group["Reason Category"] == "undocumented")).sum())
        rows.append({
            group_col: key,
            "Total": len(group),
            "Met": met,
            "Not Met (Gross)": not_met_gross,
            "Not Met (Controllable)": not_met_ctrl,
            "Not Yet Due": not_yet_due,
            "Gross SLA %": gross_pct,
            "Controllable SLA %": ctrl_pct,
            "Undocumented Misses": undocumented,
            "Undocumented %": round(100 * undocumented / not_met_gross, 1) if not_met_gross else None,
        })
    return pd.DataFrame(rows).sort_values("Controllable SLA %", ascending=True, na_position="first")


def build_reason_summary(df: pd.DataFrame) -> pd.DataFrame:
    not_met = df[df["SLA Status"] == "Not met"].copy()
    not_met["Reason Label"] = not_met["Reason"].replace("", "No reason logged")
    rows = []
    for reason, group in not_met.groupby("Reason Label"):
        category = group["Reason Category"].iloc[0] if not group.empty else "ambiguous"
        rows.append({
            "Reason": reason,
            "Category": CATEGORY_LABELS.get(category, category),
            "Count": len(group),
        })
    return pd.DataFrame(rows).sort_values("Count", ascending=False)


def project_rows(subset: pd.DataFrame) -> list[dict]:
    return [
        {
            "Project ID": r["Project ID"] or "(blank)",
            "Project Name": r["Project Name"],
            "Solution Architect": r["Solution Architect"],
            "Portfolio": r["Portfolio"],
            "Reason": r["Reason"] or "No reason logged",
            "Category": CATEGORY_LABELS.get(r["Reason Category"], r["Reason Category"]) if r["Reason Category"] else "",
            "Status": r["Status"] if "Status" in r else "",
        }
        for _, r in subset.iterrows()
    ]


def build_drilldown(df: pd.DataFrame) -> dict:
    drilldown: dict[str, list[dict]] = {}
    not_met = df[df["SLA Status"] == "Not met"].copy()
    not_met["Reason Label"] = not_met["Reason"].replace("", "No reason logged")

    for sa, group in df.groupby("Solution Architect"):
        drilldown[f"SA||{sa}"] = project_rows(group[group["SLA Status"] == "Not met"])
    for portfolio, group in df.groupby("Portfolio"):
        drilldown[f"PORTFOLIO||{portfolio}"] = project_rows(group[group["SLA Status"] == "Not met"])
    for reason, group in not_met.groupby("Reason Label"):
        drilldown[f"REASON||{reason}"] = project_rows(group)

    # Overall "everything not met" view, and one view per reason category,
    # so the headline Not-met count/category cards are drillable too.
    drilldown["ALL||Not Met"] = project_rows(not_met.sort_values(["Solution Architect", "Project ID"]))
    for category_key, category_label in CATEGORY_LABELS.items():
        subset = not_met[not_met["Reason Category"] == category_key].sort_values(["Solution Architect", "Project ID"])
        drilldown[f"CATEGORY||{category_label}"] = project_rows(subset)

    # Projects-only view (excludes standalone Ideas) so a PRJ-specific
    # per-SA chart can drill down without mixing in Idea rows.
    projects_only = df[df["Project ID"].str.upper().str.startswith("PRJ")]
    for sa, group in projects_only.groupby("Solution Architect"):
        drilldown[f"SAPRJ||{sa}"] = project_rows(group[group["SLA Status"] == "Not met"])
    return drilldown


# ---------------------------------------------------------------------------
# HTML dashboard
# ---------------------------------------------------------------------------
def make_bar_figure(summary: pd.DataFrame, group_col: str, div_id: str, title: str) -> str:
    labels = summary[group_col].tolist()
    gross = summary["Gross SLA %"].fillna(0).tolist()
    ctrl = summary["Controllable SLA %"].fillna(0).tolist()

    figure = go.Figure()
    figure.add_bar(name="Gross SLA %", x=labels, y=gross, marker_color="#c24b3d",
                    text=gross, texttemplate="%{text:.1f}%", textposition="outside",
                    hovertemplate="%{x}<br>Gross SLA: %{y:.1f}%<extra></extra>")
    figure.add_bar(name="Controllable SLA %", x=labels, y=ctrl, marker_color="#1c8c72",
                    text=ctrl, texttemplate="%{text:.1f}%", textposition="outside",
                    hovertemplate="%{x}<br>Controllable SLA: %{y:.1f}%<extra></extra>")
    figure.update_layout(
        title=title,
        barmode="group",
        template="plotly_white",
        height=max(420, 90 + len(labels) * 10),
        font={"family": "Segoe UI, Aptos, sans-serif", "size": 13, "color": "#24323d"},
        yaxis={"title": "SLA %", "range": [0, 112]},
        xaxis={"tickangle": -30},
        margin={"l": 60, "r": 30, "t": 60, "b": 120},
        legend={"orientation": "h", "y": 1.08},
        uniformtext_minsize=9,
    )
    return figure.to_html(full_html=False, include_plotlyjs="cdn" if div_id == "saChart" else False, div_id=div_id)


def make_reason_figure(reason_summary: pd.DataFrame, div_id: str) -> str:
    colors = {
        CATEGORY_LABELS["business_caused"]: "#1c8c72",
        CATEGORY_LABELS["ambiguous"]: "#e08a3e",
        CATEGORY_LABELS["undocumented"]: "#c24b3d",
    }
    bar_colors = [colors.get(c, "#7a8a99") for c in reason_summary["Category"]]
    figure = go.Figure(
        data=go.Bar(
            x=reason_summary["Reason"],
            y=reason_summary["Count"],
            marker_color=bar_colors,
            customdata=reason_summary["Category"],
            text=reason_summary["Count"],
            texttemplate="%{text}",
            textposition="outside",
            hovertemplate="%{x}<br>Count: %{y}<br>%{customdata}<extra></extra>",
        )
    )
    figure.update_layout(
        title="Missed-SLA reasons (color = business-caused / ambiguous / undocumented)",
        template="plotly_white",
        height=420,
        font={"family": "Segoe UI, Aptos, sans-serif", "size": 13, "color": "#24323d"},
        yaxis={"title": "Not-met projects"},
        xaxis={"tickangle": -20},
        margin={"l": 60, "r": 30, "t": 60, "b": 120},
        uniformtext_minsize=9,
    )
    return figure.to_html(full_html=False, include_plotlyjs=False, div_id=div_id)


def build_html(sa_summary, portfolio_summary, reason_summary, project_summary, drilldown, team_stats, output_path: Path, as_of: datetime, year: int):
    sa_html = make_bar_figure(sa_summary, "Solution Architect", "saChart", "SLA % by Solution Architect (click a bar to drill down)")
    portfolio_html = make_bar_figure(portfolio_summary, "Portfolio", "portfolioChart", "SLA % by Program Owning Portfolio (click a bar to drill down)")
    reason_html = make_reason_figure(reason_summary, "reasonChart")
    project_html = make_bar_figure(
        project_summary, "Solution Architect", "projectChart",
        "SLA % by Solution Architect \u2014 Projects only, excludes Ideas (click a bar to drill down)"
    )
    drilldown_json = json.dumps(drilldown)
    business_caused_label = CATEGORY_LABELS["business_caused"]
    ambiguous_label = CATEGORY_LABELS["ambiguous"]
    undocumented_label = CATEGORY_LABELS["undocumented"]

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>SARB SLA Reasoning Analysis</title>
<style>
  body {{ font-family: 'Segoe UI', Aptos, sans-serif; margin: 24px; color: #24323d; background: #f6f3ee; }}
  h1 {{ margin-bottom: 4px; }}
  .muted {{ color: #5c6b73; margin-bottom: 20px; }}
  .cards {{ display: flex; gap: 12px; flex-wrap: wrap; margin: 20px 0 30px; }}
  .card {{ background: #fff; border: 1px solid #d9e2e5; border-radius: 8px; padding: 14px 18px; min-width: 160px; }}
  .card strong {{ display: block; font-size: 26px; margin-top: 4px; }}
  .card.clickable {{ cursor: pointer; transition: box-shadow 0.15s, border-color 0.15s; }}
  .card.clickable:hover {{ border-color: #264653; box-shadow: 0 2px 8px rgba(0,0,0,0.12); }}
  .card.clickable small {{ display: block; color: #5c6b73; margin-top: 2px; }}
  .legend-note {{ background: #fff; border: 1px solid #d9e2e5; border-radius: 8px; padding: 14px 18px; margin-bottom: 24px; font-size: 13px; line-height: 1.5; }}
  .legend-note span {{ display: inline-block; width: 12px; height: 12px; border-radius: 2px; margin-right: 6px; }}
  .chart-block {{ background: #fff; border: 1px solid #d9e2e5; border-radius: 8px; padding: 10px 16px 4px; margin-bottom: 28px; }}
  #drilldownPanel {{ margin-top: 10px; }}
  #drilldownTitle {{ font-size: 18px; font-weight: 600; margin-bottom: 12px; }}
  table {{ border-collapse: collapse; width: 100%; background: #ffffff; margin-top: 8px; }}
  th, td {{ border: 1px solid #d8dfdc; padding: 7px 10px; text-align: left; font-size: 12.5px; }}
  th {{ background: #264653; color: #ffffff; position: sticky; top: 0; }}
  tr:nth-child(even) td {{ background: #f1f5f4; }}
  .emptyHint {{ color: #6b7a75; font-style: italic; }}
  .drilldownScroll {{ max-height: 480px; overflow-y: auto; }}
  .saGroupHeader {{ background: #e8edec; border: 1px solid #c7d2cf; border-radius: 6px; padding: 6px 10px; margin: 16px 0 4px; font-weight: 600; color: #264653; }}
  .saGroupHeader:first-child {{ margin-top: 0; }}
</style>
</head>
<body>
<h1>SARB SLA Reasoning Analysis</h1>
<div class="muted">As of {as_of.strftime('%Y-%m-%d')} | Created in {year} | PRJ projects and standalone Ideas (not converted to a Project) | NPD / TechPASS / SA work refs and Cancelled items excluded</div>

<div class="cards">
  <div class="card">Items in scope<strong>{team_stats['total']}</strong></div>
  <div class="card">Gross SLA %<strong>{team_stats['gross_pct']}%</strong></div>
  <div class="card">Controllable SLA %<strong>{team_stats['ctrl_pct']}%</strong></div>
  <div class="card clickable" id="notMetCard">Not-met items (all)<strong>{team_stats['not_met']}</strong><small>Click to list every not-met project</small></div>
  <div class="card clickable" id="businessCausedCard">...business-caused<strong>{team_stats['business_caused']} ({team_stats['business_caused_pct']}%)</strong><small>Click to drill down</small></div>
  <div class="card clickable" id="ambiguousCard">...ambiguous<strong>{team_stats['ambiguous']} ({team_stats['ambiguous_pct']}%)</strong><small>Click to drill down</small></div>
  <div class="card clickable" id="undocumentedCard">...no reason logged<strong>{team_stats['undocumented']} ({team_stats['undocumented_pct']}%)</strong><small>Click to drill down</small></div>
</div>

<div class="legend-note">
  <b>Gross SLA %</b> = Met / (Met + Not met), counting every missed item. &nbsp;
  <b>Controllable SLA %</b> = Met / (Met + Not-met items whose reason is <i>not</i> business-caused) &mdash;
  excludes items blocked by a Pending BRD, undefined scope, an NDA hold, or a future start date, since those delays
  sit with the business/project, not the Solution Architect.<br>
  Reason colors: <span style="background:#1c8c72"></span>business-caused (excluded from Controllable SLA) &nbsp;
  <span style="background:#e08a3e"></span>ambiguous / needs clearer logging (still counted) &nbsp;
  <span style="background:#c24b3d"></span>no reason logged (still counted, flagged as a data-quality gap).
</div>

<div class="chart-block">{sa_html}</div>
<div class="chart-block">{portfolio_html}</div>
<div class="chart-block">{reason_html}</div>
<div class="chart-block">{project_html}</div>

<div id="drilldownPanel">
  <div id="drilldownTitle">Click any bar above, or a "Not-met" card, to list its projects and logged reasons.</div>
  <div class="drilldownScroll" id="drilldownTableContainer"><span class="emptyHint">No selection yet.</span></div>
</div>

<script>
  const DRILLDOWN_DATA = {drilldown_json};

  function buildTable(rows) {{
    if (!rows || rows.length === 0) {{
      return '<span class="emptyHint">No not-met projects for this selection.</span>';
    }}
    let html = '<table><thead><tr><th>Project ID</th><th>Project Name</th><th>Solution Architect</th><th>Portfolio</th><th>Reason</th><th>Category</th></tr></thead><tbody>';
    rows.forEach(function (row) {{
      html += '<tr><td>' + row['Project ID'] + '</td><td>' + row['Project Name'] + '</td><td>' + row['Solution Architect'] +
        '</td><td>' + row['Portfolio'] + '</td><td>' + row['Reason'] + '</td><td>' + row['Category'] + '</td></tr>';
    }});
    html += '</tbody></table>';
    return html;
  }}

  function buildTableGroupedBySA(rows) {{
    if (!rows || rows.length === 0) {{
      return '<span class="emptyHint">No not-met projects for this selection.</span>';
    }}
    const groups = {{}};
    const order = [];
    rows.forEach(function (row) {{
      const sa = row['Solution Architect'] || '(Unassigned)';
      if (!groups[sa]) {{ groups[sa] = []; order.push(sa); }}
      groups[sa].push(row);
    }});
    order.sort(function (a, b) {{ return groups[b].length - groups[a].length || a.localeCompare(b); }});
    let html = '';
    order.forEach(function (sa) {{
      const saRows = groups[sa];
      html += '<div class="saGroupHeader">' + sa + ' &mdash; ' + saRows.length + ' not-met project(s)</div>';
      html += '<table><thead><tr><th>Project ID</th><th>Project Name</th><th>Portfolio</th><th>Reason</th><th>Category</th></tr></thead><tbody>';
      saRows.forEach(function (row) {{
        html += '<tr><td>' + row['Project ID'] + '</td><td>' + row['Project Name'] +
          '</td><td>' + row['Portfolio'] + '</td><td>' + row['Reason'] + '</td><td>' + row['Category'] + '</td></tr>';
      }});
      html += '</tbody></table>';
    }});
    return html;
  }}

  function renderDrilldown(prefix, label) {{
    const key = prefix + "||" + label;
    const rows = DRILLDOWN_DATA[key] || [];
    document.getElementById('drilldownTitle').textContent = label + " - " + rows.length + " not-met project(s)";
    const groupBySA = (prefix === 'ALL' || prefix === 'CATEGORY');
    document.getElementById('drilldownTableContainer').innerHTML = groupBySA ? buildTableGroupedBySA(rows) : buildTable(rows);
  }}

  document.addEventListener('DOMContentLoaded', function () {{
    const sa = document.getElementById('saChart');
    const portfolio = document.getElementById('portfolioChart');
    const reason = document.getElementById('reasonChart');
    const projectChart = document.getElementById('projectChart');
    if (sa) sa.on('plotly_click', function (d) {{ renderDrilldown('SA', d.points[0].x); }});
    if (portfolio) portfolio.on('plotly_click', function (d) {{ renderDrilldown('PORTFOLIO', d.points[0].x); }});
    if (reason) reason.on('plotly_click', function (d) {{ renderDrilldown('REASON', d.points[0].x); }});
    if (projectChart) projectChart.on('plotly_click', function (d) {{ renderDrilldown('SAPRJ', d.points[0].x); }});

    const notMetCard = document.getElementById('notMetCard');
    if (notMetCard) notMetCard.addEventListener('click', function () {{ renderDrilldown('ALL', 'Not Met'); }});
    const businessCausedCard = document.getElementById('businessCausedCard');
    if (businessCausedCard) businessCausedCard.addEventListener('click', function () {{ renderDrilldown('CATEGORY', {business_caused_label!r}); }});
    const ambiguousCard = document.getElementById('ambiguousCard');
    if (ambiguousCard) ambiguousCard.addEventListener('click', function () {{ renderDrilldown('CATEGORY', {ambiguous_label!r}); }});
    const undocumentedCard = document.getElementById('undocumentedCard');
    if (undocumentedCard) undocumentedCard.addEventListener('click', function () {{ renderDrilldown('CATEGORY', {undocumented_label!r}); }});
  }});
</script>
</body>
</html>
"""
    output_path.write_text(html, encoding="utf-8")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def parse_args() -> argparse.Namespace:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
    from sa_paths import output_dir_for

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=None,
                         help="Path to the refreshed Work Queue CSV/Excel export. Defaults to the "
                              "newest .csv/.xlsx file in the shared SA\\Inputs folder.")
    parser.add_argument("--year", type=int, default=None, help="Created-year filter (default: current year).")
    parser.add_argument("--as-of", type=str, default=None, help="As-of date YYYY-MM-DD (default: today).")
    parser.add_argument("--output", type=Path, default=Path("sarb-sla-reasoning.html"))
    parser.add_argument("--output-dir", type=Path, default=output_dir_for("sarb-sla-reasoning"),
                         help="Defaults to the shared SA\\Outputs\\sarb-sla-reasoning folder.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    from sa_paths import latest_input

    if args.input is None:
        found = latest_input("*.csv", "*.xlsx")
        if not found:
            raise SystemExit("No --input given and no .csv/.xlsx file found in the shared SA\\Inputs folder.")
        args.input = found
        print(f"No --input given; using newest file in SA\\Inputs: {args.input}")
    if not args.input.exists():
        raise SystemExit(f"Source file not found: {args.input}")

    as_of = datetime.strptime(args.as_of, "%Y-%m-%d") if args.as_of else datetime.now()
    year = args.year or as_of.year

    frame = load_source(args.input)
    df = build_analysis(frame, year=year, as_of=as_of)
    if df.empty:
        raise SystemExit(f"No in-scope Project/Idea records found for Created year {year}.")

    sa_summary = build_group_summary(df, "Solution Architect")
    portfolio_summary = build_group_summary(df, "Portfolio")
    reason_summary = build_reason_summary(df)
    project_only_df = df[df["Project ID"].str.upper().str.startswith("PRJ")]
    project_summary = build_group_summary(project_only_df, "Solution Architect") if not project_only_df.empty else sa_summary.iloc[0:0]
    drilldown = build_drilldown(df)

    met = int((df["SLA Status"] == "Met").sum())
    not_met = int((df["SLA Status"] == "Not met").sum())
    ctrl_not_met = int(df["Controllable Not Met"].sum())
    undocumented = int(((df["SLA Status"] == "Not met") & (df["Reason Category"] == "undocumented")).sum())
    business_caused = int(((df["SLA Status"] == "Not met") & (df["Reason Category"] == "business_caused")).sum())
    ambiguous = int(((df["SLA Status"] == "Not met") & (df["Reason Category"] == "ambiguous")).sum())
    team_stats = {
        "total": len(df),
        "not_met": not_met,
        "gross_pct": round(100 * met / (met + not_met), 1) if (met + not_met) else 0.0,
        "ctrl_pct": round(100 * met / (met + ctrl_not_met), 1) if (met + ctrl_not_met) else 0.0,
        "undocumented": undocumented,
        "undocumented_pct": round(100 * undocumented / not_met, 1) if not_met else 0.0,
        "business_caused": business_caused,
        "business_caused_pct": round(100 * business_caused / not_met, 1) if not_met else 0.0,
        "ambiguous": ambiguous,
        "ambiguous_pct": round(100 * ambiguous / not_met, 1) if not_met else 0.0,
    }

    args.output_dir.mkdir(parents=True, exist_ok=True)
    sa_summary.to_csv(args.output_dir / "sa-sla-reasoning-summary.csv", index=False)
    portfolio_summary.to_csv(args.output_dir / "portfolio-sla-reasoning-summary.csv", index=False)
    reason_summary.to_csv(args.output_dir / "reason-sla-reasoning-summary.csv", index=False)
    project_summary.to_csv(args.output_dir / "project-only-sla-reasoning-summary.csv", index=False)
    not_met_detail = df[df["SLA Status"] == "Not met"][
        ["Project ID", "Project Name", "Solution Architect", "Portfolio", "Reason", "Reason Category", "Status"]
    ]
    not_met_detail.to_csv(args.output_dir / "not-met-detail.csv", index=False)

    output_path = args.output if args.output.is_absolute() else args.output_dir / args.output.name
    build_html(sa_summary, portfolio_summary, reason_summary, project_summary, drilldown, team_stats, output_path, as_of, year)

    print(f"Items in scope: {len(df)} (Created {year})")
    print(f"Gross SLA %: {team_stats['gross_pct']} | Controllable SLA %: {team_stats['ctrl_pct']}")
    print(f"Undocumented misses: {undocumented} of {not_met} not-met ({team_stats['undocumented_pct']}%)")
    print(f"Created {args.output_dir / 'sa-sla-reasoning-summary.csv'}")
    print(f"Created {args.output_dir / 'portfolio-sla-reasoning-summary.csv'}")
    print(f"Created {args.output_dir / 'reason-sla-reasoning-summary.csv'}")
    print(f"Created {args.output_dir / 'project-only-sla-reasoning-summary.csv'}")
    print(f"Created {args.output_dir / 'not-met-detail.csv'}")
    print(f"Created {output_path}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
