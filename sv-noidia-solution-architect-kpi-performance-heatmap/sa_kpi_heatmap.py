"""Build an interactive KPI heatmap per Solution Architect (SA).

KPIs:
  1. TechPass Created   -> 'TechPass' column is populated
  2. Scenario Published -> 'Publication Status' is populated and not "Not Published"
  3. SARB Approved      -> 'SARB Approval Date' is populated
  4. SAD Approved       -> 'SAD Approved' is populated

Each KPI is expressed as a percentage of eligible work items owned by that
Solution Architect. SAD Approval is applicable only to PRJ work references;
the other KPIs include both PRJ and ID work references. The output HTML heatmap
is color coded:
  >= 70%      dark green
  50% - 70%   light green
  30% - 50%   yellow
  < 30%       light orange

Clicking a heatmap cell shows the list of projects that satisfy (contribute
to) that KPI for that Solution Architect in a drill-down table below the
chart.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go

KPIS = ("TechPass Created", "Scenario Published", "SARB Approved", "SAD Approved")

REQUIRED_COLUMNS = (
    "Work Ref",
    "Work Ref: Name",
    "Work Ref: Solution Architect",
    "Status",
    "TechPass",
    "Publication Status",
    "SARB Approval Date",
    "SAD Approved",
)

COLORSCALE = [
    [0.0, "#f4a261"],
    [0.30, "#f4a261"],
    [0.30, "#fee08b"],
    [0.50, "#fee08b"],
    [0.50, "#a6d96a"],
    [0.70, "#a6d96a"],
    [0.70, "#1a9850"],
    [1.0, "#1a9850"],
]


def parse_args() -> argparse.Namespace:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
    from sa_paths import output_dir_for

    out_dir = output_dir_for("sv-noidia-solution-architect-kpi-performance-heatmap")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        type=Path,
        default=None,
        help="Path to the source CSV (or Excel) export. If omitted, defaults to the newest "
             "matching file in the shared SA\\Inputs folder (falls back to a prompt).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=out_dir / "sv_noidia_sa_kpi_performance_heatmap.html",
        help="Output interactive HTML heatmap path.",
    )
    parser.add_argument(
        "--summary-output",
        type=Path,
        default=out_dir / "sv_noidia_sa_kpi_summary.csv",
        help="Output CSV path for the per-SA KPI percentage summary.",
    )
    parser.add_argument(
        "--exclude-sa",
        action="append",
        default=[],
        help=(
            "Solution Architect name to exclude from the heatmap (case-insensitive). "
            "Can be passed multiple times, e.g. --exclude-sa Unassigned."
        ),
    )
    parser.add_argument(
        "--status",
        action="append",
        default=[],
        help=(
            "Restrict projects to only this Status value (case-insensitive). "
            "Can be passed multiple times to include several statuses, e.g. --status WIP. "
            "If omitted, all non-cancelled projects are included."
        ),
    )
    return parser.parse_args()


def resolve_input_path(args_input: Path | None) -> Path:
    if args_input is not None:
        path = args_input
    else:
        from sa_paths import latest_input
        found = latest_input("*.csv", "*.xlsx", "*.xls")
        if found:
            print(f"No --input given; using newest file in SA\\Inputs: {found}")
            path = found
        else:
            entered = input("Enter the full path to the source file (CSV or Excel): ").strip().strip('"')
            if not entered:
                raise SystemExit("No source file path was provided.")
            path = Path(entered)
    if not path.exists():
        raise SystemExit(f"Source file not found: {path}")
    return path


def load_source(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in (".xlsx", ".xls"):
        frame = pd.read_excel(path, dtype=str)
    else:
        frame = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
    frame.columns = [str(column).strip() for column in frame.columns]
    missing = [column for column in REQUIRED_COLUMNS if column not in frame.columns]
    if missing:
        raise ValueError(f"Missing required column(s): {', '.join(missing)}")
    return frame


def is_blank(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    text = str(value).strip()
    return text == "" or text.lower() in ("nan", "none")


def clean_projects(frame: pd.DataFrame, exclude_sa: list[str] | None = None, status_filter: list[str] | None = None) -> pd.DataFrame:
    frame = frame.copy()
    frame["Status"] = frame["Status"].fillna("").astype(str).str.strip()
    if status_filter:
        allowed_lower = {status.strip().lower() for status in status_filter if status.strip()}
        frame = frame[frame["Status"].str.lower().isin(allowed_lower)].reset_index(drop=True)
    else:
        frame = frame[frame["Status"].str.lower() != "cancelled"].reset_index(drop=True)

    # Non-SAD KPIs include Projects and Ideas; SAD approval applies to Projects only.
    work_ref_upper = frame["Work Ref"].fillna("").astype(str).str.strip().str.upper()
    frame = frame[work_ref_upper.str.startswith("PRJ") | work_ref_upper.str.startswith("ID")].reset_index(drop=True)
    work_ref_upper = frame["Work Ref"].fillna("").astype(str).str.strip().str.upper()
    frame["SAD KPI Applicable"] = work_ref_upper.str.startswith("PRJ")

    frame["Solution Architect"] = frame["Work Ref: Solution Architect"].fillna("Unassigned").astype(str).str.strip()
    frame.loc[frame["Solution Architect"] == "", "Solution Architect"] = "Unassigned"

    if exclude_sa:
        excluded_lower = [name.strip().lower() for name in exclude_sa if name.strip()]
        sa_lower = frame["Solution Architect"].str.lower()
        mask = pd.Series(False, index=frame.index)
        for name in excluded_lower:
            mask |= sa_lower.str.contains(name, regex=False)
        frame = frame[~mask].reset_index(drop=True)

    frame["Project ID"] = frame["Work Ref"].fillna("").astype(str).str.strip()
    frame["Project Name"] = frame["Work Ref: Name"].fillna("").astype(str).str.strip()

    frame["KPI: TechPass Created"] = ~frame["TechPass"].apply(is_blank)

    publication_status = frame["Publication Status"].fillna("").astype(str).str.strip()
    frame["KPI: Scenario Published"] = (
        publication_status.str.lower().ne("")
        & ~publication_status.str.lower().eq("not published")
    )

    frame["KPI: SARB Approved"] = ~frame["SARB Approval Date"].apply(is_blank)
    frame["KPI: SAD Approved"] = ~frame["SAD Approved"].apply(is_blank)
    return frame


def build_summary(frame: pd.DataFrame) -> pd.DataFrame:
    kpi_columns = {
        "TechPass Created": ("KPI: TechPass Created", None),
        "Scenario Published": ("KPI: Scenario Published", None),
        "SARB Approved": ("KPI: SARB Approved", None),
        "SAD Approved": ("KPI: SAD Approved", "SAD KPI Applicable"),
    }
    rows = []
    for architect, group in frame.groupby("Solution Architect"):
        total = len(group)
        sad_eligible = group[group["SAD KPI Applicable"]]
        row = {
            "Solution Architect": architect,
            "Total Work Items": total,
            "SAD Approved Eligible Projects": len(sad_eligible),
        }
        for kpi_label, (column, eligibility_column) in kpi_columns.items():
            eligible_group = group if eligibility_column is None else group[group[eligibility_column]]
            eligible_total = len(eligible_group)
            met = int(eligible_group[column].sum())
            pct = round((met / eligible_total) * 100, 1) if eligible_total else None
            row[f"{kpi_label} %"] = pct
            denominator_label = "eligible" if eligibility_column else "total"
            row[f"{kpi_label} (met/{denominator_label})"] = f"{met}/{eligible_total}"
        rows.append(row)
    summary = pd.DataFrame(rows).sort_values("Solution Architect").reset_index(drop=True)
    return summary


def build_drilldown(frame: pd.DataFrame) -> dict:
    kpi_columns = {
        "TechPass Created": ("KPI: TechPass Created", None),
        "Scenario Published": ("KPI: Scenario Published", None),
        "SARB Approved": ("KPI: SARB Approved", None),
        "SAD Approved": ("KPI: SAD Approved", "SAD KPI Applicable"),
    }

    def project_rows(subset: pd.DataFrame) -> list[dict]:
        return [
            {
                "Project ID": row["Project ID"] or "(blank)",
                "Project Name": row["Project Name"],
                "Status": row["Status"],
            }
            for _, row in subset.iterrows()
        ]

    drilldown: dict[str, dict[str, list[dict]]] = {}
    for architect, group in frame.groupby("Solution Architect"):
        for kpi_label, (column, eligibility_column) in kpi_columns.items():
            key = f"{architect}||{kpi_label}"
            eligible_group = group if eligibility_column is None else group[group[eligibility_column]]
            drilldown[key] = {
                "met": project_rows(eligible_group[eligible_group[column]]),
                "missing": project_rows(eligible_group[~eligible_group[column]]),
            }
    return drilldown


def create_heatmap(summary: pd.DataFrame, drilldown: dict, output_path: Path) -> None:
    architects = summary["Solution Architect"].tolist()
    z = []
    text = []
    for architect in architects:
        values = [summary.loc[summary["Solution Architect"] == architect, f"{kpi} %"].iloc[0] for kpi in KPIS]
        z.append([value if pd.notna(value) else None for value in values])
        text.append([f"{value:.0f}%" if pd.notna(value) else "N/A" for value in values])

    figure = go.Figure(
        data=go.Heatmap(
            z=z,
            x=list(KPIS),
            y=architects,
            zmin=0,
            zmax=100,
            colorscale=COLORSCALE,
            text=text,
            texttemplate="%{text}",
            textfont={"size": 14, "color": "#1c1c1c"},
            hovertemplate="<b>%{y}</b><br>%{x}: %{text}<extra></extra>",
            colorbar={"title": "%", "tickvals": [15, 40, 60, 85], "ticktext": ["<30%", "30-50%", "50-70%", ">=70%"]},
            xgap=3,
            ygap=3,
        )
    )
    figure.update_layout(
        title="SV-Noidia Solution Architect KPI Performance Heatmap<br><sup>SAD Approval is measured only for PRJ projects; other KPIs include PRJ and ID work items.</sup>",
        template="plotly_white",
        height=max(500, 120 + len(architects) * 42),
        font={"family": "Segoe UI, Aptos, sans-serif", "size": 14, "color": "#24323d"},
        margin={"l": 200, "r": 60, "t": 90, "b": 80},
        xaxis={"side": "top", "tickangle": 0},
    )

    fig_html = figure.to_html(full_html=False, include_plotlyjs="cdn", div_id="kpiHeatmap")
    drilldown_json = json.dumps(drilldown)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>Solution Architect KPI Heatmap</title>
<style>
  body {{ font-family: 'Segoe UI', Aptos, sans-serif; margin: 24px; color: #24323d; background: #f6f3ee; }}
  #drilldownPanel {{ margin-top: 24px; }}
  #drilldownTitle {{ font-size: 18px; font-weight: 600; margin-bottom: 12px; }}
  #drilldownColumns {{ display: flex; gap: 24px; flex-wrap: wrap; }}
  .drilldownColumn {{ flex: 1; min-width: 320px; }}
  .drilldownColumn h3 {{ font-size: 15px; margin: 0 0 8px 0; }}
  .drilldownColumn.met h3 {{ color: #1a9850; }}
  .drilldownColumn.missing h3 {{ color: #e08a3e; }}
  table {{ border-collapse: collapse; width: 100%; background: #ffffff; }}
  th, td {{ border: 1px solid #d8dfdc; padding: 8px 12px; text-align: left; font-size: 13px; }}
  th {{ background: #264653; color: #ffffff; }}
  tr:nth-child(even) td {{ background: #f1f5f4; }}
  #emptyHint {{ color: #6b7a75; font-style: italic; }}
</style>
</head>
<body>
{fig_html}
<div id="drilldownPanel">
  <div id="drilldownTitle">Click a heatmap cell to list the eligible work items that met and missed that KPI.</div>
  <div id="drilldownColumns">
    <div class="drilldownColumn met">
      <h3 id="metHeading">Met criteria</h3>
      <div id="metTableContainer"><span class="emptyHint">No cell selected yet.</span></div>
    </div>
    <div class="drilldownColumn missing">
      <h3 id="missingHeading">Missing criteria</h3>
      <div id="missingTableContainer"><span class="emptyHint">No cell selected yet.</span></div>
    </div>
  </div>
</div>
<script>
  const DRILLDOWN_DATA = {drilldown_json};

  function buildTable(rows) {{
    if (rows.length === 0) {{
      return '<span class="emptyHint">None.</span>';
    }}
    let html = '<table><thead><tr><th>Project ID</th><th>Project Name</th><th>Status</th></tr></thead><tbody>';
    rows.forEach(function (row) {{
      html += '<tr><td>' + row['Project ID'] + '</td><td>' + row['Project Name'] + '</td><td>' + row['Status'] + '</td></tr>';
    }});
    html += '</tbody></table>';
    return html;
  }}

  function renderTable(architect, kpi) {{
    const key = architect + "||" + kpi;
    const data = DRILLDOWN_DATA[key] || {{ met: [], missing: [] }};
    const titleEl = document.getElementById('drilldownTitle');
    const metHeading = document.getElementById('metHeading');
    const missingHeading = document.getElementById('missingHeading');
    const metContainer = document.getElementById('metTableContainer');
    const missingContainer = document.getElementById('missingTableContainer');

    const total = data.met.length + data.missing.length;
    titleEl.textContent = architect + " — " + kpi + " (" + data.met.length + " of " + total + " eligible work items met)";
    metHeading.textContent = "Met criteria (" + data.met.length + ")";
    missingHeading.textContent = "Missing criteria (" + data.missing.length + ")";
    metContainer.innerHTML = buildTable(data.met);
    missingContainer.innerHTML = buildTable(data.missing);
  }}

  document.addEventListener('DOMContentLoaded', function () {{
    const plotDiv = document.getElementById('kpiHeatmap');
    plotDiv.on('plotly_click', function (data) {{
      const point = data.points[0];
      renderTable(point.y, point.x);
    }});
  }});
</script>
</body>
</html>
"""
    output_path.write_text(html, encoding="utf-8")


def main() -> None:
    args = parse_args()
    input_path = resolve_input_path(args.input)
    frame = load_source(input_path)
    frame = clean_projects(
        frame,
        exclude_sa=args.exclude_sa or ["Unassigned"],
        status_filter=args.status or ["WIP"],
    )
    if frame.empty:
        raise SystemExit("No WIP projects or ideas were found in the source file.")

    summary = build_summary(frame)
    drilldown = build_drilldown(frame)

    args.summary_output.parent.mkdir(parents=True, exist_ok=True)
    summary.to_csv(args.summary_output, index=False)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    create_heatmap(summary, drilldown, args.output)

    print(f"WIP work items included: {len(frame)}")
    print(f"Solution Architects: {frame['Solution Architect'].nunique()}")
    print(f"Created {args.summary_output}")
    print(f"Created {args.output}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
