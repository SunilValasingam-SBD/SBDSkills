"""Create an interactive Solution Architect idea/project engagement dashboard."""

from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go

REQUIRED_COLUMNS = ("Work Ref", "Idea ID", "Program Owning Portfolio")
DATE_CANDIDATES = (
    "Created",
    "Created Date",
    "Initiative Created",
    "Created On",
    "Date Created",
)
SA_CANDIDATES = (
    "Work Ref: Solution Architect",
    "Solution Architect",
    "SA",
    "Assigned to",
)
# Consolidate related Portfolio names into a single reporting group for the Portfolio-based chart.
PORTFOLIO_GROUP_MAP = {
    "CIO Office": "HR & Legal Compliance",
    "HR, Legal, & Compliance": "HR & Legal Compliance",
    "Mergers, Acquisitions & Divestitures": "HR & Legal Compliance",
    "Commercial & Regions": "Customer Engagement",
    "Customer Engagement": "Customer Engagement",
}


def parse_args() -> argparse.Namespace:
    import sys
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
    from sa_paths import output_dir_for

    out_dir = output_dir_for("sa-engagement-analysis")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, help="Path to the source CSV. If omitted, defaults to the "
                         "newest .csv file in the shared SA\\Inputs folder (falls back to a prompt).")
    parser.add_argument("--year", type=int, default=date.today().year, help="Initiative creation year.")
    parser.add_argument("--created-column", help="Explicit creation-date column name.")
    parser.add_argument("--sa-column", help="Explicit Solution Architect column name.")
    parser.add_argument(
        "--output",
        type=Path,
        default=out_dir / "sa_engagement_dashboard.html",
        help="Interactive HTML output path.",
    )
    parser.add_argument(
        "--summary-output",
        type=Path,
        default=out_dir / "sa_engagement_summary.csv",
        help="Grouped summary CSV output path.",
    )
    parser.add_argument(
        "--detail-output",
        type=Path,
        default=out_dir / "sa_engagement_detail.csv",
        help="Normalized detail CSV output path.",
    )
    return parser.parse_args()


def resolve_input_path(value: Path | None) -> Path:
    entered = value
    if entered is None:
        import sys
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
        from sa_paths import latest_input
        found = latest_input("*.csv")
        if found:
            print(f"No --input given; using newest file in SA\\Inputs: {found}")
            entered = found
        else:
            text = input("Enter the full path to the source CSV: ").strip().strip('"')
            if not text:
                raise SystemExit("No source file path was provided.")
            entered = Path(text)
    if not entered.exists():
        raise SystemExit(f"Source file not found: {entered}")
    if entered.suffix.lower() != ".csv":
        raise SystemExit("The input must be a CSV file.")
    return entered


def clean_text(series: pd.Series, fallback: str = "Unassigned") -> pd.Series:
    result = series.fillna("").astype(str).str.strip()
    return result.mask(result.str.lower().isin({"", "nan", "none", "null"}), fallback)


def is_blank(value: object) -> bool:
    return str(value).strip().lower() in {"", "nan", "none", "null"}


def select_column(columns: list[str], explicit: str | None, candidates: tuple[str, ...], label: str) -> str:
    if explicit:
        if explicit not in columns:
            raise SystemExit(f"{label} column '{explicit}' was not found. Available columns: {', '.join(columns)}")
        return explicit
    exact = {column.casefold(): column for column in columns}
    matches = [exact[name.casefold()] for name in candidates if name.casefold() in exact]
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise SystemExit(
            f"Multiple {label} columns were found ({', '.join(matches)}). "
            f"Pass --{label.replace(' ', '-')}-column to choose one."
        )
    raise SystemExit(
        f"Could not detect a {label} column. Expected one of: {', '.join(candidates)}. "
        f"Available columns: {', '.join(columns)}"
    )


def load_and_classify(path: Path, year: int, created_column: str | None, sa_column: str | None) -> tuple[pd.DataFrame, str, str]:
    frame = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
    frame.columns = [str(column).strip() for column in frame.columns]
    missing = [column for column in REQUIRED_COLUMNS if column not in frame.columns]
    if missing:
        raise SystemExit(f"Missing required column(s): {', '.join(missing)}")

    date_column = select_column(list(frame.columns), created_column, DATE_CANDIDATES, "creation date")
    architect_column = select_column(list(frame.columns), sa_column, SA_CANDIDATES, "SA")
    created = pd.to_datetime(frame[date_column], errors="coerce")
    if created.notna().sum() == 0:
        raise SystemExit(f"Creation date column '{date_column}' contains no parseable dates.")

    work_ref = frame["Work Ref"].fillna("").astype(str).str.strip()
    upper_ref = work_ref.str.upper()
    frame = frame.loc[created.dt.year.eq(year) & (upper_ref.str.startswith("ID") | upper_ref.str.startswith("PRJ"))].copy()
    frame["Work Ref"] = work_ref.loc[frame.index]
    frame["Created Date"] = created.loc[frame.index].dt.strftime("%Y-%m-%d")
    frame["Engagement Type"] = frame["Work Ref"].str.upper().str.startswith("PRJ").map({True: "Project", False: "Idea"})
    frame["Idea ID"] = frame["Idea ID"].fillna("").astype(str).str.strip()
    frame["E2E Engagement"] = (
        frame["Engagement Type"].eq("Project")
        & ~frame["Idea ID"].map(is_blank)
    )
    frame["Project-only Engagement"] = frame["Engagement Type"].eq("Project") & ~frame["E2E Engagement"]
    frame["Portfolio"] = clean_text(frame["Program Owning Portfolio"]).str.rstrip("/").str.split("/").str[-1].str.strip()
    frame["Portfolio"] = frame["Portfolio"].mask(frame["Portfolio"].eq(""), "Unassigned")
    frame["Portfolio"] = frame["Portfolio"].replace(PORTFOLIO_GROUP_MAP)
    frame["Solution Architect"] = clean_text(frame[architect_column])
    return frame, date_column, architect_column


def summarize(frame: pd.DataFrame, dimension: str) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for label, group in frame.groupby(dimension, dropna=False):
        total = group["Work Ref"].nunique()
        projects = group.loc[group["Engagement Type"].eq("Project"), "Work Ref"].nunique()
        ideas = group.loc[group["Engagement Type"].eq("Idea"), "Work Ref"].nunique()
        e2e = group.loc[group["E2E Engagement"], "Work Ref"].nunique()
        rows.append(
            {
                "View": "Portfolio" if dimension == "Portfolio" else "Solution Architect",
                "Name": label,
                "Total Engagement": total,
                "Ideas": ideas,
                "Projects": projects,
                "E2E Projects": e2e,
                "Project-only Projects": max(projects - e2e, 0),
                "E2E % of Total": round(e2e / total * 100, 1) if total else 0.0,
                "Engaged % of Total": round((ideas + e2e) / total * 100, 1) if total else 0.0,
            }
        )
    return pd.DataFrame(rows).sort_values(["View", "Total Engagement", "Name"], ascending=[True, False, True])


FONT_FAMILY = "'Segoe UI', 'Helvetica Neue', Arial, sans-serif"
SOFT_PALETTE = ["#5B8FB9", "#4CAF9E", "#F2A65A", "#9575CD"]
SOFT_E2E_COLOR = "#E4572E"
SOFT_AVERAGE_COLOR = "#2E8B7A"
CATEGORY_PALETTE = [
    "#5B8FB9", "#4CAF9E", "#F2A65A", "#9575CD", "#E4572E", "#2E8B7A",
    "#F4B942", "#7D8CC4", "#C86B85", "#6FB07F", "#D98E04", "#8E7CC3",
]


def build_engagement_figure(
    combined: pd.DataFrame,
    views: list[str],
    title: str,
    count_metrics: tuple[str, ...],
    numerator_columns: tuple[str, ...],
    pct_column: str,
    pct_axis_title: str,
    x_axis_title: str,
    annotation_text: str,
) -> go.Figure:
    figure = go.Figure()
    buttons = []
    traces_per_view = len(count_metrics) + 2
    max_rows = 0
    for view in views:
        subset = combined[combined["View"].eq(view)].sort_values("Total Engagement", ascending=False)
        max_rows = max(max_rows, len(subset))
        visible = [False] * (len(views) * traces_per_view)
        offset = views.index(view) * traces_per_view
        for index, metric in enumerate(count_metrics):
            figure.add_trace(
                go.Bar(
                    x=subset["Name"],
                    y=subset[metric],
                    name=metric,
                    orientation="v",
                    visible=view == views[0],
                    marker={"color": SOFT_PALETTE[index % len(SOFT_PALETTE)], "line": {"width": 0}},
                    text=subset[metric],
                    texttemplate="%{text}",
                    textposition="inside",
                    insidetextanchor="start",
                    textfont={"color": "#FFFFFF", "size": 11, "family": FONT_FAMILY},
                    customdata=subset[[pct_column, "Project-only Projects"]].to_numpy(),
                    hovertemplate=(
                        "<b>%{x}</b><br>" + metric + ": %{y}<br>"
                        f"{pct_axis_title}: " + "%{customdata[0]:.1f}%<br>"
                        "Project-only projects: %{customdata[1]}<extra></extra>"
                    ),
                )
            )
            visible[offset + index] = True
        figure.add_trace(
            go.Scatter(
                x=subset["Name"],
                y=subset[pct_column],
                name=pct_axis_title,
                mode="lines+markers+text",
                text=[f"{value:.1f}%" for value in subset[pct_column]],
                textposition="top center",
                yaxis="y2",
                visible=view == views[0],
                line={"color": SOFT_E2E_COLOR, "width": 2, "shape": "spline"},
                marker={"color": SOFT_E2E_COLOR, "size": 9},
                hovertemplate="<b>%{x}</b><br>" + pct_axis_title + ": %{y:.1f}%<extra></extra>",
            )
        )
        visible[offset + len(count_metrics)] = True
        team_numerator = sum(subset[column].sum() for column in numerator_columns)
        team_total = subset["Total Engagement"].sum()
        team_pct = round(team_numerator / team_total * 100, 1) if team_total else 0.0
        figure.add_trace(
            go.Scatter(
                x=subset["Name"],
                y=[team_pct] * len(subset),
                name=f"Team {pct_axis_title} (Average)",
                mode="lines",
                line={"color": SOFT_AVERAGE_COLOR, "dash": "dash", "width": 2},
                yaxis="y2",
                visible=view == views[0],
                hovertemplate=f"Team average {pct_axis_title.lower()}: {team_pct:.1f}%<extra></extra>",
            )
        )
        visible[offset + len(count_metrics) + 1] = True
        buttons.append(
            {
                "label": view,
                "method": "update",
                "args": [
                    {"visible": visible},
                    {"title": f"{title} by {view}", "xaxis": {"title": view}},
                ],
            }
        )
    initial_visible = [True] * traces_per_view + [False] * traces_per_view
    for trace, is_visible in zip(figure.data, initial_visible):
        trace.visible = is_visible
    figure.update_layout(
        title={"text": f"{title} ({views[0]})", "font": {"family": FONT_FAMILY, "size": 22, "color": "#3B3B3B"}},
        font={"family": FONT_FAMILY, "size": 13, "color": "#4A4A4A"},
        barmode="group",
        height=max(600, 30 * max_rows + 220),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FAFAF8",
        yaxis={"title": "Count", "domain": [0, 1], "gridcolor": "#E8E6E1", "zerolinecolor": "#E8E6E1"},
        yaxis2={
            "title": pct_axis_title,
            "overlaying": "y",
            "side": "right",
            "range": [0, 100],
            "showgrid": False,
        },
        xaxis={"title": x_axis_title, "tickangle": -45, "gridcolor": "#E8E6E1"},
        legend={"orientation": "h", "y": -0.35, "font": {"size": 12, "color": "#5A5A5A"}},
        updatemenus=[
            {
                "buttons": buttons,
                "direction": "down",
                "x": 0.8,
                "y": 1.15,
                "xanchor": "left",
                "font": {"family": FONT_FAMILY, "size": 12, "color": "#4A4A4A"},
                "bgcolor": "#F5F4F1",
                "bordercolor": "#D8D6D0",
            }
        ],
        margin={"l": 80, "r": 80, "t": 150, "b": 180},
        annotations=[
            {
                "text": annotation_text,
                "showarrow": False,
                "x": 0,
                "y": 1.08,
                "xref": "paper",
                "yref": "paper",
                "xanchor": "left",
                "font": {"family": FONT_FAMILY, "size": 12, "color": "#8A8A8A"},
            }
        ],
    )
    return figure


def build_monthly_figure(frame: pd.DataFrame, year: int) -> go.Figure:
    engaged = frame.loc[~frame["Project-only Engagement"]].copy()
    engaged["Month"] = pd.to_datetime(engaged["Created Date"]).dt.strftime("%b")
    month_order = list(pd.date_range(f"{year}-01-01", periods=12, freq="MS").strftime("%b"))
    dims = ["Portfolio", "Solution Architect"]
    dim_categories = {dim: sorted(frame[dim].unique()) for dim in dims}
    trace_counts = [len(dim_categories[dim]) + 1 for dim in dims]
    total_traces = sum(trace_counts)
    offsets = [sum(trace_counts[:index]) for index in range(len(dims))]

    figure = go.Figure()
    buttons = []
    for dim_index, dim in enumerate(dims):
        categories = dim_categories[dim]
        pivot = (
            engaged.groupby([dim, "Month"])["Work Ref"]
            .nunique()
            .unstack(fill_value=0)
            .reindex(index=categories, fill_value=0)
            .reindex(columns=month_order, fill_value=0)
        )
        visible = [False] * total_traces
        offset = offsets[dim_index]
        for cat_index, category in enumerate(categories):
            values = pivot.loc[category]
            labels = [str(v) if v else "" for v in values]
            figure.add_trace(
                go.Bar(
                    x=month_order,
                    y=values,
                    name=str(category),
                    marker={"color": CATEGORY_PALETTE[cat_index % len(CATEGORY_PALETTE)], "line": {"width": 0}},
                    text=labels,
                    textposition="inside",
                    insidetextanchor="middle",
                    textfont={"color": "#FFFFFF", "size": 10, "family": FONT_FAMILY},
                    visible=dim == dims[0],
                    hovertemplate=f"<b>{category}</b><br>" + "%{x}: %{y} engaged<extra></extra>",
                )
            )
            visible[offset + cat_index] = True
        total_series = pivot.sum(axis=0)
        figure.add_trace(
            go.Scatter(
                x=month_order,
                y=total_series,
                name="Total Engaged (Team)",
                mode="lines+markers+text",
                text=[str(v) for v in total_series],
                textposition="top center",
                textfont={"color": SOFT_AVERAGE_COLOR, "size": 12, "family": FONT_FAMILY},
                line={"color": SOFT_AVERAGE_COLOR, "width": 3, "dash": "dot"},
                marker={"color": SOFT_AVERAGE_COLOR, "size": 8},
                visible=dim == dims[0],
                hovertemplate="<b>%{x}</b><br>Total team engaged: %{y}<extra></extra>",
            )
        )
        visible[offset + len(categories)] = True
        buttons.append(
            {
                "label": dim,
                "method": "update",
                "args": [{"visible": visible}, {"title": f"Engaged Initiatives by Month ({dim} view)"}],
            }
        )
    initial_visible = [True] * trace_counts[0] + [False] * (total_traces - trace_counts[0])
    for trace, is_visible in zip(figure.data, initial_visible):
        trace.visible = is_visible
    figure.update_layout(
        title={
            "text": f"Engaged Initiatives by Month ({dims[0]} view)",
            "font": {"family": FONT_FAMILY, "size": 22, "color": "#3B3B3B"},
        },
        font={"family": FONT_FAMILY, "size": 13, "color": "#4A4A4A"},
        barmode="stack",
        height=650,
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FAFAF8",
        xaxis={"title": "Month", "gridcolor": "#E8E6E1"},
        yaxis={"title": "Engaged Initiatives", "gridcolor": "#E8E6E1", "zerolinecolor": "#E8E6E1"},
        legend={"orientation": "h", "y": -0.3, "font": {"size": 11, "color": "#5A5A5A"}},
        updatemenus=[
            {
                "buttons": buttons,
                "direction": "down",
                "x": 0.8,
                "y": 1.15,
                "xanchor": "left",
                "font": {"family": FONT_FAMILY, "size": 12, "color": "#4A4A4A"},
                "bgcolor": "#F5F4F1",
                "bordercolor": "#D8D6D0",
            }
        ],
        margin={"l": 80, "r": 40, "t": 130, "b": 140},
        annotations=[
            {
                "text": "Engaged = Ideas + E2E Projects (all rows except Project-only); idea-stage engagement counts immediately",
                "showarrow": False,
                "x": 0,
                "y": 1.08,
                "xref": "paper",
                "yref": "paper",
                "xanchor": "left",
                "font": {"family": FONT_FAMILY, "size": 12, "color": "#8A8A8A"},
            }
        ],
    )
    return figure


def build_dashboard(frame: pd.DataFrame, output: Path, year: int, date_column: str, sa_column: str) -> None:
    views = ["Portfolio", "Solution Architect"]
    portfolio = summarize(frame, "Portfolio")
    architect = summarize(frame, "Solution Architect")
    combined = pd.concat([portfolio, architect], ignore_index=True)
    overview_annotation = (
        f"Source creation column: {date_column} | SA column: {sa_column} | "
        "Engaged % = (Ideas + E2E Projects) / Total Engagement"
    )
    figure_overview = build_engagement_figure(
        combined,
        views,
        title=f"SA Engagement Overview ({year})",
        count_metrics=("Total Engagement", "Ideas", "Projects", "E2E Projects"),
        numerator_columns=("Ideas", "E2E Projects"),
        pct_column="Engaged % of Total",
        pct_axis_title="Engaged % of Total",
        x_axis_title="Portfolio",
        annotation_text=overview_annotation,
    )

    figure_monthly = build_monthly_figure(frame, year)

    projects_only_frame = frame.loc[frame["Engagement Type"].eq("Project")]
    portfolio_po = summarize(projects_only_frame, "Portfolio")
    architect_po = summarize(projects_only_frame, "Solution Architect")
    combined_po = pd.concat([portfolio_po, architect_po], ignore_index=True)
    projects_only_annotation = (
        "Idea-only rows excluded; E2E % of Total = E2E Projects / Projects "
        f"(source creation column: {date_column} | SA column: {sa_column})"
    )
    figure_projects_only = build_engagement_figure(
        combined_po,
        views,
        title=f"E2E Engagement Excluding Ideas ({year})",
        count_metrics=("Projects", "E2E Projects", "Project-only Projects"),
        numerator_columns=("E2E Projects",),
        pct_column="E2E % of Total",
        pct_axis_title="E2E % of Total",
        x_axis_title="Portfolio",
        annotation_text=projects_only_annotation,
    )

    section_style = "font-family:'Segoe UI','Helvetica Neue',Arial,sans-serif;color:#3B3B3B;margin:24px 0 4px 12px;"
    html = "\n".join(
        [
            "<html><head><meta charset='utf-8'><title>SA Engagement Dashboard</title></head><body>",
            f"<h2 style=\"{section_style}\">1. Engagement Overview (Portfolio / SA)</h2>",
            figure_overview.to_html(full_html=False, include_plotlyjs=True),
            f"<h2 style=\"{section_style}\">2. Engaged Initiatives by Month (Portfolio / SA)</h2>",
            figure_monthly.to_html(full_html=False, include_plotlyjs=False),
            f"<h2 style=\"{section_style}\">3. E2E Engagement Excluding Idea-only Rows (Portfolio / SA)</h2>",
            figure_projects_only.to_html(full_html=False, include_plotlyjs=False),
            "</body></html>",
        ]
    )
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(html, encoding="utf-8")


def main() -> None:
    args = parse_args()
    source = resolve_input_path(args.input)
    frame, date_column, sa_column = load_and_classify(source, args.year, args.created_column, args.sa_column)
    summary = pd.concat([summarize(frame, "Portfolio"), summarize(frame, "Solution Architect")], ignore_index=True)
    args.summary_output.parent.mkdir(parents=True, exist_ok=True)
    args.detail_output.parent.mkdir(parents=True, exist_ok=True)
    summary.to_csv(args.summary_output, index=False, encoding="utf-8-sig")
    detail_columns = [
        "Work Ref", "Idea ID", "Created Date", "Portfolio", "Solution Architect",
        "Engagement Type", "E2E Engagement", "Project-only Engagement",
    ]
    frame[detail_columns].to_csv(args.detail_output, index=False, encoding="utf-8-sig")
    build_dashboard(frame, args.output, args.year, date_column, sa_column)
    print(f"Included {frame['Work Ref'].nunique()} distinct initiatives ({len(frame)} rows) for {args.year}.")
    print(f"Detected creation date: {date_column}; SA: {sa_column}")
    print(f"Dashboard: {args.output}")
    print(f"Summary: {args.summary_output}")
    print(f"Detail: {args.detail_output}")


if __name__ == "__main__":
    main()
