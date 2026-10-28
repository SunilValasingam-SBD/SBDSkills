"""Build a SARB Due Date heatmap: Solution Architect x Month, with margin totals."""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go

DATE_FORMATS = ("%m/%d/%Y %H:%M", "%m/%d/%Y", "%d-%b-%y", "%d-%b-%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S")

REQUIRED_COLUMNS = (
    "Work Ref",
    "Work Ref: Name",
    "Work Ref: Solution Architect",
    "SARB Due Date",
)

# Default visible window: August 2026 through the end of 2027, nothing else.
RANGE_START = date(2026, 8, 1)
RANGE_END = date(2027, 12, 31)


def parse_date(value) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    text = str(value).strip()
    if not text:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def load_rows(path: Path) -> list[dict]:
    if path.suffix.lower() in (".xlsx", ".xls"):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.worksheets[0]
        headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
        rows = []
        for raw in ws.iter_rows(min_row=2, values_only=True):
            rows.append(dict(zip(headers, raw)))
        return rows
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


# Work List exports used across skills mix several classifications in one file
# (Projects, Ideas, TechPass requests, SA NPD, NPD) distinguished only by the
# "Work Ref" prefix. This skill only cares about Projects and Ideas.
WORK_REF_PREFIX_RE = re.compile(r"^(PRJ|ID)\d", re.IGNORECASE)


def is_project_or_idea(work_ref: str) -> bool:
    return bool(WORK_REF_PREFIX_RE.match(work_ref.strip()))


def build_dataframe(rows: list[dict], include_approved: bool, as_of: date,
                     range_start: date = RANGE_START, range_end: date = RANGE_END) -> pd.DataFrame:
    has_rating_column = any("SARB Rating" in r for r in rows)
    records = []
    for r in rows:
        work_ref = str(r.get("Work Ref") or "").strip()
        if not is_project_or_idea(work_ref):
            continue  # skip TechPass / SA NPD / NPD and anything else that isn't a Project or Idea

        if not include_approved and has_rating_column:
            rating = str(r.get("SARB Rating") or "").strip()
            if rating:
                continue  # already approved/rated -> no longer a pending due date

        due = parse_date(r.get("SARB Due Date"))
        if due is None:
            continue
        if due < range_start or due > range_end:
            continue  # out of scope: only Aug 2026 through end of 2027
        start = parse_date(r.get("Work Ref: Start Date")) or parse_date(r.get("Created"))

        name = str(r.get("Work Ref: Name") or "").strip()
        sa = str(r.get("Work Ref: Solution Architect") or "").strip() or "(Unassigned)"
        reason = str(r.get("SARB Reason") or "").strip()

        records.append({
            "Work Ref": work_ref,
            "Name": name,
            "SA": sa,
            "Start": start,
            "Due": due,
            "Reason": reason,
            "Status": "Overdue" if due < as_of else "Upcoming",
            "Label": work_ref,
        })

    df = pd.DataFrame.from_records(records)
    if df.empty:
        return df

    df["Due"] = pd.to_datetime(df["Due"])
    df["Month"] = df["Due"].dt.to_period("M")
    return df


def month_range(range_start: date, range_end: date) -> list:
    return list(pd.period_range(start=range_start, end=range_end, freq="M"))


PROFESSIONAL_FONT = "'Segoe UI Rounded', 'Varela Round', 'Century Gothic', 'Calibri', 'Trebuchet MS', sans-serif"


def build_heatmap_chart(df: pd.DataFrame, as_of: date,
                         range_start: date = RANGE_START, range_end: date = RANGE_END):
    months = month_range(range_start, range_end)
    month_labels = [m.strftime("%b %Y") for m in months]
    as_of_period = pd.Period(as_of, freq="M")

    # Order SAs by total workload (busiest at top).
    sa_order = df.groupby("SA").size().sort_values(ascending=False).index.tolist()

    total = df.groupby(["SA", "Month"]).size().unstack(fill_value=0).reindex(index=sa_order, columns=months, fill_value=0)
    overdue = (
        df[df["Status"] == "Overdue"].groupby(["SA", "Month"]).size()
        .unstack(fill_value=0).reindex(index=sa_order, columns=months, fill_value=0)
    )

    # Numeric axis positions so overdue cells can be outlined precisely with shapes.
    x_idx = list(range(len(months)))
    y_idx = list(range(len(sa_order)))

    # Build per-cell hover text (compact) and a richer click-detail payload (customdata)
    # listing every individual item, for the interactive click-to-expand panel.
    hover = [[""] * len(months) for _ in sa_order]
    cell_text = [[""] * len(months) for _ in sa_order]
    click_detail = [[""] * len(months) for _ in sa_order]
    overdue_cells = []
    zmax = max(1, int(df.groupby(["SA", "Month"]).size().max()))
    for i, sa in enumerate(sa_order):
        for j, m in enumerate(months):
            cell_items = df[(df["SA"] == sa) & (df["Month"] == m)]
            if cell_items.empty:
                continue
            count = len(cell_items)
            od = int((cell_items["Status"] == "Overdue").sum())
            items_sorted = cell_items.sort_values("Due")
            lines = [f"{r['Work Ref']} - {r['Name'][:45]} ({r['Status']}, due {r['Due']:%d-%b})"
                      for _, r in items_sorted.iterrows()]
            hover[i][j] = f"<b>{sa} - {month_labels[j]}</b><br>" + "<br>".join(lines)
            list_items = "".join(
                f'<li><span class="ref">{r["Work Ref"]}</span> - {r["Name"]}'
                f'<br><span class="meta">Status: {r["Status"]} &#8226; Due: {r["Due"]:%d-%b-%Y}</span></li>'
                for _, r in items_sorted.iterrows()
            )
            click_detail[i][j] = (
                f'<div class="detail-heading">{sa} &#8226; {month_labels[j]} '
                f'({count} item{"s" if count != 1 else ""}, {od} overdue)</div>'
                f'<ul class="detail-list">{list_items}</ul>'
            )
            text_color = "#ffffff" if count / zmax >= 0.6 else "#1a1a1a"
            cell_text[i][j] = f'<b style="color:{text_color}">{count}</b>'
            if od:
                overdue_cells.append((i, j))

    grand_total = int(total.values.sum())
    grand_overdue = int(overdue.values.sum())

    fig = go.Figure(
        data=[
            go.Heatmap(
                z=total.values,
                x=x_idx,
                y=y_idx,
                text=cell_text,
                texttemplate="%{text}",
                textfont={"size": 13, "family": PROFESSIONAL_FONT, "color": "#1a1a1a"},
                hoverinfo="text",
                hovertext=hover,
                customdata=click_detail,
                colorscale="Blues",
                showscale=False,
                zmin=0,
                xgap=2, ygap=2,
            )
        ]
    )

    fig.update_xaxes(tickvals=x_idx, ticktext=[f"<b>{lbl}</b>" for lbl in month_labels], tickangle=-45,
                      tickfont=dict(size=13, family=PROFESSIONAL_FONT), title="")
    fig.update_yaxes(tickvals=y_idx, ticktext=[f"<b>{lbl}</b>" for lbl in sa_order], autorange="reversed",
                      tickfont=dict(size=13, family=PROFESSIONAL_FONT), title="")

    # Outline every cell that contains at least one overdue item in orange.
    for i, j in overdue_cells:
        fig.add_shape(
            type="rect", xref="x", yref="y",
            x0=j - 0.5, x1=j + 0.5, y0=i - 0.5, y1=i + 0.5,
            line=dict(color="#e67e22", width=3), fillcolor="rgba(0,0,0,0)",
            layer="above",
        )

    # Mark today's exact position within its month (not just the month boundary),
    # so "today" visually lands where it actually falls (e.g. late Oct, not early Oct).
    if as_of_period in months:
        idx = months.index(as_of_period)
        days_in_month = as_of_period.days_in_month
        day_fraction = (as_of.day - 1) / days_in_month
        today_x = idx - 0.5 + day_fraction
        fig.add_vline(x=today_x, line_width=2, line_dash="dash", line_color="#444444")

    # Single "Total" label in the top-left corner, above the Solution Architect column.
    fig.add_annotation(
        text=f"<b>Total: {grand_total}</b>",
        xref="paper", yref="paper", x=0, y=1.22, xanchor="left", yanchor="bottom",
        showarrow=False, font=dict(size=15, family=PROFESSIONAL_FONT, color="#1a2733"),
    )

    fig.update_layout(
        font=dict(family=PROFESSIONAL_FONT, size=13, color="#2c3e50"),
        title=dict(
            text=(f"<b>SARB Due Dates by Solution Architect &amp; Month - Aug 2026 through 2027</b><br>"
                  f"<sup>As of {as_of:%d-%b-%Y} &#8226; {grand_overdue} overdue (orange outline) "
                  f"&#8226; dashed line = today</sup>"),
            font=dict(family=PROFESSIONAL_FONT, size=20, color="#1a2733"),
            x=0.5, xanchor="center", y=0.95, yanchor="top",
        ),
        height=max(550, 42 * len(sa_order) + 190),
        margin=dict(l=10, r=10, t=140, b=80),
    )
    return fig


def render_interactive_html(fig: go.Figure) -> str:
    """Render the figure plus a click-to-expand details panel below it.

    Clicking any heatmap cell shows the full list of items in that
    Solution Architect / Month cell (reference, name, status, due date)
    in the panel, using each cell's `customdata` payload.
    """
    plot_html = fig.to_html(
        full_html=False,
        include_plotlyjs="cdn",
        div_id="sarb-heatmap",
        config={"responsive": True},
        post_script=(
            "var gd = document.getElementById('sarb-heatmap');"
            "gd.on('plotly_click', function(evt) {"
            "  var pt = evt.points && evt.points[0];"
            "  var panel = document.getElementById('sarb-detail-panel');"
            "  if (!pt || pt.customdata === undefined || pt.customdata === '') {"
            "    panel.innerHTML = '<div class=\"detail-placeholder\">No items in this cell.</div>';"
            "    return;"
            "  }"
            "  panel.innerHTML = pt.customdata;"
            "  panel.scrollIntoView({behavior: 'smooth', block: 'nearest'});"
            "});"
        ),
    )
    font = PROFESSIONAL_FONT
    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>SARB Due Dates</title>
<style>
  body {{ font-family: {font}; margin: 0; padding: 20px; background: #fafbfc; color: #2c3e50; }}
  #sarb-detail-panel {{
    margin-top: 18px; padding: 16px 20px; border: 1px solid #d9e2ec; border-radius: 8px;
    background: #ffffff; box-shadow: 0 1px 3px rgba(0,0,0,0.06); min-height: 56px;
  }}
  .detail-placeholder {{ color: #8492a6; font-style: italic; }}
  .detail-heading {{ font-weight: bold; font-size: 15px; color: #1a2733; margin-bottom: 10px; }}
  .detail-list {{ list-style: none; margin: 0; padding: 0; }}
  .detail-list li {{
    padding: 8px 0; border-bottom: 1px solid #eef1f4; font-size: 13px; line-height: 1.5;
  }}
  .detail-list li:last-child {{ border-bottom: none; }}
  .detail-list .ref {{ font-weight: bold; color: #2471a3; }}
  .detail-list .meta {{ color: #6b7a8f; font-size: 12px; }}
  .hint {{ color: #8492a6; font-size: 13px; margin: 8px 0 0 4px; }}
</style>
</head>
<body>
{plot_html}
<div class="hint">Click any cell in the heatmap to see its item details below.</div>
<div id="sarb-detail-panel"><div class="detail-placeholder">Click a cell above to see its item details here.</div></div>
</body>
</html>"""


def parse_args() -> argparse.Namespace:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
    from sa_paths import output_dir_for

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=None,
                         help="Path to the source CSV/Excel export with SARB Due Date, Work Ref, "
                              "Work Ref: Solution Architect, etc. Defaults to the newest matching "
                              "file in the shared SA\\Inputs folder.")
    parser.add_argument("--as-of", type=str, default=None, help="As-of date YYYY-MM-DD (default: today).")
    parser.add_argument("--include-approved", action="store_true",
                         help="Also include rows that already have a SARB Rating set (default: excluded, "
                              "since an approved item no longer has a pending due date).")
    parser.add_argument("--output-dir", type=Path, default=output_dir_for("sarb-due-date-gantt"),
                         help="Defaults to the shared SA\\Outputs\\sarb-due-date-gantt folder.")
    args = parser.parse_args()

    if args.input is None:
        from sa_paths import latest_input
        found = latest_input("*.csv", "*.xlsx", "*.xls")
        if not found:
            raise SystemExit("No --input given and no .csv/.xlsx file found in the shared SA\\Inputs folder.")
        print(f"No --input given; using newest file in SA\\Inputs: {found}")
        args.input = found
    return args


def main() -> None:
    args = parse_args()
    if not args.input.exists():
        raise SystemExit(f"Source file not found: {args.input}")

    as_of = datetime.strptime(args.as_of, "%Y-%m-%d").date() if args.as_of else date.today()

    rows = load_rows(args.input)
    if not rows:
        raise SystemExit("No rows found in the source file.")

    missing = [c for c in REQUIRED_COLUMNS if c not in rows[0]]
    if missing:
        raise SystemExit(f"Missing required column(s): {', '.join(missing)}")

    df = build_dataframe(rows, include_approved=args.include_approved, as_of=as_of)
    if df.empty:
        raise SystemExit("No rows with a SARB Due Date (and no SARB Rating, unless --include-approved) were found.")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    heatmap_path = args.output_dir / "SARB_Due_Dates_Heatmap.html"

    fig = build_heatmap_chart(df, as_of)
    heatmap_path.write_text(render_interactive_html(fig), encoding="utf-8")

    overdue = int((df["Status"] == "Overdue").sum())
    upcoming = int((df["Status"] == "Upcoming").sum())
    print(f"Included {len(df)} items ({overdue} overdue, {upcoming} upcoming) as of {as_of:%Y-%m-%d}.")
    print(f"Heatmap: {heatmap_path}")


if __name__ == "__main__":
    main()
