"""Create an aging-by-week chart for TechPass tickets."""

from __future__ import annotations

import argparse
import csv
from collections import Counter
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches
import plotly.graph_objects as go
from plotly.subplots import make_subplots

AGING_BUCKETS = ("Week 1", "Week 2", "Week 3", "Week 4", "Week 5", "Above 5 Weeks")


def parse_args() -> argparse.Namespace:
    import sys
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
    from sa_paths import output_dir_for

    out_dir = output_dir_for("techpass-aging")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        type=Path,
        default=None,
        help="Path to the TechPass workbook. Defaults to the newest .xlsx file in the shared SA\\Inputs folder.",
    )
    parser.add_argument(
        "--as-of",
        type=date.fromisoformat,
        default=date.today(),
        help="Calculation date in YYYY-MM-DD format (default: today).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=out_dir / "techpass_aging_by_architect.html",
        help="Output HTML chart path.",
    )
    parser.add_argument(
        "--pptx-output",
        type=Path,
        default=out_dir / "techpass_aging_summary.pptx",
        help="Output one-slide PowerPoint path.",
    )
    args = parser.parse_args()
    if args.input is None:
        from sa_paths import latest_input
        found = latest_input("*.xlsx")
        if not found:
            raise SystemExit("No --input given and no .xlsx file found in the shared SA\\Inputs folder.")
        print(f"No --input given; using newest file in SA\\Inputs: {found}")
        args.input = found
    return args


def load_tickets(input_path: Path, as_of: date) -> list[dict]:
    # Some exported workbooks declare incorrect row spans, which causes
    # openpyxl's optimized read_only parser to truncate rows/columns.
    # Loading normally avoids that issue.
    workbook = load_workbook(input_path, read_only=False, data_only=True)
    worksheet = workbook.active
    headers = [cell.value for cell in next(worksheet.iter_rows())]
    required = {"Number", "State", "Assigned to", "Created"}
    missing = required - set(headers)
    if missing:
        raise ValueError(f"Missing required column(s): {', '.join(sorted(missing))}")

    tickets = []
    for values in worksheet.iter_rows(values_only=True):
        row = dict(zip(headers, values))
        if not row.get("Number") or row.get("Number") == "Number":
            continue
        created = row.get("Created")
        if isinstance(created, datetime):
            created = created.date()
        if not isinstance(created, date):
            raise ValueError(f"Invalid Created date for {row.get('Number')}: {created!r}")
        age_days = max(0, (as_of - created).days)
        row["Created"] = created
        row["Age (days)"] = age_days
        row["Age (weeks)"] = age_days // 7
        row["Aging bucket"] = AGING_BUCKETS[min(age_days // 7, 5)]
        row["Solution Architect"] = str(row.get("Assigned to") or "Unassigned").strip()
        row["Status"] = str(row.get("State") or "Unknown").strip()
        tickets.append(row)
    return tickets


def write_summary(tickets: list[dict], output_path: Path) -> Path:
    summary_path = output_path.with_name("techpass_aging_summary.csv")
    with summary_path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(["Solution Architect", *AGING_BUCKETS, "Total Tickets"])
        counts = Counter((ticket["Solution Architect"], ticket["Aging bucket"]) for ticket in tickets)
        architects = sorted({ticket["Solution Architect"] for ticket in tickets})
        for architect in architects:
            bucket_counts = [counts[(architect, bucket)] for bucket in AGING_BUCKETS]
            writer.writerow([architect, *bucket_counts, sum(bucket_counts)])
    return summary_path


def create_chart(tickets: list[dict], as_of: date, output_path: Path) -> go.Figure:
    architects = sorted({ticket["Solution Architect"] for ticket in tickets})
    colors = ["#2a9d8f", "#4f9d69", "#e9c46a", "#f4a261", "#e76f51", "#9b5de5"]
    counts = Counter((ticket["Solution Architect"], ticket["Aging bucket"]) for ticket in tickets)
    figure = make_subplots(
        rows=2,
        cols=1,
        row_heights=[0.54, 0.46],
        vertical_spacing=0.16,
        specs=[[{"type": "scatter"}], [{"type": "table"}]],
        subplot_titles=("Ticket counts by Solution Architect and aging week", "Ticket count by aging bucket"),
    )
    for bucket, color in zip(AGING_BUCKETS, colors):
        points = [
            (architect, counts[(architect, bucket)])
            for architect in architects
            if counts[(architect, bucket)]
        ]
        figure.add_trace(
            go.Scatter(
                name=bucket,
                x=[point[0] for point in points],
                y=[bucket for _ in points],
                mode="markers+text",
                text=[point[1] for point in points],
                textposition="middle center",
                textfont={"color": "#ffffff", "size": 12},
                marker={"size": 28, "color": color, "line": {"width": 1, "color": "#ffffff"}},
                customdata=[point[1] for point in points],
                hovertemplate="<b>%{x}</b><br>Aging: %{y}<br>Tickets: %{customdata}<extra></extra>",
            ),
            row=1,
            col=1,
        )

    figure.add_trace(
        go.Table(
            header={
                "values": ["<b>Solution Architect</b>", *[f"<b>{bucket}</b>" for bucket in AGING_BUCKETS], "<b>Total</b>"],
                "fill_color": "#264653",
                "font": {"color": "white", "size": 12},
                "align": "left",
                "height": 42,
            },
            cells={
                "values": [
                    architects,
                    *[
                        [sum(1 for ticket in tickets if ticket["Solution Architect"] == architect and ticket["Aging bucket"] == bucket) for architect in architects]
                        for bucket in AGING_BUCKETS
                    ],
                    [sum(1 for ticket in tickets if ticket["Solution Architect"] == architect) for architect in architects],
                ],
                "fill_color": [["#f1f5f4" if index % 2 == 0 else "#ffffff" for index in range(len(architects))]],
                "font": {"color": "#24323d", "size": 12},
                "align": "left",
                "height": 40,
            },
        ),
        row=2,
        col=1,
    )
    figure.update_layout(
        title=f"TechPass Ticket Aging | {len(tickets)} tickets | As of {as_of:%Y-%m-%d}",
        template="plotly_white",
        height=max(1150, 900 + len(architects) * 24),
        font={"family": "Aptos, Segoe UI, sans-serif", "size": 15, "color": "#24323d"},
        paper_bgcolor="#f6f3ee",
        plot_bgcolor="#ffffff",
        legend_title_text="Aging bucket",
        margin={"l": 95, "r": 35, "t": 100, "b": 130},
    )
    figure.update_xaxes(title_text="Solution Architect", type="category", categoryorder="array", categoryarray=architects, tickangle=-35, row=1, col=1)
    figure.update_yaxes(title_text="Aging week", type="category", categoryorder="array", categoryarray=list(AGING_BUCKETS), row=1, col=1)
    figure.write_html(output_path, include_plotlyjs=True, full_html=True)
    return figure


def create_powerpoint(tickets: list[dict], as_of: date, output_path: Path) -> None:
    architects = sorted({ticket["Solution Architect"] for ticket in tickets})
    counts = Counter((ticket["Solution Architect"], ticket["Aging bucket"]) for ticket in tickets)
    colors = ["#2a9d8f", "#4f9d69", "#e9c46a", "#f4a261", "#e76f51", "#9b5de5"]
    canvas = Image.new("RGB", (1800, 1000), "#f6f3ee")
    draw = ImageDraw.Draw(canvas)
    font_path = r"C:\Windows\Fonts\segoeui.ttf"
    bold_path = r"C:\Windows\Fonts\segoeuib.ttf"
    font = ImageFont.truetype(font_path, 24)
    small_font = ImageFont.truetype(font_path, 19)
    bold_font = ImageFont.truetype(bold_path, 25)
    title_font = ImageFont.truetype(bold_path, 34)
    draw.text((45, 28), f"TechPass Ticket Aging | {len(tickets)} tickets | As of {as_of:%Y-%m-%d}", fill="#24323d", font=title_font)

    chart_left, chart_top, chart_right, chart_bottom = 80, 125, 1720, 545
    plot_left, plot_top, plot_right, plot_bottom = 300, 180, 1660, 455
    draw.text((chart_left, 115), "Ticket counts by Solution Architect and aging week", fill="#24323d", font=bold_font)
    for index, bucket in enumerate(AGING_BUCKETS):
        y = plot_top + index * (plot_bottom - plot_top) / (len(AGING_BUCKETS) - 1)
        draw.line((plot_left, y, plot_right, y), fill="#d8dfdc", width=2)
        draw.text((chart_left, y - 13), bucket, fill="#24323d", font=small_font)
    for index, architect in enumerate(architects):
        x = plot_left + index * (plot_right - plot_left) / max(1, len(architects) - 1)
        draw.line((x, plot_top, x, plot_bottom), fill="#edf1ef", width=1)
        draw.text((x, plot_bottom + 18), architect, fill="#24323d", font=small_font, anchor="ma")
        for bucket_index, (bucket, color) in enumerate(zip(AGING_BUCKETS, colors)):
            count = counts[(architect, bucket)]
            if count:
                y = plot_top + bucket_index * (plot_bottom - plot_top) / (len(AGING_BUCKETS) - 1)
                radius = 27
                draw.ellipse((x - radius, y - radius, x + radius, y + radius), fill=color, outline="white", width=2)
                draw.text((x, y), str(count), fill="white", font=bold_font, anchor="mm")
    draw.line((plot_left, plot_bottom, plot_right, plot_bottom), fill="#24323d", width=2)

    table_left, table_top, table_right = 80, 610, 1720
    draw.text((table_left, 560), "Ticket count by aging bucket", fill="#24323d", font=bold_font)
    columns = ["Solution Architect", *AGING_BUCKETS, "Total"]
    widths = [310, 180, 180, 180, 180, 180, 220, 110]
    x_positions = [table_left]
    for width in widths[:-1]:
        x_positions.append(x_positions[-1] + width)
    row_height = 40
    for row_index, row in enumerate([columns] + [[architect, *[counts[(architect, bucket)] for bucket in AGING_BUCKETS], sum(counts[(architect, bucket)] for bucket in AGING_BUCKETS)] for architect in architects]):
        y = table_top + row_index * row_height
        fill = "#264653" if row_index == 0 else ("#ffffff" if row_index % 2 else "#f1f5f4")
        draw.rectangle((table_left, y, table_right, y + row_height), fill=fill, outline="#d8dfdc")
        for column_index, value in enumerate(row):
            draw.text((x_positions[column_index] + 8, y + row_height / 2), str(value), fill="white" if row_index == 0 else "#24323d", font=small_font, anchor="lm")

    chart_image = output_path.with_suffix(".png")
    canvas.save(chart_image)
    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    slide.shapes.add_picture(str(chart_image), Inches(0), Inches(0), width=Inches(13.333), height=Inches(7.5))
    presentation.save(output_path)
    chart_image.unlink()


def main() -> None:
    args = parse_args()
    tickets = load_tickets(args.input, args.as_of)
    if not tickets:
        raise SystemExit("No tickets were found in TP.xlsx.")
    summary_path = write_summary(tickets, args.output)
    figure = create_chart(tickets, args.as_of, args.output)
    create_powerpoint(tickets, args.as_of, args.pptx_output)
    print(f"Created {args.output}")
    print(f"Created {summary_path}")
    print(f"Created {args.pptx_output}")
    print(f"Tickets included: {len(tickets)}")


if __name__ == "__main__":
    main()