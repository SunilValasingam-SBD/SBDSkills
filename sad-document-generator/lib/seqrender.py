"""
Lightweight UML-style sequence diagram renderer using matplotlib.
Usage: render_sequence(participants, messages, title, out_path)
  participants: list of str (actor/system names, left to right)
  messages: list of tuples (from_name, to_name, label, style)
     style: 'sync' (solid, filled arrow), 'return' (dashed, open arrow), 'async' (solid, open arrow)
  notes: optional list of (participant_name_or_None, text, position_index) -> rendered as a note box
         placed above the message at that index (index = message index to insert before)
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, Rectangle
import textwrap

def render_sequence(participants, messages, title, out_path, notes=None, width=11):
    notes = notes or []
    n = len(participants)
    x_positions = {name: (i + 1) * (width / (n + 1)) for i, name in enumerate(participants)}

    # compute vertical layout: header + one row per message (plus extra rows for notes)
    row_h = 0.62
    header_h = 1.3
    note_rows = {}
    for note in notes:
        idx = note[2]
        note_rows[idx] = note_rows.get(idx, 0) + 1

    total_rows = len(messages) + sum(note_rows.values())
    fig_h = header_h + total_rows * row_h + 0.8
    fig, ax = plt.subplots(figsize=(width, fig_h))
    ax.set_xlim(0, width)
    ax.set_ylim(0, fig_h)
    ax.axis("off")

    top_y = fig_h - 0.5

    # Title
    ax.text(width / 2, fig_h - 0.15, title, ha="center", va="top", fontsize=13, fontweight="bold")

    # Participant boxes
    box_w = width / (n + 1) * 0.86
    box_h = 0.55
    for name in participants:
        x = x_positions[name]
        ax.add_patch(Rectangle((x - box_w / 2, top_y - box_h), box_w, box_h,
                                facecolor="#2E5C8A", edgecolor="black", zorder=3))
        wrapped = "\n".join(textwrap.wrap(name, 16))
        ax.text(x, top_y - box_h / 2, wrapped, ha="center", va="center",
                fontsize=9, color="white", fontweight="bold", zorder=4)

    lifeline_top = top_y - box_h
    lifeline_bottom = 0.6
    for name in participants:
        x = x_positions[name]
        ax.plot([x, x], [lifeline_top, lifeline_bottom], color="#888888",
                linestyle=(0, (4, 3)), linewidth=1, zorder=1)

    # Draw messages top-down
    cursor_y = lifeline_top - 0.25
    msg_idx = 0
    for note_before in [n for n in notes if n[2] == 0]:
        cursor_y = _draw_note(ax, note_before, x_positions, width, cursor_y, row_h)

    for msg in messages:
        frm, to, label, style = msg
        y = cursor_y
        x1, x2 = x_positions[frm], x_positions[to]
        dashed = style == "return"
        arrowstyle = "-|>" if style != "async" else "->"
        color = "#333333" if style != "return" else "#555555"
        if x1 == x2:
            # self-message loop
            ax.annotate("", xy=(x1 + 0.5, y - 0.18), xytext=(x1, y),
                        arrowprops=dict(arrowstyle=arrowstyle, color=color,
                                        linestyle="dashed" if dashed else "solid",
                                        connectionstyle="arc3,rad=-1.2"))
            ax.text(x1 + 0.15, y + 0.18, label, fontsize=8, ha="left", va="bottom")
        else:
            arrow = FancyArrowPatch((x1, y), (x2, y), arrowstyle=arrowstyle,
                                    color=color, mutation_scale=14,
                                    linestyle="dashed" if dashed else "solid",
                                    linewidth=1.3, zorder=2)
            ax.add_patch(arrow)
            mid = (x1 + x2) / 2
            ax.text(mid, y + 0.08, label, fontsize=8, ha="center", va="bottom", zorder=5)
        cursor_y -= row_h
        msg_idx += 1
        for note_after in [nt for nt in notes if nt[2] == msg_idx]:
            cursor_y = _draw_note(ax, note_after, x_positions, width, cursor_y, row_h)

    plt.tight_layout()
    plt.savefig(out_path, dpi=170, bbox_inches="tight")
    plt.close(fig)


def _draw_note(ax, note, x_positions, width, cursor_y, row_h):
    participant, text, _ = note
    wrapped = "\n".join(textwrap.wrap(text, 90))
    lines = wrapped.count("\n") + 1
    h = 0.28 * lines + 0.15
    if participant and participant in x_positions:
        cx = x_positions[participant]
    else:
        cx = width / 2
    box_w = min(width * 0.9, 0.42 * len(max(wrapped.split("\n"), key=len)) + 0.6)
    ax.add_patch(Rectangle((cx - box_w / 2, cursor_y - h), box_w, h,
                            facecolor="#FFF4CC", edgecolor="#C9A227", zorder=3))
    ax.text(cx, cursor_y - h / 2, wrapped, ha="center", va="center", fontsize=7.5,
            style="italic", zorder=4)
    return cursor_y - h - 0.18
