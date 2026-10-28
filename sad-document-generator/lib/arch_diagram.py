# -*- coding: utf-8 -*-
"""
Generic architecture-diagram drawing helpers (current/future state, Level 2 style boxes & arrows),
used when Ardoq export images are not available and a conceptual substitute diagram is needed.

Usage pattern (in a per-project script):
    from arch_diagram import new_figure, box, arrow, save

    fig, ax = new_figure(11, 7, "Current State - Conceptual Architecture")
    box(ax, 4.2, 5.0, 2.6, 0.8, "System Name\nRole description")
    arrow(ax, (5.5, 5.0), (5.5, 4.2), "data flow label")
    save(fig, r"out\arch_current_state.png")
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import textwrap

PALETTE = {
    "core": "#1F3B5C",       # primary/target system
    "existing": "#2E5C8A",   # existing SBD-standard system
    "satellite": "#5B8BB0",  # satellite / secondary system
    "legacy": "#8A8A8A",     # legacy / retiring / future-roadmap
}


def new_figure(width, height, title):
    fig, ax = plt.subplots(figsize=(width, height))
    ax.set_xlim(0, width)
    ax.set_ylim(0, height)
    ax.axis("off")
    ax.text(width / 2, height - 0.3, title, ha="center", fontsize=13, fontweight="bold")
    return fig, ax


def box(ax, x, y, w, h, text, color="core", fontsize=9, textcolor="white"):
    fill = PALETTE.get(color, color)
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.06",
                                 facecolor=fill, edgecolor="black", linewidth=1, zorder=3))
    wrapped = "\n".join(textwrap.wrap(text, 18))
    ax.text(x + w / 2, y + h / 2, wrapped, ha="center", va="center", fontsize=fontsize,
            color=textcolor, fontweight="bold", zorder=4)


def arrow(ax, p1, p2, label="", style="-|>", color="#333333"):
    a = FancyArrowPatch(p1, p2, arrowstyle=style, color=color, mutation_scale=14,
                         linewidth=1.3, zorder=2, connectionstyle="arc3,rad=0.05")
    ax.add_patch(a)
    if label:
        mx, my = (p1[0] + p2[0]) / 2, (p1[1] + p2[1]) / 2
        ax.text(mx, my + 0.12, label, fontsize=7.5, ha="center", va="bottom", zorder=5)


def save(fig, out_path, dpi=170):
    plt.tight_layout()
    fig.savefig(out_path, dpi=dpi, bbox_inches="tight")
    plt.close(fig)
