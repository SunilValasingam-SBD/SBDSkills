"""
TEMPLATE / TOOLKIT for building professional swimlane process-flow diagrams as an
editable PowerPoint (.pptx) using python-pptx native shapes and connectors.

This file is a starting point, not a fixed script: copy it, then edit the `build_slide()`
calls at the bottom (lane names/heights, node text/positions, connections) to match the
process you were given. Keep the helper functions (add_box, connect, draw_lane, add_title,
add_legend, set_shadow_soft) as-is unless you have a specific reason to change them --
they encode fixes for real rendering bugs (see NOTES at the bottom of this file).

Usage:
    py build_swimlane_pptx.py
    py render_and_verify.py   # REQUIRED self-check step before delivering, see SKILL.md
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
import os

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_PATH = os.path.join(OUT_DIR, "Flow_Diagram.pptx")

FONT = "Calibri"
NAVY = RGBColor(0x1F, 0x38, 0x64)
DARK_TEXT = RGBColor(0x22, 0x26, 0x2B)
GRAY_TEXT = RGBColor(0x6B, 0x71, 0x78)
LINE_COLOR = RGBColor(0x59, 0x63, 0x6E)

# Add/rename lanes and give each a fill + tab (accent) color.
LANE_FILL = {
    "Lane A": RGBColor(0xEB, 0xF2, 0xFA),
    "Lane B": RGBColor(0xEC, 0xF6, 0xEE),
    "Lane C": RGBColor(0xFD, 0xF5, 0xE4),
    "Outcome": RGBColor(0xF2, 0xF2, 0xF3),
}
LANE_TAB = {
    "Lane A": RGBColor(0x2E, 0x5A, 0x8C),
    "Lane B": RGBColor(0x3E, 0x7D, 0x4E),
    "Lane C": RGBColor(0xB4, 0x88, 0x0C),
    "Outcome": RGBColor(0x5B, 0x64, 0x70),
}
KIND_STYLE = {
    "process": (RGBColor(0xFF, 0xFF, 0xFF), RGBColor(0x2E, 0x5A, 0x8C)),
    "decision": (RGBColor(0xFD, 0xF3, 0xD9), RGBColor(0xB4, 0x88, 0x0C)),
    "stop": (RGBColor(0xE9, 0xE9, 0xEB), RGBColor(0x6B, 0x71, 0x78)),
    "success": (RGBColor(0xE5, 0xF3, 0xE8), RGBColor(0x2E, 0x7D, 0x32)),
    "error": (RGBColor(0xFB, 0xEA, 0xE9), RGBColor(0xC0, 0x39, 0x2B)),
}
KIND_SHAPE = {
    "process": MSO_SHAPE.ROUNDED_RECTANGLE,
    "decision": MSO_SHAPE.DIAMOND,
    "stop": MSO_SHAPE.ROUNDED_RECTANGLE,
    "success": MSO_SHAPE.ROUNDED_RECTANGLE,
    "error": MSO_SHAPE.ROUNDED_RECTANGLE,
}

# IMPORTANT: verified empirically against real PowerPoint rendering (see NOTES at bottom).
# Do NOT assume 0=top,1=right,2=bottom,3=left -- that assumption is WRONG and causes
# connectors to attach to the opposite side, cutting straight through box text.
CXN_TOP, CXN_RIGHT, CXN_BOTTOM, CXN_LEFT = 0, 3, 2, 1


def set_shadow_soft(shape):
    """Apply a soft outer shadow via raw XML (subtle, professional look)."""
    el = shape._element.spPr
    effectLst = el.makeelement(qn("a:effectLst"), {})
    shadow = el.makeelement(qn("a:outerShdw"), {
        "blurRad": "45000", "dist": "20000", "dir": "5400000", "rotWithShape": "0"
    })
    clr = el.makeelement(qn("a:srgbClr"), {"val": "1F2937"})
    alpha = el.makeelement(qn("a:alpha"), {"val": "22000"})
    clr.append(alpha)
    shadow.append(clr)
    effectLst.append(shadow)
    el.append(effectLst)


def add_box(slide, x, y, w, h, kind, text, font_size=12):
    """Add a shape with word-wrapped, vertically-centered text. Height should already
    account for the number of wrapped lines -- undersized boxes cause text to spill
    past the shape edge (looks like a rendering bug but is really a sizing bug)."""
    fill_color, line_color = KIND_STYLE[kind]
    shape = slide.shapes.add_shape(KIND_SHAPE[kind], Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.color.rgb = line_color
    shape.line.width = Pt(1.4)
    set_shadow_soft(shape)

    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.04)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lines = text.split("\n")
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.CENTER
    r0 = p0.add_run()
    r0.text = lines[0]
    r0.font.size = Pt(font_size)
    r0.font.name = FONT
    r0.font.color.rgb = DARK_TEXT
    for line in lines[1:]:
        p = tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = line
        r.font.size = Pt(font_size)
        r.font.name = FONT
        r.font.color.rgb = DARK_TEXT
    return shape


def connect(slide, shape_from, site_from, shape_to, site_to, label=None, label_pos=None):
    """Elbow connector bound to shape connection sites (auto-routes at right angles).
    CRITICAL: leave at least ~0.3-0.4in of clear space between shape_from and shape_to
    along the connection axis. If two shapes are touching or overlapping, the connector
    has zero length and becomes invisible -- it will look like the boxes are just
    sitting on top of each other with no line/arrow at all."""
    conn = slide.shapes.add_connector(MSO_CONNECTOR.ELBOW,
                                       Inches(0), Inches(0), Inches(1), Inches(1))
    conn.begin_connect(shape_from, site_from)
    conn.end_connect(shape_to, site_to)
    conn.line.color.rgb = LINE_COLOR
    conn.line.width = Pt(1.4)
    ln = conn.line._get_or_add_ln()
    tail = ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"})
    ln.append(tail)

    if label:
        lx, ly = label_pos
        tb = slide.shapes.add_textbox(Inches(lx), Inches(ly), Inches(0.9), Inches(0.3))
        tf = tb.text_frame
        tf.margin_left = Inches(0.02)
        tf.margin_right = Inches(0.02)
        tf.margin_top = Inches(0.0)
        tf.margin_bottom = Inches(0.0)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = label
        r.font.size = Pt(10.5)
        r.font.bold = True
        r.font.name = FONT
        r.font.color.rgb = NAVY
        tb.fill.solid()
        tb.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        tb.line.color.rgb = RGBColor(0xD7, 0xDB, 0xE0)
        tb.line.width = Pt(0.5)
    return conn


def draw_lane(slide, name, x, y, w, h):
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    rect.fill.solid()
    rect.fill.fore_color.rgb = LANE_FILL[name]
    rect.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    rect.line.width = Pt(2)
    rect.shadow.inherit = False

    tab_w, tab_h = 1.5, min(0.55, h * 0.42)
    tab = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Inches(x - tab_w - 0.15), Inches(y + h / 2 - tab_h / 2),
                                  Inches(tab_w), Inches(tab_h))
    tab.fill.solid()
    tab.fill.fore_color.rgb = LANE_TAB[name]
    tab.line.fill.background()
    set_shadow_soft(tab)
    tf = tab.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = name
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.name = FONT
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    return rect


def add_title(slide, x, w, y, title, subtitle):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(0.55))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = title
    r.font.size = Pt(26)
    r.font.bold = True
    r.font.name = FONT
    r.font.color.rgb = NAVY

    tb2 = slide.shapes.add_textbox(Inches(x), Inches(y + 0.52), Inches(w), Inches(0.4))
    tf2 = tb2.text_frame
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = subtitle
    r2.font.size = Pt(15)
    r2.font.name = FONT
    r2.font.color.rgb = GRAY_TEXT

    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y + 1.0),
                                       Inches(x + w), Inches(y + 1.0))
    line.line.color.rgb = RGBColor(0xD7, 0xDB, 0xE0)
    line.line.width = Pt(1.2)


def add_legend(slide, x_center, y, items=None):
    """items: list of (label, kind) tuples. Defaults to the 5 built-in kinds."""
    if items is None:
        items = [
            ("Process step", "process"),
            ("Decision point", "decision"),
            ("Stop / no action", "stop"),
            ("Success outcome", "success"),
            ("Error outcome", "error"),
        ]
    swatch = 0.22
    gap_after_swatch = 0.08
    gap_between = 0.35
    widths = [swatch + gap_after_swatch + 0.085 * len(lbl) for lbl, _ in items]
    total_w = sum(widths) + gap_between * (len(items) - 1)

    panel_w = total_w + 0.9
    panel_x = x_center - panel_w / 2
    panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(panel_x), Inches(y),
                                    Inches(panel_w), Inches(0.55))
    panel.fill.solid()
    panel.fill.fore_color.rgb = RGBColor(0xFA, 0xFA, 0xFB)
    panel.line.color.rgb = RGBColor(0xD7, 0xDB, 0xE0)
    panel.line.width = Pt(1)

    cx = x_center - total_w / 2
    cy = y + 0.16
    for (label, kind), w in zip(items, widths):
        fill_color, line_color = KIND_STYLE[kind]
        sw = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(cy),
                                     Inches(swatch), Inches(swatch))
        sw.fill.solid()
        sw.fill.fore_color.rgb = fill_color
        sw.line.color.rgb = line_color
        sw.line.width = Pt(1.2)

        tb = slide.shapes.add_textbox(Inches(cx + swatch + gap_after_swatch), Inches(cy - 0.05),
                                       Inches(w), Inches(0.32))
        tf = tb.text_frame
        tf.margin_left = Inches(0.0)
        tf.margin_top = Inches(0.0)
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = label
        r.font.size = Pt(11)
        r.font.name = FONT
        r.font.color.rgb = DARK_TEXT
        cx += w + gap_between


class LaneLayout:
    """Helper that lays out lanes top-to-bottom and gives you each lane's vertical center.
    Give decision diamonds and any lane with tall/stacked content extra height
    (>= node_height + 0.4in clearance on each side) so connectors never get squeezed
    to zero length between adjacent lanes."""

    def __init__(self, slide, lane_names, lane_heights, x, w, y0):
        self.tops = {}
        y = y0
        for name, h in zip(lane_names, lane_heights):
            draw_lane(slide, name, x, y, w, h)
            self.tops[name] = (y, h)
            y += h
        self.bottom = y

    def cy(self, name):
        top, h = self.tops[name]
        return top + h / 2


def build_example_slide(prs):
    """Illustrative example -- replace node text/positions/connections with your process."""
    slide_w = prs.slide_width / 914400
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    margin = 0.5
    lane_x = 2.1
    lane_w = slide_w - lane_x - margin
    add_title(slide, margin, slide_w - 2 * margin, 0.25, "Example Process Flow", "Subtitle here")

    lane_names = ["Lane A", "Lane B", "Lane C", "Outcome"]
    # Give the lane(s) holding decision diamonds or multi-line boxes extra height.
    lane_heights = [1.4, 1.6, 1.5, 1.15]
    lanes = LaneLayout(slide, lane_names, lane_heights, lane_x, lane_w, 1.55)

    n1 = add_box(slide, lane_x + 0.25, lanes.cy("Lane A") - 0.32, 1.7, 0.64, "process",
                 "1. Do something", 11)
    n2 = add_box(slide, lane_x + 2.55, lanes.cy("Lane B") - 0.65, 2.6, 1.3, "decision",
                 "2. Some condition?", 10.5)
    n_ok = add_box(slide, lane_x + 5.5, lanes.cy("Lane A") - 0.5, 2.0, 1.0, "success",
                    "3a. Success path", 10.5)
    n_err = add_box(slide, lane_x + 5.5, lanes.cy("Outcome") - 0.42, 2.0, 0.84, "error",
                     "3b. Error path", 10.5)

    connect(slide, n1, CXN_RIGHT, n2, CXN_LEFT)
    connect(slide, n2, CXN_TOP, n_ok, CXN_BOTTOM, "Yes",
            (n2.left / 914400 + 0.6, lanes.cy("Lane A") + 0.2))
    connect(slide, n2, CXN_BOTTOM, n_err, CXN_TOP, "No",
            (n2.left / 914400 + 0.6, lanes.cy("Lane B") + lane_heights[1] / 2 - 0.15))

    add_legend(slide, lane_x + lane_w / 2, lanes.bottom + 0.35)


def main():
    prs = Presentation()
    prs.slide_width = Inches(19.5)   # widen if content overflows; check in render_and_verify.py
    prs.slide_height = Inches(9.5)

    build_example_slide(prs)

    prs.save(OUT_PATH)
    print("Saved:", OUT_PATH)


if __name__ == "__main__":
    main()


# ---------------------------------------------------------------------------------------
# NOTES / LESSONS LEARNED (read before modifying the constants above)
# ---------------------------------------------------------------------------------------
# 1. Connection-site indices for ROUNDED_RECTANGLE and DIAMOND autoshapes are NOT
#    0=top,1=right,2=bottom,3=left as one might assume. Verified empirically via
#    PowerPoint COM (see render_and_verify.py) that the correct mapping is
#    TOP=0, RIGHT=3, BOTTOM=2, LEFT=1. Using the wrong indices makes connectors attach to
#    the opposite side of a shape, drawing a line straight through the shape's own text.
# 2. python-pptx writes elbow connectors with default (unrouted) bend points. Opening the
#    file with PowerPoint COM and calling shape.RerouteConnections() on every connector
#    (msoConnector shapes, Type == 2) forces PowerPoint to recompute clean right-angle
#    paths -- do this before exporting preview images or handing off the file.
# 3. Always leave >= 0.3-0.4in gap between any two shapes that are directly connected,
#    especially across adjacent lanes stacked vertically. If you enlarge one shape (e.g. a
#    decision diamond) for text-fit reasons, also enlarge its lane height accordingly --
#    otherwise the shape can end up touching its neighbor with a zero-length, invisible
#    connector (looks like "shapes just sitting on top of each other with no arrow").
# 4. Multi-line text boxes should have height computed from the wrapped line count, not a
#    fixed guess -- undersized boxes let text spill past the shape's edge.
# 5. ALWAYS self-verify by rendering to PNG (render_and_verify.py) and visually inspecting
#    before delivering. Do not assume the generated XML renders as intended.
