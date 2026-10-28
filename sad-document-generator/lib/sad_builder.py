# -*- coding: utf-8 -*-
"""
Reusable Word-document builder for Solution Architecture Documents (SAD), following the
SAD-Template.pdf structure (see ../SAD-Template-structure.md for the full section guidance text).

This is a generic library — it contains no project-specific content. A per-project script
imports SADDocument and calls its methods with that project's own extracted/synthesized content.

Basic usage:
    from sad_builder import SADDocument

    doc = SADDocument()
    doc.cover_page(
        title="Solution Architecture Document",
        subtitle="<Project Name> - <Short Description> (<Project Ref>)",
        meta_rows=[("Project Name", "..."), ("Business Unit", "..."), ...],
    )
    doc.heading("1. Solution Design Summary", 1)
    doc.heading("1.1 Solution Design Overview", 2)
    doc.para("...")
    ...
    doc.add_sequence_diagram("2.2.1 <Requirement Title> (BR#)", "<description>",
                             image_path=".../01_xxx.png", wsd_script_path=".../01_xxx.wsd.txt")
    ...
    doc.save(r"C:\\path\\to\\Project Name - Solution Architecture Document.docx")

All lower-level helpers (heading, para, bullets, table, image, wsd_script) are exposed directly
so any section of the template can be authored freely; a few convenience wrappers are provided
for the most structurally repetitive parts (cover page, sequence diagram blocks).
"""
import os
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SBD_BLUE = RGBColor(0x1F, 0x3B, 0x5C)
SBD_ACCENT = RGBColor(0x2E, 0x5C, 0x8A)

# The nine standard section headings from SAD-Template.pdf, for reference/validation.
TEMPLATE_SECTIONS = [
    "1 Solution Design Summary",
    "1.1 Solution Design Overview",
    "1.2 Solution Design Constraints",
    "1.3 Solution Design Risks / Contingencies",
    "2 Solution Design",
    "2.1 Solution Architecture Diagram",
    "2.2 Process Sequence Diagrams",
    "2.3 Solution Design Details",
    "2.3.1 Systems",
    "2.3.2 System Connections",
    "2.3.3 Integration Approaches",
    "2.3.4 Middleware & Transfer Technologies",
    "2.3.5 Data Objects",
    "2.3.6 Data Strategy",
    "2.3.7 System Hosting",
    "2.3.8 System Existence",
    "2.3.9 Deployment & Support",
    "2.4 Solution Design Interface Inventory",
]


class SADDocument:
    def __init__(self, base_font="Calibri", base_size=10.5):
        self.doc = docx.Document()
        style = self.doc.styles["Normal"]
        style.font.name = base_font
        style.font.size = Pt(base_size)

    # ---------- low-level helpers ----------
    def heading(self, text, level, color=SBD_BLUE):
        h = self.doc.add_heading(text, level=level)
        for run in h.runs:
            run.font.color.rgb = color
        return h

    def para(self, text, bold=False, italic=False, size=None):
        p = self.doc.add_paragraph()
        r = p.add_run(text)
        r.bold = bold
        r.italic = italic
        if size:
            r.font.size = Pt(size)
        return p

    def bullets(self, items):
        for it in items:
            self.doc.add_paragraph(it, style="List Bullet")

    def numbered(self, items):
        for it in items:
            self.doc.add_paragraph(it, style="List Number")

    def table(self, headers, rows, header_fill="1F3B5C"):
        table = self.doc.add_table(rows=1, cols=len(headers))
        table.style = "Light Grid Accent 1"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr_cells = table.rows[0].cells
        for i, h in enumerate(headers):
            hdr_cells[i].text = ""
            p = hdr_cells[i].paragraphs[0]
            r = p.add_run(h)
            r.bold = True
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            shd = OxmlElement('w:shd')
            shd.set(qn('w:fill'), header_fill)
            hdr_cells[i]._tc.get_or_add_tcPr().append(shd)
        for row in rows:
            cells = table.add_row().cells
            for i, val in enumerate(row):
                cells[i].text = str(val)
                for p in cells[i].paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(9)
        return table

    def image(self, path, width=6.4, caption=None):
        self.doc.add_picture(path, width=Inches(width))
        last = self.doc.paragraphs[-1]
        last.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if caption:
            cap = self.doc.add_paragraph()
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = cap.add_run(caption)
            r.italic = True
            r.font.size = Pt(9)

    def wsd_script(self, title, script_path):
        self.para(f"WebSequenceDiagrams.com editable script - {title}:", italic=True, size=9)
        with open(script_path, "r", encoding="utf-8") as f:
            script = f.read()
        p = self.doc.add_paragraph()
        r = p.add_run(script)
        r.font.name = "Consolas"
        r.font.size = Pt(8.5)
        p.paragraph_format.left_indent = Inches(0.3)

    def page_break(self):
        self.doc.add_page_break()

    # ---------- structural convenience wrappers ----------
    def cover_page(self, title, subtitle, meta_rows):
        title_p = self.doc.add_paragraph()
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = title_p.add_run(title)
        r.font.size = Pt(28)
        r.bold = True
        r.font.color.rgb = SBD_BLUE

        sub_p = self.doc.add_paragraph()
        sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = sub_p.add_run(subtitle)
        r.font.size = Pt(15)
        r.font.color.rgb = SBD_ACCENT

        self.doc.add_paragraph()
        self.table(["Field", "Value"], meta_rows)
        self.page_break()

    def add_sequence_diagram(self, heading_text, description, image_path=None,
                              wsd_script_path=None, level=3):
        """One '2.2.x <High-level requirement> (BR#)' block: heading, description, image,
        and (optionally) the editable websequencediagrams.com script as a fallback/companion."""
        self.heading(heading_text, level)
        if description:
            self.para(description)
        if image_path and os.path.exists(image_path):
            self.image(image_path, width=6.3)
        if wsd_script_path and os.path.exists(wsd_script_path):
            self.wsd_script(heading_text, wsd_script_path)
        self.doc.add_paragraph()

    def save(self, out_path):
        os.makedirs(os.path.dirname(out_path), exist_ok=True) if os.path.dirname(out_path) else None
        self.doc.save(out_path)
        return out_path
