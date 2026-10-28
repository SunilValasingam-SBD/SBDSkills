# -*- coding: utf-8 -*-
"""
Generic input-document text extractor for the SAD document generator skill.

Supports: .docx, .pptx, .xlsx/.xlsm, .pdf, .txt/.md
Usage:
    py extract_docs.py --input "path\to\file1.docx" "path\to\file2.pptx" --out-dir "path\to\extract"
    py extract_docs.py --input-dir "path\to\folder" --out-dir "path\to\extract"

Each input file is converted to a plain-text .txt file (same base name) in --out-dir,
preserving headings/paragraph styles, table contents, and slide/sheet structure so the
model can read the raw content easily.
"""
import argparse
import os
import sys


def extract_docx(path):
    import docx
    d = docx.Document(path)
    lines = []
    for p in d.paragraphs:
        style = p.style.name if p.style else ""
        if p.text.strip():
            lines.append(f"[{style}] {p.text}")
    for t_idx, table in enumerate(d.tables):
        lines.append(f"--- TABLE {t_idx + 1} ---")
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            lines.append(" | ".join(cells))
    return "\n".join(lines)


def extract_pptx(path):
    import pptx
    p = pptx.Presentation(path)
    lines = []
    for i, slide in enumerate(p.slides):
        lines.append(f"--- SLIDE {i + 1} ---")
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    text = "".join(run.text for run in para.runs)
                    if text.strip():
                        lines.append(text)
            if shape.has_table:
                lines.append("[TABLE]")
                for row in shape.table.rows:
                    cells = [c.text.strip() for c in row.cells]
                    lines.append(" | ".join(cells))
            if shape.shape_type == 13:  # picture
                lines.append(f"[IMAGE: {shape.name}]")
    return "\n".join(lines)


def extract_xlsx(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    lines = []
    for ws in wb.worksheets:
        lines.append(f"--- SHEET: {ws.title} ---")
        for row in ws.iter_rows(values_only=True):
            if any(c is not None and str(c).strip() != "" for c in row):
                lines.append(" | ".join("" if c is None else str(c) for c in row))
    return "\n".join(lines)


def extract_pdf(path):
    import pypdf
    reader = pypdf.PdfReader(path)
    lines = []
    for i, page in enumerate(reader.pages):
        lines.append(f"--- PAGE {i + 1} ---")
        lines.append(page.extract_text() or "")
    return "\n".join(lines)


def extract_txt(path):
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        return f.read()


def extract_msg(path):
    import extract_msg
    msg = extract_msg.Message(path)
    lines = [
        f"From: {msg.sender}",
        f"To: {msg.to}",
        f"Cc: {msg.cc}" if msg.cc else "",
        f"Date: {msg.date}",
        f"Subject: {msg.subject}",
        "--- BODY ---",
        msg.body or "",
    ]
    if msg.attachments:
        lines.append("--- ATTACHMENTS ---")
        for att in msg.attachments:
            lines.append(f"[ATTACHMENT: {att.longFilename or att.shortFilename}]")
    msg.close()
    return "\n".join(l for l in lines if l)


EXTRACTORS = {
    ".docx": extract_docx,
    ".pptx": extract_pptx,
    ".xlsx": extract_xlsx,
    ".xlsm": extract_xlsx,
    ".pdf": extract_pdf,
    ".txt": extract_txt,
    ".md": extract_txt,
    ".msg": extract_msg,
}


def gather_files(inputs, input_dir):
    files = list(inputs or [])
    if input_dir:
        for root, _, names in os.walk(input_dir):
            for name in names:
                ext = os.path.splitext(name)[1].lower()
                if ext in EXTRACTORS:
                    files.append(os.path.join(root, name))
    return files


def main():
    ap = argparse.ArgumentParser(description="Extract text content from SAD input documents.")
    ap.add_argument("--input", nargs="*", default=[], help="One or more input file paths.")
    ap.add_argument("--input-dir", default=None, help="Folder to scan recursively for supported files.")
    ap.add_argument("--out-dir", required=True, help="Directory to write extracted .txt files to.")
    args = ap.parse_args()

    os.makedirs(args.out_dir, exist_ok=True)
    files = gather_files(args.input, args.input_dir)
    if not files:
        print("No input files found. Provide --input file(s) and/or --input-dir.", file=sys.stderr)
        sys.exit(1)

    manifest = []
    for in_path in files:
        ext = os.path.splitext(in_path)[1].lower()
        fn = EXTRACTORS.get(ext)
        base = os.path.splitext(os.path.basename(in_path))[0]
        out_path = os.path.join(args.out_dir, base + ".txt")
        if not fn:
            print(f"SKIP (unsupported type): {in_path}")
            continue
        try:
            text = fn(in_path)
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(text)
            print(f"OK: {os.path.basename(in_path)} -> {out_path} ({len(text)} chars)")
            manifest.append((in_path, out_path, len(text)))
        except Exception as e:
            print(f"FAIL: {in_path}: {e}")

    print(f"\nDONE. Extracted {len(manifest)}/{len(files)} files to {args.out_dir}")


if __name__ == "__main__":
    main()
