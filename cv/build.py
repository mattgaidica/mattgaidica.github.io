#!/usr/bin/env python3
"""Build WashU Medicine CV outputs from cv.yaml.

Produces:
  - GAIDICA_WashU_CV.docx  (primary submission format)
  - GAIDICA_WashU_CV.pdf   (via pandoc + XeLaTeX/tectonic)
  - _build/cv.md           (intermediate markdown)

Usage:
  python build.py              # docx + pdf
  python build.py --docx-only
  python build.py --pdf-only
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

import yaml
from docx import Document
from docx.enum.text import WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, Twips

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "cv.yaml"
DOCX_OUT = ROOT / "GAIDICA_WashU_CV.docx"
PDF_OUT = ROOT / "GAIDICA_WashU_CV.pdf"
BUILD_DIR = ROOT / "_build"
MD_OUT = BUILD_DIR / "cv.md"

# WashU-ish layout matching the existing template (~0.7" margins, date | body).
DATE_COL_IN = 1.25
CONTENT_COL_IN = 5.85
PAGE_WIDTH_IN = 8.5
MARGIN_IN = 0.7


def load_cv(path: Path = SOURCE) -> dict:
    with path.open(encoding="utf-8") as f:
        return yaml.safe_load(f)


def set_run_font(run, font: str, size_pt: float, bold: bool = False, underline: bool = False):
    run.bold = bold
    run.underline = underline
    run.font.name = font
    run.font.size = Pt(size_pt)
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)
    rFonts.set(qn("w:cs"), font)


def set_paragraph_spacing(paragraph, before: int = 0, after: int = 40, line: float = 1.0):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = line


def add_page_number_field(paragraph):
    """Insert PAGE and NUMPAGES fields into a paragraph."""

    def add_field(instr: str):
        run = paragraph.add_run()
        fld_begin = OxmlElement("w:fldChar")
        fld_begin.set(qn("w:fldCharType"), "begin")
        run._r.append(fld_begin)

        run = paragraph.add_run()
        instr_text = OxmlElement("w:instrText")
        instr_text.set(qn("xml:space"), "preserve")
        instr_text.text = f" {instr} "
        run._r.append(instr_text)

        run = paragraph.add_run()
        fld_separate = OxmlElement("w:fldChar")
        fld_separate.set(qn("w:fldCharType"), "separate")
        run._r.append(fld_separate)

        run = paragraph.add_run("1")

        run = paragraph.add_run()
        fld_end = OxmlElement("w:fldChar")
        fld_end.set(qn("w:fldCharType"), "end")
        run._r.append(fld_end)

    paragraph.add_run("Page ")
    add_field("PAGE")
    paragraph.add_run(" of ")
    add_field("NUMPAGES")


def configure_document(doc: Document, meta: dict):
    font = meta.get("font", "Times New Roman")
    size = float(meta.get("font_size_pt", 10.5))

    style = doc.styles["Normal"]
    style.font.name = font
    style.font.size = Pt(size)
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)

    for section in doc.sections:
        section.page_width = Inches(PAGE_WIDTH_IN)
        section.page_height = Inches(11)
        section.left_margin = Inches(MARGIN_IN)
        section.right_margin = Inches(MARGIN_IN)
        section.top_margin = Inches(MARGIN_IN)
        section.bottom_margin = Inches(0.65)
        section.different_first_page_header_footer = False

        footer = section.footer
        footer.is_linked_to_previous = False
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.clear()
        set_paragraph_spacing(p, before=0, after=0)
        run = p.add_run(f"{meta.get('footer_name', meta['name'])}\t{meta['revised']}\t")
        set_run_font(run, font, 9)
        # tab stops: left name, center date, right page
        tabs = p.paragraph_format.tab_stops
        usable = PAGE_WIDTH_IN - 2 * MARGIN_IN
        tabs.add_tab_stop(Inches(usable / 2), WD_TAB_ALIGNMENT.CENTER)
        tabs.add_tab_stop(Inches(usable), WD_TAB_ALIGNMENT.RIGHT)
        add_page_number_field(p)
        for run in p.runs:
            set_run_font(run, font, 9)


def add_heading(doc: Document, title: str, meta: dict, *, section: bool):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=10, after=2)
    run = p.add_run(title)
    set_run_font(
        run,
        meta.get("font", "Times New Roman"),
        float(meta.get("font_size_pt", 10.5)),
        bold=True,
        underline=section,
    )
    return p


def add_note(doc: Document, text: str, meta: dict):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=4, after=2)
    run = p.add_run(text)
    set_run_font(
        run,
        meta.get("font", "Times New Roman"),
        float(meta.get("font_size_pt", 10.5)),
        bold=True,
    )


def add_entry_table(doc: Document, entries: list[dict], meta: dict):
    if not entries:
        return
    font = meta.get("font", "Times New Roman")
    size = float(meta.get("font_size_pt", 10.5))
    table = doc.add_table(rows=0, cols=2)
    table.autofit = False
    table.allow_autofit = False

    # Prefer fixed column widths (twips).
    date_tw = Twips(int(DATE_COL_IN * 1440))
    body_tw = Twips(int(CONTENT_COL_IN * 1440))

    for entry in entries:
        row = table.add_row()
        c0, c1 = row.cells
        c0.width = date_tw
        c1.width = body_tw

        # Clear default paragraph and write date.
        p0 = c0.paragraphs[0]
        p0.clear()
        set_paragraph_spacing(p0, before=2, after=2)
        r0 = p0.add_run(str(entry.get("dates", "") or ""))
        set_run_font(r0, font, size)

        text = str(entry.get("text", "") or "")
        lines = text.split("\n") if text else [""]
        p1 = c1.paragraphs[0]
        p1.clear()
        set_paragraph_spacing(p1, before=2, after=2)
        for i, line in enumerate(lines):
            if i:
                p1.add_run().add_break()
            r = p1.add_run(line)
            set_run_font(r, font, size)

    # Remove table borders for WashU look.
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)

    # Small spacer after table.
    spacer = doc.add_paragraph()
    set_paragraph_spacing(spacer, before=0, after=2)


def add_bibliography(doc: Document, items: list[str], meta: dict):
    font = meta.get("font", "Times New Roman")
    size = float(meta.get("font_size_pt", 10.5))
    for i, cite in enumerate(items, start=1):
        p = doc.add_paragraph()
        set_paragraph_spacing(p, before=2, after=6)
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.first_line_indent = Inches(-0.25)
        run = p.add_run(f"{i}) {cite}")
        set_run_font(run, font, size)


def build_docx(data: dict, dest: Path = DOCX_OUT) -> Path:
    meta = data["meta"]
    doc = Document()
    configure_document(doc, meta)

    # Title + revision date
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=0, after=0)
    run = p.add_run(meta["name"])
    set_run_font(run, meta.get("font", "Times New Roman"), 14, bold=True)

    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=0, after=8)
    run = p.add_run(str(meta["revised"]))
    set_run_font(run, meta.get("font", "Times New Roman"), float(meta.get("font_size_pt", 10.5)))

    for block in data.get("body", []):
        btype = block.get("type")
        if btype in ("section", "subsection"):
            add_heading(doc, block["title"], meta, section=(btype == "section"))
            add_entry_table(doc, block.get("entries") or [], meta)
        elif btype == "bibliography":
            add_heading(doc, block["title"], meta, section=False)
            add_bibliography(doc, block.get("items") or [], meta)
        elif btype == "note":
            add_note(doc, block["text"], meta)
        elif btype == "entries":
            add_entry_table(doc, block.get("entries") or [], meta)

    dest.parent.mkdir(parents=True, exist_ok=True)
    doc.save(dest)
    return dest


def yaml_to_markdown(data: dict) -> str:
    meta = data["meta"]
    lines = [
        f"# {meta['name']}",
        "",
        str(meta["revised"]),
        "",
    ]
    for block in data.get("body", []):
        btype = block.get("type")
        if btype == "section":
            lines += [f"## {block['title']}", ""]
        elif btype in ("subsection", "bibliography"):
            lines += [f"### {block['title']}", ""]
        elif btype == "note":
            lines += [f"*{block['text']}*", ""]

        for entry in block.get("entries") or []:
            dates = entry.get("dates", "")
            text = str(entry.get("text", "")).replace("\n", "  \n")
            lines.append(f"**{dates}**  ")
            lines.append(text)
            lines.append("")

        for i, cite in enumerate(block.get("items") or [], start=1):
            lines.append(f"{i}. {cite}")
            lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def build_markdown(data: dict, dest: Path = MD_OUT) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(yaml_to_markdown(data), encoding="utf-8")
    return dest


def find_pdf_engine() -> str | None:
    for engine in ("xelatex", "tectonic", "pdflatex", "lualatex"):
        if shutil.which(engine):
            return engine
    return None


def build_pdf(md_path: Path, dest: Path = PDF_OUT) -> Path:
    if not shutil.which("pandoc"):
        raise RuntimeError("pandoc not found; install pandoc to build PDF")
    engine = find_pdf_engine()
    if not engine:
        raise RuntimeError("No LaTeX engine found (xelatex/tectonic/pdflatex)")

    # CI often lacks Times New Roman; override with CV_MAINFONT (e.g. Liberation Serif).
    mainfont = os.environ.get("CV_MAINFONT", "Times New Roman")

    cmd = [
        "pandoc",
        str(md_path),
        "-o",
        str(dest),
        f"--pdf-engine={engine}",
        "-V",
        "geometry:margin=0.7in",
        "-V",
        "fontsize=11pt",
        "-V",
        f"mainfont={mainfont}",
        "--shift-heading-level-by=0",
    ]
    # tectonic does not use all -V the same way; still fine as best-effort
    subprocess.run(cmd, check=True, cwd=ROOT)
    return dest


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build WashU CV from cv.yaml")
    parser.add_argument("--docx-only", action="store_true")
    parser.add_argument("--pdf-only", action="store_true")
    parser.add_argument("--source", type=Path, default=SOURCE)
    args = parser.parse_args(argv)

    if not args.source.exists():
        print(f"Missing source: {args.source}", file=sys.stderr)
        return 1

    data = load_cv(args.source)
    do_docx = not args.pdf_only
    do_pdf = not args.docx_only

    md_path = build_markdown(data)
    print(f"Wrote {md_path.relative_to(ROOT)}")

    if do_docx:
        path = build_docx(data)
        print(f"Wrote {path.relative_to(ROOT)}")

    if do_pdf:
        try:
            path = build_pdf(md_path)
            print(f"Wrote {path.relative_to(ROOT)}")
        except (RuntimeError, subprocess.CalledProcessError) as exc:
            print(f"PDF build failed: {exc}", file=sys.stderr)
            return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
