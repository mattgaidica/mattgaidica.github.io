#!/usr/bin/env python3
"""Build WashU Medicine CV outputs from cv.yaml.

Produces:
  - GAIDICA_WashU_CV.docx  (primary submission format)
  - GAIDICA_WashU_CV.pdf   (via pandoc + XeLaTeX/tectonic)
  - api/v1/cv.json         (machine-readable API)
  - api/v1/index.json      (API discovery)
  - _build/cv.md           (intermediate markdown)

Usage:
  python build.py              # api + docx + pdf
  python build.py --docx-only
  python build.py --pdf-only
  python build.py --api-only
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from datetime import datetime, timezone
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
API_DIR = ROOT / "api" / "v1"
API_CV_OUT = API_DIR / "cv.json"
API_INDEX_OUT = API_DIR / "index.json"
BUILD_DIR = ROOT / "_build"
MD_OUT = BUILD_DIR / "cv.md"

# Public site origin (CNAME). Override with CV_BASE_URL if needed.
BASE_URL = os.environ.get("CV_BASE_URL", "https://gaidi.ca").rstrip("/")

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


def _section_by_title(data: dict, title: str) -> dict | None:
    for block in data.get("body") or []:
        if block.get("title") == title:
            return block
    return None


def _entries_as_dict(block: dict | None) -> dict[str, str]:
    """Map Contact-style date labels → text (first line if multiline)."""
    out: dict[str, str] = {}
    if not block:
        return out
    for entry in block.get("entries") or []:
        key = str(entry.get("dates") or "").strip()
        text = str(entry.get("text") or "").strip()
        if key:
            out[key] = text
    return out


def public_formats() -> dict[str, str]:
    return {
        "json": f"{BASE_URL}/cv/api/v1/cv.json",
        "discovery": f"{BASE_URL}/cv/api/v1/index.json",
        "yaml": f"{BASE_URL}/cv/cv.yaml",
        "docx": f"{BASE_URL}/cv/GAIDICA_WashU_CV.docx",
        "pdf": f"{BASE_URL}/cv/GAIDICA_WashU_CV.pdf",
    }


def build_api_payload(data: dict) -> dict:
    meta = data.get("meta") or {}
    contact = _entries_as_dict(_section_by_title(data, "Contact Information"))
    present = _section_by_title(data, "Present Position")
    present_text = ""
    if present and present.get("entries"):
        present_text = str(present["entries"][0].get("text") or "")

    return {
        "api_version": "1",
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "source": "cv.yaml",
        "formats": public_formats(),
        "profile": {
            "name": meta.get("name"),
            "revised": meta.get("revised"),
            "email": contact.get("Email"),
            "phone": contact.get("Phone"),
            "office_address": contact.get("Office Address"),
            "present_position": present_text,
        },
        "cv": data,
    }


def build_api_index(data: dict) -> dict:
    meta = data.get("meta") or {}
    return {
        "api_version": "1",
        "name": "Matthew Gaidica CV API",
        "description": (
            "Static JSON API for the WashU Medicine CV. "
            "Fetch cv.json for the full structured record; use profile fields for quick facts."
        ),
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "cv_revised": meta.get("revised"),
        "endpoints": [
            {
                "path": "/cv/api/v1/index.json",
                "url": f"{BASE_URL}/cv/api/v1/index.json",
                "method": "GET",
                "description": "API discovery document",
            },
            {
                "path": "/cv/api/v1/cv.json",
                "url": f"{BASE_URL}/cv/api/v1/cv.json",
                "method": "GET",
                "description": "Full CV as JSON (profile + structured body from cv.yaml)",
            },
            {
                "path": "/cv/cv.yaml",
                "url": f"{BASE_URL}/cv/cv.yaml",
                "method": "GET",
                "description": "Editable YAML source of truth",
            },
            {
                "path": "/cv/GAIDICA_WashU_CV.docx",
                "url": f"{BASE_URL}/cv/GAIDICA_WashU_CV.docx",
                "method": "GET",
                "description": "WashU Medicine Word CV",
            },
            {
                "path": "/cv/GAIDICA_WashU_CV.pdf",
                "url": f"{BASE_URL}/cv/GAIDICA_WashU_CV.pdf",
                "method": "GET",
                "description": "PDF export",
            },
        ],
        "agent_notes": [
            "Prefer cv.json for structured parsing.",
            "Dates use MM/YYYY; list order is chronological (oldest → newest) within sections.",
            "body[] blocks are section | subsection | bibliography | note | entries.",
            "No auth required. Cache using generated_at or profile.revised.",
        ],
    }


def build_api(data: dict) -> tuple[Path, Path]:
    API_DIR.mkdir(parents=True, exist_ok=True)
    payload = build_api_payload(data)
    index = build_api_index(data)
    API_CV_OUT.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    API_INDEX_OUT.write_text(json.dumps(index, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    return API_CV_OUT, API_INDEX_OUT


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
    parser.add_argument("--api-only", action="store_true")
    parser.add_argument("--source", type=Path, default=SOURCE)
    args = parser.parse_args(argv)

    if not args.source.exists():
        print(f"Missing source: {args.source}", file=sys.stderr)
        return 1

    exclusive = sum(bool(x) for x in (args.docx_only, args.pdf_only, args.api_only))
    if exclusive > 1:
        print("Use only one of --docx-only, --pdf-only, --api-only", file=sys.stderr)
        return 1

    data = load_cv(args.source)
    # Always emit API JSON (cheap, agent-facing). Flags limit docx/pdf only.
    if args.api_only:
        do_api, do_docx, do_pdf = True, False, False
    elif args.docx_only:
        do_api, do_docx, do_pdf = True, True, False
    elif args.pdf_only:
        do_api, do_docx, do_pdf = True, False, True
    else:
        do_api, do_docx, do_pdf = True, True, True

    if do_api:
        cv_path, index_path = build_api(data)
        print(f"Wrote {cv_path.relative_to(ROOT)}")
        print(f"Wrote {index_path.relative_to(ROOT)}")

    md_path = None
    if do_docx or do_pdf:
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
