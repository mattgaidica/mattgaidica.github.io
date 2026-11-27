# WashU CV pipeline

Edit **`cv.yaml`** (source of truth), then rebuild DOCX/PDF.

WashU Medicine expects a Word CV for faculty packets ([Faculty Packets](https://faculty.med.wustl.edu/appointments-promotions/faculty-packets/)): Times New Roman 12pt or Arial 11pt, dates as `MM/YYYY` when known (year-only `YYYY` is allowed), chronological oldest→newest. This pipeline mirrors the section layout already in `GAIDICA_WashU_CV.docx`.

## Quick start

```bash
cd cv
make setup    # once
make all      # → API JSON + GAIDICA_WashU_CV.docx + GAIDICA_WashU_CV.pdf
```

Requirements beyond Python packages: [pandoc](https://pandoc.org/) and a LaTeX engine (`xelatex`, `tectonic`, or `pdflatex`) for PDF.

```bash
# DOCX only (no LaTeX needed; still writes API JSON)
python build.py --docx-only

# API JSON only
python build.py --api-only
```

## Public API (static)

GitHub Pages cannot run a live server, so the “API” is versioned static JSON rebuilt with the CV:

| Endpoint | Description |
|----------|-------------|
| https://gaidi.ca/cv/api/v1/index.json | Discovery + agent notes |
| https://gaidi.ca/cv/api/v1/cv.json | Full structured CV |
| https://gaidi.ca/cv/cv.yaml | YAML source |
| https://gaidi.ca/cv/GAIDICA_WashU_CV.docx | WashU Word CV |
| https://gaidi.ca/cv/GAIDICA_WashU_CV.pdf | PDF |

```bash
curl -s https://gaidi.ca/cv/api/v1/cv.json | jq '.profile'
```

Human docs: https://gaidi.ca/cv/api/v1/

## Layout of `cv.yaml`

```yaml
meta:
  name: Matthew Gaidica, PhD
  revised: 08/2026          # bump when you update
  footer_name: Gaidica, Matthew
  font: Times New Roman
  font_size_pt: 10.5

body:
  - type: section            # bold + underline
    title: Honors and Awards
    entries:
      - dates: 2025
        text: Explorers Club, Fellow

  - type: subsection         # bold only
    title: Courses
    entries: [...]

  - type: bibliography
    title: Original, peer reviewed articles in refereed journals
    items:
      - "Ellens DJ, Gaidica M, ..."
```

Entry `text` may be multiline (`|`). Keep dates oldest→newest within each section.

## Bot / agent workflow

1. Edit `cv.yaml` only (add grants, pubs, awards, bump `meta.revised`).
2. Commit and push — GitHub Actions rebuilds `GAIDICA_WashU_CV.docx` / `.pdf` and commits them back.
3. Optionally run `make all` locally first if you want to preview before pushing.

Do not hand-edit the DOCX; it will be overwritten on the next build. A one-time snapshot of the prior Word file is in `archive/`.

## GitHub Actions

[`.github/workflows/cv.yml`](../.github/workflows/cv.yml) runs on pushes to `master`/`main` that touch:

- `cv/cv.yaml`
- `cv/build.py`
- `cv/requirements.txt`
- the workflow file itself

It also supports **Run workflow** (manual) from the Actions tab. Output-only commits are ignored by the path filter, so the bot commit does not loop.

CI uses `Liberation Serif` for PDF (`CV_MAINFONT`); local builds default to Times New Roman when available.

## Outputs

| File | Role |
|------|------|
| `cv.yaml` | Editable source |
| `api/v1/cv.json` | Machine-readable API payload |
| `api/v1/index.json` | API discovery |
| `GAIDICA_WashU_CV.docx` | WashU submission / packet |
| `GAIDICA_WashU_CV.pdf` | Shareable / site copy |
| `_build/cv.md` | Intermediate markdown (gitignored) |
