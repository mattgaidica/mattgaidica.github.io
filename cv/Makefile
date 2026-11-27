PYTHON ?= .venv/bin/python
PIP ?= .venv/bin/pip

.PHONY: all setup api docx pdf clean

all: setup-check
	$(PYTHON) build.py

setup:
	python3 -m venv .venv
	$(PIP) install -r requirements.txt

api: setup-check
	$(PYTHON) build.py --api-only

docx: setup-check
	$(PYTHON) build.py --docx-only

pdf: setup-check
	$(PYTHON) build.py --pdf-only

setup-check:
	@test -x $(PYTHON) || (echo "Run: make setup" && exit 1)

clean:
	rm -rf _build
	rm -f GAIDICA_WashU_CV.pdf
	rm -f api/v1/cv.json api/v1/index.json
