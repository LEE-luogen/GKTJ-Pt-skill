# Patient Template Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the first working template-driven patient report pipeline with a fixed `.docx` template and editable Word chart updates.

**Architecture:** Reuse one real patient sample `.docx` as the first unified template, introduce payload v2 in `build_payload.py`, fill text content through a template manifest, then update embedded Word chart workbooks and chart XML directly inside the output `.docx`.

**Tech Stack:** Python 3, `python-docx`, `openpyxl`, `lxml`, `zipfile`, `unittest`

---

### Task 1: Add Template Assets And Manifest

**Files:**
- Create: `templates/patient-unified-v1.docx`
- Create: `templates/patient-unified-v1.manifest.json`

- [ ] Copy one patient sample report into the repo as the first fixed template.
- [ ] Record stable paragraph indexes and section anchors in a manifest.

### Task 2: Add Payload V2

**Files:**
- Modify: `scripts/build_payload.py`
- Test: `tests/test_template_pipeline.py`

- [ ] Add `--schema-version` handling.
- [ ] Build payload v2 with `meta`, `template`, `intro`, `chapter2`, `feedback`, `summary`, `attachment`, `checks`.
- [ ] Add validation for chart slot and question references.

### Task 3: Add Template Renderer

**Files:**
- Create: `scripts/render_from_template.py`
- Test: `tests/test_template_pipeline.py`

- [ ] Load template and manifest.
- [ ] Replace intro, chapter 2, feedback, summary, and attachment text.
- [ ] Save a rendered `.docx` without touching chart binaries yet.

### Task 4: Add Word Chart Updater

**Files:**
- Create: `scripts/update_word_charts.py`
- Test: `tests/test_template_pipeline.py`

- [ ] Update embedded workbook data for each chart slot.
- [ ] Update chart XML caches to match workbook values.
- [ ] Save the updated `.docx`.

### Task 5: Verify End-To-End

**Files:**
- Modify: `README.md`
- Test: `tests/test_template_pipeline.py`

- [ ] Run the template pipeline in tests against the committed template.
- [ ] Verify paragraphs changed and chart workbook values changed.
- [ ] Document the first-version limitations.
