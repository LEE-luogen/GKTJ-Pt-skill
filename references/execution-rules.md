# Execution Rules

## Step 1: Parse Questionnaire
- Use `scripts/parse_questionnaire.py` for spreadsheet inputs.
- Save structured output as `questionnaire.json`.
- Validate: max 11 questions, each question has options with codes and counts.
- If questionnaire has more than 11 questions, terminate with error.
- Normalize questionnaire option schemas: accept `label`/`pct` and `code`/`percentage`, strip duplicate option prefixes, and preserve original order.

## Step 2: Generate Report Content
- AI generates `report_content.md` or `report_content.jsonl`.
- Follow `references/section-rules.md` for structure.
- Follow `references/expression-modules.md` for patient expression patterns.
- Follow `references/compliance-rules.md` for wording rules.
- Word count targets:
  - 报告背景: 300-350 characters
  - 数据来源: 150-200 characters
  - Each chapter 2 analysis: 300-400 characters

## Step 3: Build Payload
- Use `scripts/build_payload.py --schema-version v2`.
- Input: `questionnaire.json` + `report_content.md` (or `.jsonl`).
- Output: `report_payload_v2.json`.
- CLI flags for overrides:
  - `--product`, `--region`, `--time`, `--report-type`
  - `--survey-period`, `--sample-size`
  - `--issued-count`, `--valid-count`
- Validation checks (errors block generation):
  - 11-question maximum
  - Intro word count: background ≥300, data source ≥150
  - Chapter 2 analysis: 300-400 chars per question
  - Percentage values: 0-100 range
  - Chart slot binding: chart_ref follows pattern chart_slot_NN
  - Required fields: meta, intro, chapter2, feedback, summary, attachment
- Validation checks (warnings, don't block):
  - Background >350 chars
  - Data source >200 chars
  - Single-choice percentage sum not exactly 100%
  - Body text percentage references not found in options

## Step 4: Render Report
- Use `scripts/render_from_template.py`.
- Input: `report_payload_v2.json` + template DOCX.
- Output: `.docx` file with cover, TOC, native charts, page numbers.
- Preserve template numbering: chapter and subsection text contains no literal Chinese number or bullet prefix; visible numbering comes exclusively from `w:numPr`.
- CLI: `python render_from_template.py payload.json -o output.docx`
- Template: `templates/patient-unified-v1.docx` (default, override with `--template`)
- Rendering steps:
  1. Copy customer DOCX template.
  2. Replace cover text (product, disease, region, time, company).
  3. Replace chapter 1 (报告背景 + 数据来源).
  4. Replace chapter 2 headings and analysis text (charts preserved).
  5. Replace chapters 3-5 text.
  6. Insert TOC field after cover section.
  7. Set updateFields=true in settings.xml.
  8. Call update_word_charts.py to update all 11 native charts.
  9. Validate output.

## Step 5: Verify Output
- Chart count must equal question count (max 11).
- TOC field exists in document.
- updateFields=true in settings.xml.
- Charts are native Word charts (not PNG images).
- When option count exceeds the chart's seeded series count, clone the final series and rewrite `idx`, `order`, workbook formulas, cache values, and workbook cells for the added option.
- Chart labels, legends, and axes use Songti at exactly 10pt.
- Number format 0.00% in workbooks.
- Cover page exists with correct text.
- Page numbers: cover has none, content has continuous numbering.
- Attachment: no percentages shown.
- No absolute efficacy or safety claims.
- Output DOCX opens without repair prompt.

## Exception Handling
- If questionnaire >11 questions → terminate.
- If chart slot missing workbook relationship → terminate.
- If percentages unparseable → terminate.
- If output DOCX cannot open → terminate.
- If background <300 chars → terminate.
- If data source <150 chars → terminate.
- If chart font is not Songti 10pt → terminate.
- If TOC cannot refresh → preserve field, mark in validation.
