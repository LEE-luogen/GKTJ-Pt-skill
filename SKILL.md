---
name: "GKTJ-Pt-skill"
description: "Use when generating patient-facing questionnaire analysis reports from uploaded survey spreadsheets or questionnaire tables, especially when the output must include fixed sections, native Word charts, controlled typography, and restrained patient-facing wording. Uses customer DOCX template as the output baseline."
---

# GKTJ Patient Survey Report Generator

## Overview
This skill generates patient-facing questionnaire analysis reports using the customer DOCX template as the output baseline. The template provides a cover page, 11 native Word charts with embedded Excel workbooks, page numbering, and professional formatting. The renderer replaces text content while preserving the template's chart structure, colors, and layout.

Primary and secondary heading bullets/numbers are owned by the template's Word numbering (`w:numPr`). Heading text must never include literal prefixes such as `一、` or `·`, otherwise Word displays duplicated numbering.

## When to Use
- Uploaded attachment contains patient questionnaire data, especially `.xlsx`, `.csv`, or copied survey tables.
- Output must follow a fixed patient report structure rooted in:
  `引言` / `数据信息分析` / `反馈意见分析` / `综合分析与建议` / `附件`.
- Every question in chapter 2 needs a matching native Word chart.
- The report must stay in patient viewpoint throughout.
- The user wants `.docx` artifact with editable charts, not PNG images.

Do not use this skill for doctor reports, clinical trial manuscripts, or unstructured brainstorming.

## Template Facts
- Template file: `templates/patient-unified-v1.docx`
- Manifest: `templates/patient-unified-v1.manifest.json`
- Source: Customer DOCX from 爱诺-国开天津
- 1 cover page
- 2 sections (cover + content)
- 11 native Word charts (horizontal bar, multi-series per option)
- 11 embedded Excel workbooks
- Footer with PAGE field for page numbering
- Max 11 questions per report
- Chart→Workbook mapping is NOT sequential (must resolve via OOXML relationships)
- Questionnaire input may use either `label`/`pct` or `code`/`percentage`; normalize both before payload construction.
- If a question has more options than the pre-seeded chart series, clone the last template series and update its formulas, cache, and workbook column rather than dropping the option.

## Workflow
1. Normalize inputs:
   - Required: `品种`, `地区`, questionnaire attachment.
   - Optional: `时间`, `问卷收集时间`, `发放问卷数`, `有效问卷数/样本数`, custom execution note, custom output directory.
2. Parse the questionnaire data first.
   - Use `scripts/parse_questionnaire.py` for spreadsheet inputs.
   - Save structured output as `questionnaire.json` before writing report sections.
   - Validate: max 11 questions, each question has options with codes and counts.
3. Derive one statement-style title per question.
   - These titles are the chapter 2 dimension headings.
   - Each heading must be a concise summary of the question meaning, preferably 4-7 Chinese characters, not exceeding 9 characters.
   - Do not copy the full question stem or generate a long sentence-like heading.
4. Generate report content in this order:
   - `引言` (报告背景 + 数据来源)
   - `2、数据信息分析` question by question
   - `3.1 积极反馈`
   - `3.2 待改进反馈`
   - `4、综合分析与建议`
   - `5、附件-问卷题目内容`
   - When drafting prose, use the patient expression modules in `references/expression-modules.md`.
   - In `引言`, default to two subheads: `报告背景` and `数据来源`.
   - If the user provides questionnaire collection time and counts, make sure they are reflected in the `数据来源` paragraph instead of being omitted.
5. Build a report payload JSON through a script, not by hand.
   - AI should first write `report_content.md` or `report_content.jsonl`.
   - Use `scripts/build_payload.py --schema-version v2` to convert that draft plus `questionnaire.json` into a valid `report_payload_v2.json`.
   - Validation checks:
     - 11-question maximum
     - Intro word count: background 300-350 chars, data source 150-200 chars
     - Chapter 2 analysis: 300-400 chars per question
     - Percentage values: 0-100 range, single-choice sum within 0.05 of 100%
     - Chart slot binding: chart_ref follows pattern chart_slot_NN
     - Body text percentage references must exist in question options
6. Render artifacts.
   - Use `scripts/render_from_template.py` with payload v2 JSON.
   - The renderer:
     1. Copies the customer DOCX template.
     2. Replaces cover text (product, disease, region, time, company).
     3. Replaces chapter 1 (报告背景 + 数据来源).
     4. Replaces chapter 2 headings and analysis text (charts preserved).
     5. Replaces chapters 3-5 text.
     6. Inserts TOC field after cover section.
     7. Sets updateFields=true in settings.xml.
     8. Calls `update_word_charts.py` to update all 11 native charts and workbooks.
     9. Validates output.
   - Charts are native Word charts (not PNG images) and are editable in Word.
   - Number format: 0.00% in both chart labels and Excel workbooks.
   - Chart font: Songti, exactly 10pt for labels, legends, and axes.
7. Verify before delivery.
   - Chart count must equal chapter 2 question count (max 11).
   - Attachment must preserve original question and option meaning, but must not show percentages.
   - No absolute efficacy or safety claims.
   - TOC field exists and updateFields=true.
   - Output DOCX opens without repair prompt.

## Required Output Rules
- Patient viewpoint only. Do not rewrite as doctor judgement.
- Do not invent sample size, hospitals, institutions, or dates.
- If sample size or time is missing, use restrained wording and omit unsupported facts.
- In chapter 2, analyze percentage relationships; do not mechanically enumerate A/B/C/D and stop there.
- In chapter 2, each question should preferentially use one of these 3 analysis angles:
  - analyze option percentages and give a conclusion plus feasible suggestion
  - analyze the percentage of major options one by one and explain possible causes
  - analyze the overall percentage structure and summarize the group-level pattern
- The model may randomly choose any one of the 3 angles question by question, but the chosen angle must still fit the actual data distribution.
- Adjacent chapter 2 items must not reuse the same opening structure.
- Each chapter 2 item must be controlled within 300-400 Chinese characters.
- In chapter 2, always decide the pattern first, then write:
  - overall recognition
  - conditional recognition
  - behavioural split
  - pain-point attention
- Translate percentage structure into:
  patient experience, self-management behaviour, convenience, understanding, support needs.
- Do not turn patient questionnaire feedback into efficacy proof or clinical recommendation.

## Percentage Rules
- Internal values: 0-100 range (e.g., 25.99)
- Excel workbook: 0-1 range (e.g., 0.2599), format "0.00%"
- Chart labels: display as "25.99%"
- Body text: display as "25.99%"
- No integer percentages, no smart trailing-zero removal
- Single-choice questions: sum must be within 0.05 of 100%
- Multi-choice questions: sum may exceed 100%

## Section Rules
- Read `references/section-rules.md` before generating text.
- Read `references/compliance-rules.md` before finalizing text.
- Read `references/execution-rules.md` before building the report payload.
- Read `references/expression-modules.md` before drafting narrative paragraphs.

## Scripts
- `scripts/parse_questionnaire.py`
  - Reads questionnaire spreadsheets and emits normalized JSON.
- `scripts/build_payload.py`
  - Reads `questionnaire.json` plus structured report content and emits a validated `report_payload_v2.json`.
  - Supports `--schema-version v2` for the new payload format.
  - Validates: 11-question limit, intro word counts, percentages, chart slot binding.
- `scripts/render_from_template.py`
  - Main renderer: copies customer DOCX template, replaces text, inserts TOC, updates charts.
  - Uses `scripts/update_word_charts.py` for native chart updates.
  - Produces `.docx` with cover page, TOC, native charts, page numbers.
- `scripts/update_word_charts.py`
  - Updates native Word charts and embedded Excel workbooks.
  - Resolves chart→workbook mapping via OOXML relationships.
  - Preserves multi-series bar chart structure (each option = one series).
  - Sets chart text to Songti 10pt and number format to 0.00%.
  - Outputs detailed slot validation results.

## Expected File Flow
- Input:
  - attachment spreadsheet
- Intermediate:
  - `questionnaire.json` (normalized questionnaire data)
  - `report_content.md` or `report_content.jsonl` (AI-generated report draft)
  - `report_payload_v2.json` (validated payload)
- Output:
  - `问卷调研分析报告-{product}-患者端-{region}.docx` (final report with cover, TOC, native charts)

## Exceptions and Guards
- If the questionnaire has more than 11 questions, terminate with error.
- If a chart slot has no matching workbook relationship, terminate with error.
- If percentages cannot be parsed, terminate with error.
- If output DOCX cannot be opened, terminate with error.
- If background text is under 300 characters, terminate with error.
- If data source text is under 150 characters, terminate with error.
- If chart text is not Songti 10pt after update, terminate with an error.
- If TOC cannot be refreshed, preserve the TOC field and mark in validation.

## References
- `references/section-rules.md` — chapter structure and heading rules
- `references/expression-modules.md` — patient expression patterns
- `references/compliance-rules.md` — wording compliance rules
- `references/execution-rules.md` — execution rules
- `templates/patient-unified-v1.docx` — customer DOCX template
- `templates/patient-unified-v1.manifest.json` — template structure manifest
