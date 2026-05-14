---
name: "GKTJ-Pt-skill"
description: "Use when generating patient-facing questionnaire analysis reports from uploaded survey spreadsheets or questionnaire tables, especially when the output must include fixed sections, consistent charts, controlled Word typography, and restrained patient-facing wording."
---

# GKTJ Patient Survey Report Generator

## Overview
This skill generates patient-facing questionnaire analysis reports with a fixed structure, controlled wording, and `.docx` output. The current default workflow is body-first: no cover page, one image chart per question, and fixed Word typography for patient reports.

## When to Use
- Uploaded attachment contains patient questionnaire data, especially `.xlsx`, `.csv`, or copied survey tables.
- Output must follow a fixed patient report structure rooted in:
  `引言` / `数据信息分析` / `反馈意见分析` / `附件`.
- Every question in chapter 2 needs a matching chart.
- The report must stay in patient viewpoint throughout.
- The user wants `.md` and `.docx` artifacts, not just chat output.

Do not use this skill for doctor reports, clinical trial manuscripts, or unstructured brainstorming.

## Workflow
1. Normalize inputs:
   - Required: `品种`, `地区`, questionnaire attachment.
   - Optional: `时间`, `问卷收集时间`, `发放问卷数`, `有效问卷数/样本数`, custom execution note, custom output directory.
2. Parse the questionnaire data first.
   - Use `scripts/parse_questionnaire.py` for spreadsheet inputs.
   - Save structured output as `questionnaire.json` before writing report sections.
3. Derive one statement-style title per question.
   - These titles are the chapter 2 dimension headings.
   - Each heading must be a concise summary of the question meaning, preferably 4-7 Chinese characters, not exceeding 9 characters.
   - Do not copy the full question stem or generate a long sentence-like heading.
4. Generate report content in this order:
   - `引言`
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
   - Use `scripts/build_payload.py` to convert that draft plus `questionnaire.json` into a valid `report_payload.json`.
6. Render artifacts.
   - Use `scripts/render_report.py` for direct payload v1 rendering, or `scripts/render_from_template.py` for payload v2 body-only rendering.
   - Both current report paths generate PNG chart images and insert them into the report.
   - Do not generate a cover page.
7. Verify before delivery.
   - Chart count must equal chapter 2 question count.
   - Attachment must preserve original question and option meaning, but must not show percentages.
   - No absolute efficacy or safety claims.

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
- Each chapter 2 item must be controlled within 300-350 Chinese characters.
- In chapter 2, always decide the pattern first, then write:
  - overall recognition
  - conditional recognition
  - behavioural split
  - pain-point attention
- Translate percentage structure into:
  patient experience, self-management behaviour, convenience, understanding, support needs.
- Do not turn patient questionnaire feedback into efficacy proof or clinical recommendation.

## Section Rules
- Read `references/section-rules.md` before generating text.
- Read `references/compliance-rules.md` before finalizing text.
- Read `references/execution-rules.md` before building the report payload.
- Read `references/expression-modules.md` before drafting narrative paragraphs.

## Scripts
- `scripts/parse_questionnaire.py`
  - Reads questionnaire spreadsheets and emits normalized JSON.
- `scripts/build_payload.py`
  - Reads `questionnaire.json` plus structured report content and emits a validated `report_payload.json`.
- `scripts/render_report.py`
  - Direct renderer for the current patient body-only flow.
  - Generates markdown, PNG charts, `.docx`, and a summary JSON.
- `scripts/render_from_template.py`
  - Renders payload v2 into the same body-only `.docx` structure without a cover page.
- `scripts/update_word_charts.py`
  - Legacy utility kept only for old editable-chart experiments; not part of the current patient default flow.

## Expected File Flow
- Input:
  - attachment spreadsheet
- Intermediate:
  - `questionnaire.json`
  - `report_content.md` or `report_content.jsonl`
  - `report_payload.json` or `report_payload.v2.json`
- Output:
  - `report_draft.md`
  - `report_final.md`
  - `charts/chart_XX.png` or `*_charts/chart_XX.png`
  - `report_rendered.docx`
  - `问卷调研分析报告-{{品种}}-患者端-{{地区}}.docx`
  - `report_summary.json`

## Common Mistakes
- Writing in doctor viewpoint.
- Letting chapter 3 repeat chapter 2 item by item.
- Asking the model to handwrite a long `report_payload.json` with many Chinese paragraphs.
- Turning patient feedback into “证明疗效” or “安全性确证”.
- Forgetting that every chapter 2 item needs one chart and only one chart.
- Showing percentages again inside the attachment section.
- Reintroducing a cover page or editable Word chart path for patient reports.

## Final Delivery
Reply with:
- markdown final path
- docx path
- chapter completeness check
- chart count check
- chart style consistency check
- missing or uncertain data
- unresolved issues
