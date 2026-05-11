# Execution Rules

## Standard Run
1. Parse attachment into `questionnaire.json`.
2. Derive question titles.
3. Draft patient report text as `report_content.md` or `report_content.jsonl`.
4. Build `report_payload.json` via `scripts/build_payload.py`.
5. Render markdown, charts, docx, summary.
6. Run final checks.

## Payload Construction Rules
- Default path: AI writes a structured intermediate draft, not a full JSON payload.
- Use `json.dumps(..., ensure_ascii=False)` through the script output path. Do not manually escape long Chinese paragraphs inside JSON strings.
- `build_payload.py` must:
  - attach normalized question options back into chapter 2 items
  - auto-fill `attachment_questions`
  - validate chapter 2 item count against `questionnaire.json`
  - apply default `report_type`
  - preserve optional `time`, `data_issue`, and `unresolved` notes when provided

## Emergency Fallback
- Direct JSON is allowed only as a backup path.
- It must be provided as a `.json` file or fenced ` ```json ` block and pass `json.loads` validation before rendering.
- This is not the default workflow.

## Required Checks
- chapter count correct
- chapter order correct
- chart count equals chapter 2 question count
- chart style uniform
- no charts in attachment section
- no unsupported data in prose
- no absolute efficacy or safety claims
- no cover page inserted by default
- explicit typography check:
  - 一级标题 20pt 绿色
  - 二级标题 16pt 黑色
  - 正文 14pt、28pt 首行缩进、1.5 倍行距

## Default Output Paths
- `04_outputs/questionnaire.json`
- `04_outputs/report_content.md`
- `04_outputs/report_payload.json`
- `04_outputs/report_draft.md`
- `04_outputs/report_final.md`
- `04_outputs/charts/`
- `04_outputs/问卷调研分析报告-{{品种}}-患者端-{{地区}}.docx`
- `04_outputs/report_summary.json`
