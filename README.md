# GKTJ-Pt-skill

A reusable skill for generating patient-facing questionnaire analysis reports with:

- fixed chapter structure
- patient-viewpoint writing rules
- restrained medical/pharmaceutical wording
- script-built payload JSON from structured intermediate drafts
- body-only `.docx` rendering without a cover page
- PNG chart images inserted into the report, one chart per question

## Structure

- `SKILL.md`
  - Top-level workflow and triggering rules
- `agents/`
  - UI metadata
- `references/`
  - Section rules, compliance rules, execution rules, expression modules, payload schema
- `scripts/`
  - Questionnaire parsing, payload building, and report rendering
- `assets/`
  - Reserved for future static templates or visual assets; current default renderer does not use a cover page
- `templates/`
  - fixed patient `.docx` template assets kept for compatibility experiments
- `docs/superpowers/specs/`
  - design docs for larger migrations

## Typical Inputs

- Product name
- Region
- Optional time
- Optional survey period / issued count / valid count / sample size
- Patient questionnaire spreadsheet or normalized questionnaire table

## Typical Outputs

- `questionnaire.json`
- `report_content.md`
- `report_payload.json` or `report_payload.v2.json`
- `report_draft.md`
- `report_final.md`
- chart images for the legacy renderer
- `问卷调研分析报告-{{品种}}-患者端-{{地区}}.docx`
- `report_summary.json`

## Scripts

### Parse questionnaire

```bash
python3 scripts/parse_questionnaire.py input.xlsx -o output/questionnaire.json
```

### Build payload

```bash
python3 scripts/build_payload.py output/questionnaire.json output/report_content.md -o output/report_payload.json --product "厄贝沙坦氢氯噻嗪片" --region "河南" --survey-period "2026年5月1日至2026年5月31日" --issued-count 2250 --valid-count 2205
```

### Render report

```bash
python3 scripts/render_report.py output/report_payload.json --output-dir output/
```

### Build payload v2

```bash
python3 scripts/build_payload.py output/questionnaire.json output/report_content.md -o output/report_payload.v2.json --schema-version v2
```

### Render from template

```bash
python3 scripts/render_from_template.py output/report_payload.v2.json -o output/report_rendered.docx
```

## Doctor vs Patient

- Reused from doctor skill:
  - parse / payload / render engineering chain
  - chart rendering
  - docx typography
- Rewritten for patient skill:
  - `SKILL.md`
  - patient section rules
  - patient compliance rules
  - patient expression modules

## Notes

- This repository is optimized for patient-facing questionnaire reports, not doctor reports.
- The legacy workflow is `questionnaire.json -> report_content.md -> report_payload.json -> final artifacts`.
- The payload v2 workflow is `questionnaire.json -> report_content.md -> report_payload.v2.json -> render_from_template.py`.
- The current patient default output starts from正文首页 and does not generate a cover page.
- Every chapter 2 item must produce one PNG chart image and one matching inserted chart in the final `.docx`.
- The introduction should render `报告背景` and `数据来源` subheadings; if survey period / counts are provided, they will be injected into the data-source paragraph.
- `update_word_charts.py` is kept only as a legacy experiment path and is not part of the default patient workflow.

## Reusing This Pattern

When creating future report skills:
1. Define the audience first.
2. Fix the chapter structure second.
3. Lock compliance boundaries third.
4. Reuse scripts last.
