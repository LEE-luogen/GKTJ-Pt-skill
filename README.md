# GKTJ-Pt-skill

A reusable skill for generating patient-facing questionnaire analysis reports with:

- fixed chapter structure
- patient-viewpoint writing rules
- restrained medical/pharmaceutical wording
- script-built payload JSON from structured intermediate drafts
- chart generation
- explicit Word typography rendering

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

## Typical Inputs

- Product name
- Region
- Optional time
- Patient questionnaire spreadsheet or normalized questionnaire table

## Typical Outputs

- `questionnaire.json`
- `report_content.md`
- `report_payload.json`
- `report_draft.md`
- `report_final.md`
- chart images
- `问卷调研分析报告-{{品种}}-患者端-{{地区}}.docx`
- `report_summary.json`

## Scripts

### Parse questionnaire

```bash
python3 scripts/parse_questionnaire.py input.xlsx -o output/questionnaire.json
```

### Build payload

```bash
python3 scripts/build_payload.py output/questionnaire.json output/report_content.md -o output/report_payload.json --product "厄贝沙坦氢氯噻嗪片" --region "河南"
```

### Render report

```bash
python3 scripts/render_report.py output/report_payload.json --output-dir output/
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
- The default workflow is `questionnaire.json -> report_content.md -> report_payload.json -> final artifacts`.
- The default `.docx` output starts from the正文首页 and does not generate a cover page.
- Core typography is explicit rather than style-name-driven: 宋体 20pt/16pt/14pt for heading hierarchy and正文.

## Reusing This Pattern

When creating future report skills:
1. Define the audience first.
2. Fix the chapter structure second.
3. Lock compliance boundaries third.
4. Reuse scripts last.
