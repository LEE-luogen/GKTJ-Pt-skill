# GKTJ-Pt-skill

A reusable skill for generating patient-facing questionnaire analysis reports with:

- fixed chapter structure
- patient-viewpoint writing rules
- restrained medical/pharmaceutical wording
- script-built payload JSON from structured intermediate drafts
- a migration path toward template-driven `.docx` rendering
- Word native editable chart support as the target chart mode

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
  - fixed patient `.docx` template and manifest for the template-driven flow
- `docs/superpowers/specs/`
  - design docs for larger migrations

## Typical Inputs

- Product name
- Region
- Optional time
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
python3 scripts/build_payload.py output/questionnaire.json output/report_content.md -o output/report_payload.json --product "厄贝沙坦氢氯噻嗪片" --region "河南"
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

### Update Word native charts

```bash
python3 scripts/update_word_charts.py output/report_rendered.docx output/report_payload.v2.json
```

## Template-Driven Direction

The repository is being upgraded from direct Word composition to a template-driven flow:

1. define one unified patient template
2. emit payload v2 with explicit content variables and chart data
3. fill a fixed `.docx` template
4. update pre-seeded Word native editable charts

See [references/template-spec.md](references/template-spec.md) and [docs/superpowers/specs/2026-05-14-patient-template-design.md](docs/superpowers/specs/2026-05-14-patient-template-design.md).

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
- The template-driven workflow is `questionnaire.json -> report_content.md -> report_payload.v2.json -> render_from_template.py -> update_word_charts.py`.
- The target workflow is template-driven and should stop expanding `render_report.py`.
- The legacy `.docx` output starts from正文首页 and does not generate a cover page.
- The first template-driven version is intentionally fixed to one committed patient template and one manifest.

## Reusing This Pattern

When creating future report skills:
1. Define the audience first.
2. Fix the chapter structure second.
3. Lock compliance boundaries third.
4. Reuse scripts last.
