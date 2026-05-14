# Patient Template-Driven Report Design

**Goal**

Convert `GKTJ-Pt-skill` from direct lightweight report rendering into a template-driven patient report workflow that fills a fixed `.docx` template and updates Word native editable charts.

**Scope**

- In scope: unified patient template only
- In scope: placeholder rules, chart slot rules, payload v2 contract
- In scope: future renderer split into template fill and chart update stages
- Out of scope: doctor template, multi-template routing, final code implementation in this pass

## Current State

The existing pipeline is:

1. `parse_questionnaire.py`
2. `build_payload.py`
3. `render_report.py`

This flow is built around a lightweight four-chapter document and explicit Word formatting in code. It does not bind to a customer template, and it renders charts through `matplotlib` PNG files inserted by `python-docx`.

## Target Architecture

The new workflow is:

1. `parse_questionnaire.py` produces normalized questionnaire data.
2. AI or rule-based drafting produces structured content.
3. `build_payload.py` emits payload v2 with template-aware fields and chart data.
4. `render_from_template.py` fills a fixed patient template.
5. `update_word_charts.py` updates pre-seeded Word native charts.
6. Validation checks confirm slot alignment and Word editability.

This split matters because content insertion and chart updating are different responsibilities. The old renderer mixed formatting, content structure, and chart generation in one place.

## Design Decisions

### 1. Keep one unified patient template first

Do not split into three patient template variants yet. The first version should optimize for stable production output and a narrow feedback loop.

### 2. Use the current skill structure as the primary backbone

The current patient report logic remains the semantic source of truth:

- `引言`
- `数据信息分析`
- `反馈意见分析`
- `附件`

The customer template influences layout and optional section framing, but does not replace the content model wholesale.

### 3. Make placeholders semantic, not visual

Placeholder naming must reflect meaning such as `chapter2.item_01.analysis`, not styling concepts such as `green_title_1` or physical coordinates.

### 4. Treat charts as named slots

Charts are not generated ad hoc. The template owns chart positions and style. The payload owns chart data. The update script owns the binding between the two.

## File-Level Plan

- `references/template-spec.md`
  - Source of truth for template structure, placeholders, chart slots, payload v2
- `references/report-payload-schema.md`
  - Updated workflow and schema guidance
- `SKILL.md`
  - Updated high-level operating model
- `README.md`
  - Updated repository overview and future script responsibilities

Future implementation files, not created in this pass:

- `templates/patient-unified-v1.docx`
- `scripts/render_from_template.py`
- `scripts/update_word_charts.py`

## Risks And Constraints

### Word chart editability risk

The largest technical risk is not content generation. It is whether the chart update path can reliably mutate the embedded chart workbook and preserve editability in Microsoft Word.

### Template drift risk

If the fixed template changes without a matching slot contract, rendering will silently degrade. This is why slot naming must be explicit and validated.

### Attachment expansion risk

The attachment section is potentially long and structurally repetitive. The first implementation should allow controlled duplication logic rather than requiring a fully static placeholder grid.

## Acceptance Criteria For The Next Implementation Phase

- A single patient `.docx` template exists under `templates/`.
- Placeholder naming follows `references/template-spec.md`.
- Payload v2 can represent chapter text, feedback items, attachment items, and chart data.
- `render_from_template.py` can fill at least one end-to-end sample report.
- `update_word_charts.py` can update 3 to 5 template charts and keep them editable in Word.
- Final output visually stays close to the customer patient template.
