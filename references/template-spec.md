# Template-Driven Patient Report Spec

## Goal

Upgrade the patient report skill from direct Word composition into a template-driven pipeline:

1. Extract a single unified patient report structure.
2. Build a structured payload with content variables and chart data.
3. Render content into a fixed `.docx` template.
4. Update pre-seeded Word native editable charts inside that template.

This first version targets only the unified patient template. Doctor reports are out of scope.

## Unified Template Structure

The template keeps the current skill's report logic as the primary backbone, then expands it toward the shared structure found in the Guokai Tianjin patient cases.

### Required chapters

1. `引言`
2. `2、数据信息分析`
3. `3、反馈意见分析`
4. `3.1 积极反馈`
5. `3.2 待改进反馈`
6. `4、附件-问卷题目内容`

### Allowed substructure additions

The unified patient template may include these optional blocks when they are present in the fixed customer template:

- `报告背景`
- `数据来源`
- `综合分析与建议`

These blocks are template-level sections, not excuses to change the payload shape ad hoc. They must be represented by explicit placeholders.

## Template Placeholder Rules

Use double-brace placeholders in body text:

```text
{{meta.product}}
{{meta.region}}
{{meta.time}}
{{intro.p1}}
{{chapter2.item_01.title}}
{{chapter2.item_01.analysis}}
{{feedback.positive.item_01.title}}
{{feedback.positive.item_01.body}}
```

### Naming rules

- Prefixes are fixed: `meta`, `intro`, `chapter2`, `feedback`, `attachment`, `summary`, `chart`.
- Use lowercase ASCII keys only.
- Use `_01`, `_02` style zero-padded indexes for repeated blocks.
- Placeholder names describe semantic content, not formatting intent.
- A placeholder must map to exactly one payload field.
- Do not encode punctuation inside the placeholder name.

### Placeholder groups

#### Meta

- `{{meta.report_title}}`
- `{{meta.product}}`
- `{{meta.region}}`
- `{{meta.time}}`
- `{{meta.report_type}}`
- `{{meta.sample_size}}`

#### Introduction

- `{{intro.p1}}`
- `{{intro.p2}}`
- `{{intro.p3}}`
- `{{intro.p4}}`

#### Chapter 2 repeated content

- `{{chapter2.item_01.heading}}`
- `{{chapter2.item_01.title}}`
- `{{chapter2.item_01.analysis}}`
- `{{chart.slot_01.caption}}`

`heading` is the display string as shown in the template, for example `· 用药依从性`.

#### Feedback repeated content

- `{{feedback.positive.item_01.title}}`
- `{{feedback.positive.item_01.body}}`
- `{{feedback.negative.item_01.title}}`
- `{{feedback.negative.item_01.body}}`

#### Attachment repeated content

- `{{attachment.item_01.question}}`
- `{{attachment.item_01.option_a}}`
- `{{attachment.item_01.option_a_pct}}`

Attachment placeholders exist for template planning, but the first implementation may render the attachment block through controlled paragraph/table duplication instead of fully enumerating static placeholders.

## Chart Slot Rules

Charts must be Word native editable charts embedded in the template before rendering. Script code updates chart workbook data and title/caption bindings. PNG insertion is not allowed.

### Slot identifiers

- `chart_slot_01`
- `chart_slot_02`
- `chart_slot_03`

### Slot contract

Each chart slot corresponds to exactly one `chapter2` item and contains:

- a chart object already present in the `.docx` template
- a nearby caption placeholder such as `{{chart.slot_01.caption}}`
- stable binding metadata recorded in the payload and chart update layer

### Matching rules

- `chart_slot_01` maps to `chapter2.items[0]`
- slot order follows chapter 2 order
- one question maps to one chart slot
- first version supports only single-series horizontal bar style matching the customer template

## Payload V2 Structure

The new payload is structured around template filling and chart updates rather than direct renderer formatting.

```json
{
  "meta": {
    "product": "厄贝沙坦氢氯噻嗪片",
    "region": "新疆维吾尔自治区",
    "time": "2026.05",
    "report_type": "患者端问卷调研分析报告",
    "report_title": "问卷调研分析报告-患者",
    "sample_size": "N=120"
  },
  "template": {
    "template_id": "patient-unified-v1",
    "template_file": "templates/patient-unified-v1.docx",
    "version": "v1"
  },
  "intro": {
    "paragraphs": ["...", "..."]
  },
  "chapter2": {
    "items": [
      {
        "slot": "item_01",
        "heading": "· 用药依从性",
        "title": "用药依从性",
        "analysis": "正文",
        "question_ref": "q01",
        "chart_ref": "chart_slot_01",
        "chart": {
          "type": "bar_horizontal",
          "title": "图1 用药依从性",
          "categories": ["A. ...", "B. ..."],
          "series": [
            {
              "name": "占比",
              "values": [48.98, 32.65]
            }
          ]
        }
      }
    ]
  },
  "feedback": {
    "positive": [
      {"slot": "item_01", "title": "用药体验整体较好", "body": "正文"}
    ],
    "negative": [
      {"slot": "item_01", "title": "依从性管理仍需加强", "body": "正文"}
    ]
  },
  "attachment": {
    "questions": [
      {
        "question_ref": "q01",
        "number": 1,
        "question": "原始题目",
        "options": [
          {"code": "A", "text": "选项内容", "pct": "48.98%"}
        ]
      }
    ]
  },
  "checks": {
    "data_issue": null,
    "unresolved": null
  }
}
```

## Rendering Responsibilities

### `build_payload.py`

- Keep questionnaire normalization as the upstream source of truth.
- Upgrade output from flat report sections to payload v2.
- Add explicit `template`, `chart_ref`, `question_ref`, and repeated block slot metadata.

### `render_from_template.py`

- Load fixed `.docx` template.
- Replace scalar placeholders.
- Expand repeated content blocks for chapter 2, feedback, and attachment.
- Preserve original Word styling from the template instead of re-creating typography in code.

### `update_word_charts.py`

- Locate pre-seeded chart objects in the rendered `.docx`.
- Update chart embedded workbook data using payload chart definitions.
- Update chart titles and validate slot-to-question alignment.

## First-Version Constraints

- Only one unified patient template.
- No doctor template work.
- No multi-template dispatch logic.
- No fallback to PNG charts.
- No attempt to auto-infer customer template structure from arbitrary `.docx` files at runtime.

## Validation Checklist

- Chapter 2 item count equals chart slot count.
- Every chart slot in payload exists in the template.
- Every repeated content block has deterministic ordering.
- Final `.docx` opens in Word without repair prompts.
- Charts remain editable in Word after output generation.
