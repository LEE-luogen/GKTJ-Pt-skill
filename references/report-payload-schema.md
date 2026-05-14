# Report Content And Payload Schema

## Status

This document now defines two states:

- `legacy v1`: current direct-render payload used by `render_report.py`
- `template-driven v2`: target payload for fixed `.docx` template filling and Word native chart updates

The migration target is `template-driven v2`. New implementation work should not expand `render_report.py` further.

## Default Workflow

Do not ask the model to handwrite a full `report_payload.json` as the primary path.

### Legacy v1 chain

1. `scripts/parse_questionnaire.py` -> `questionnaire.json`
2. AI writes `report_content.md` or `report_content.jsonl`
3. `scripts/build_payload.py` -> `report_payload.json`
4. `scripts/render_report.py` -> markdown, charts, docx, summary

### Target v2 chain

1. `scripts/parse_questionnaire.py` -> `questionnaire.json`
2. AI writes structured patient report content
3. `scripts/build_payload.py` -> `report_payload.v2.json`
4. `scripts/render_from_template.py` -> filled `.docx`
5. `scripts/update_word_charts.py` -> final `.docx` with editable Word native charts
6. validation -> summary

This avoids quote escaping failures when the content contains long Chinese paragraphs, Chinese quotation marks, English double quotes, or multiple line breaks.
The new renderer must preserve the fixed patient template instead of reconstructing layout in code.

## Recommended `report_content.md` Format

```md
---
product: 厄贝沙坦氢氯噻嗪片
region: 河南
time: 2025.10
report_type: 患者端问卷调研分析报告
data_issue:
unresolved:
---

## 1、引言

第一段。

第二段。

## 2、数据信息分析

### 1. 用药依从性

单题分析正文。

### 2. 血压控制体验

单题分析正文。

## 3、反馈意见分析

### 3.1 积极反馈

- 用药体验整体较好：正文。
- 便利性认可度较高：正文。

### 3.2 待改进反馈

- 依从性管理仍需加强：正文。
- 用药认知仍有误区：正文。
```

## Markdown Rules
- `引言` is paragraph-based.
- `数据信息分析` must contain one subheading per questionnaire item, in the same order as `questionnaire.json`.
- `积极反馈` and `待改进反馈` should use `标题：正文` blocks.
- `附件-问卷题目内容` should not be authored in the markdown intermediate draft. It is auto-filled from `questionnaire.json`.

## Optional `report_content.jsonl` Format

```jsonl
{"type":"meta","product":"厄贝沙坦氢氯噻嗪片","region":"河南","time":"2025.10"}
{"type":"introduction","text":"第一段。"}
{"type":"introduction","text":"第二段。"}
{"type":"data_analysis","number":1,"title":"用药依从性","analysis":"单题分析正文。"}
{"type":"positive_feedback","title":"用药体验整体较好","body":"正文。"}
{"type":"negative_feedback","title":"依从性管理仍需加强","body":"正文。"}
```

## Legacy v1 `report_payload.json`

```json
{
  "product": "厄贝沙坦氢氯噻嗪片",
  "region": "河南",
  "time": "2025.10",
  "report_type": "患者端问卷调研分析报告",
  "introduction": ["段落1", "段落2"],
  "data_analysis": [
    {
      "number": 1,
      "title": "用药依从性",
      "analysis": "正文",
      "question": "原始题目",
      "options": [
        {"label": "A", "text": "选项内容", "pct": "48.98%"}
      ]
    }
  ],
  "positive_feedback": [
    {"title": "标题", "body": "正文"}
  ],
  "negative_feedback": [
    {"title": "标题", "body": "正文"}
  ],
  "attachment_questions": [
    {
      "number": 1,
      "question": "原始题目",
      "options": [
        {"label": "A", "text": "选项内容", "pct": "48.98%"}
      ]
    }
  ],
  "checks": {
    "data_issue": null,
    "unresolved": null
  }
}
```

## Validation Rules
- `data_analysis` item count must equal `questionnaire.json.questions` count.
- `attachment_questions` is auto-generated from `questionnaire.json`.
- `report_type` defaults to `患者端问卷调研分析报告` if omitted.
- `time` may be omitted.
- If a direct JSON payload is used as a fallback path, it must pass `json.loads` validation before rendering.

## Template-Driven v2 Payload

Use the structure in [template-spec.md](./template-spec.md). Key differences from v1:

- top-level `meta`, `template`, `intro`, `chapter2`, `feedback`, `attachment`, `checks`
- each chapter 2 item carries `question_ref` and `chart_ref`
- chart data is stored as structured values, not image instructions
- attachment rendering is template-aware instead of always using the old hardcoded paragraph layout

## Migration Rules

- Do not add new customer-facing fields to legacy v1 unless needed for backward compatibility.
- All new template work should be modeled in v2 first.
- The old `render_report.py` remains a legacy renderer until template rendering is proven.
