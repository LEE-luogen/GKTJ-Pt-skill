from __future__ import annotations

import json
import shutil
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from openpyxl import load_workbook

from scripts.build_payload import build_payload_v2, parse_markdown_content, validate_payload_v2
from scripts.render_from_template import render_from_template
from scripts.update_word_charts import chart_targets, update_docx_charts, workbook_target


FIXTURE_DIR = Path(__file__).resolve().parents[1]
TEMPLATE_PATH = FIXTURE_DIR / "templates" / "patient-unified-v1.docx"
MANIFEST_PATH = FIXTURE_DIR / "templates" / "patient-unified-v1.manifest.json"


class Namespace:
    def __init__(self, **kwargs):
        self.__dict__.update(kwargs)


def sample_questionnaire() -> dict:
    questions = []
    for index in range(1, 11):
        questions.append(
            {
                "number": index,
                "question": f"第{index}题原始题目",
                "total": "120",
                "options": [
                    {"label": "A", "text": f"选项A{index}", "count": "60", "pct": "50.00%"},
                    {"label": "B", "text": f"选项B{index}", "count": "36", "pct": "30.00%"},
                    {"label": "C", "text": f"选项C{index}", "count": "24", "pct": "20.00%"},
                ],
            }
        )
    return {"questions": questions}


def sample_markdown() -> str:
    chapter2 = []
    for index in range(1, 11):
        chapter2.append(f"### {index}. 维度{index}\n\n维度{index}分析第一句。维度{index}分析第二句。")
    return """---
product: 测试产品
region: 测试地区
time: 2026.05
report_type: 患者端问卷调研分析报告
---

## 1、引言

第一段引言。

第二段引言。

## 2、数据信息分析

{chapter2}

## 3、反馈意见分析

### 3.1 积极反馈

- 优势一：积极反馈内容一。
- 优势二：积极反馈内容二。
- 优势三：积极反馈内容三。
- 优势四：积极反馈内容四。

### 3.2 待改进反馈

- 问题一：待改进内容一。
- 问题二：待改进内容二。
- 问题三：待改进内容三。
- 问题四：待改进内容四。
- 问题五：待改进内容五。

## 综合分析与建议

### 建议一

建议一正文。

### 建议二

建议二正文。

### 建议三

建议三正文。

### 建议四

建议四正文。

### 建议五

建议五正文。
""".format(chapter2="\n\n".join(chapter2))


class TemplatePipelineTest(unittest.TestCase):
    def test_build_payload_v2(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            report_content = Path(temp_dir) / "content.md"
            report_content.write_text(sample_markdown(), encoding="utf-8")
            meta, content = parse_markdown_content(report_content)
            payload = build_payload_v2(
                sample_questionnaire(),
                meta,
                content,
                Namespace(
                    product=None,
                    region=None,
                    time=None,
                    report_type=None,
                    data_issue=None,
                    unresolved=None,
                    template_id="patient-unified-v1",
                    template_file="templates/patient-unified-v1.docx",
                ),
            )
            validate_payload_v2(payload)
            self.assertEqual(payload["meta"]["product"], "测试产品")
            self.assertEqual(len(payload["chapter2"]["items"]), 10)
            self.assertEqual(payload["chapter2"]["items"][0]["chart_ref"], "chart_slot_01")
            self.assertEqual(payload["attachment"]["questions"][0]["question_ref"], "q01")

    def test_render_and_update_charts(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            output_docx = Path(temp_dir) / "rendered.docx"
            report_content = Path(temp_dir) / "content.md"
            report_content.write_text(sample_markdown(), encoding="utf-8")
            meta, content = parse_markdown_content(report_content)
            payload = build_payload_v2(
                sample_questionnaire(),
                meta,
                content,
                Namespace(
                    product=None,
                    region=None,
                    time=None,
                    report_type=None,
                    data_issue=None,
                    unresolved=None,
                    template_id="patient-unified-v1",
                    template_file="templates/patient-unified-v1.docx",
                ),
            )
            render_from_template(payload, TEMPLATE_PATH, MANIFEST_PATH, output_docx)
            update_docx_charts(output_docx, payload)

            document = Document(output_docx)
            texts = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]
            self.assertIn("第一段引言。", texts)
            self.assertIn("维度1", texts)
            self.assertIn("优势一：积极反馈内容一。", texts)
            self.assertIn("建议一", texts)
            self.assertIn("1.第1题原始题目", texts)

            with ZipFile(output_docx) as archive:
                first_chart = chart_targets(archive)[0]
                workbook_name = workbook_target(archive, first_chart)
                workbook = load_workbook(BytesIO(archive.read(workbook_name)))
                worksheet = workbook[workbook.sheetnames[0]]
                self.assertEqual(worksheet["B1"].value, "A.选项A1")
                self.assertAlmostEqual(float(worksheet["B2"].value), 0.5)


if __name__ == "__main__":
    unittest.main()
