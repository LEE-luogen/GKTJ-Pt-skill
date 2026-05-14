#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

DEFAULT_REPORT_TYPE = "患者端问卷调研分析报告"


def normalize_space(text: str) -> str:
    return re.sub(r"[ \t\u3000]+", " ", text.strip())


def normalize_pct(text: str) -> str:
    value = str(text).strip()
    if not value:
        return value
    if value.endswith("%"):
        number = value[:-1].strip()
    else:
        number = value
    try:
        return f"{float(number):.2f}%"
    except ValueError:
        return value


def clean_dimension_title(raw: str) -> str:
    text = normalize_space(raw)
    text = re.sub(r"^[·•]\s*", "", text)
    text = re.sub(r"^[（(]?\d+[）).、\s]+", "", text)
    return text.strip()


def question_options(question: dict) -> list[dict]:
    options = []
    for option in question.get("options", []):
        text = str(option.get("text") or "").strip()
        count = option.get("count")
        pct = normalize_pct(option.get("pct", ""))
        if text == option.get("label") and count in (None, "", "None") and pct in ("", "0%", "0.0%", "0.00%"):
            continue
        options.append(
            {
                "label": str(option.get("label", "")).strip(),
                "text": text,
                "pct": pct,
            }
        )
    return options


def split_front_matter(text: str) -> tuple[dict, str]:
    stripped = text.lstrip("\ufeff")
    match = re.match(r"^\s*---\s*\n(.*?)\n\s*---\s*\n?(.*)$", stripped, flags=re.DOTALL)
    if not match:
        return {}, stripped
    meta_block = match.group(1)
    body = match.group(2)
    meta = {}
    for line in meta_block.splitlines():
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        meta[key.strip()] = value.strip().strip('"').strip("'")
    return meta, body


def parse_paragraphs(lines: list[str]) -> list[str]:
    paragraphs: list[str] = []
    buf: list[str] = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            if buf:
                paragraphs.append(normalize_space(" ".join(buf)))
                buf = []
            continue
        buf.append(stripped)
    if buf:
        paragraphs.append(normalize_space(" ".join(buf)))
    return paragraphs


def split_blocks(lines: list[str]) -> list[str]:
    return parse_paragraphs(lines)


def parse_label_body_blocks(lines: list[str], section_name: str) -> list[dict]:
    items = []
    for raw in lines:
        stripped = raw.strip()
        if not stripped:
            continue
        stripped = re.sub(r"^[-*]\s*", "", stripped)
        if "：" in stripped:
            title, body = stripped.split("：", 1)
        elif ":" in stripped:
            title, body = stripped.split(":", 1)
        else:
            raise ValueError(f"{section_name} block missing colon separator: {stripped}")
        items.append({"title": normalize_space(title), "body": normalize_space(body)})
    return items


def parse_markdown_content(path: Path) -> tuple[dict, dict]:
    meta, body = split_front_matter(path.read_text(encoding="utf-8"))
    lines = body.splitlines()

    intro_lines: list[str] = []
    pos_lines: list[str] = []
    neg_lines: list[str] = []
    data_items: list[dict] = []
    summary_items: list[dict] = []

    current_section: str | None = None
    current_item: dict | None = None
    current_summary_item: dict | None = None

    def flush_item() -> None:
        nonlocal current_item
        if not current_item:
            return
        current_item["analysis"] = normalize_space(" ".join(current_item.pop("lines")))
        data_items.append(current_item)
        current_item = None

    def flush_summary_item() -> None:
        nonlocal current_summary_item
        if not current_summary_item:
            return
        current_summary_item["body"] = normalize_space(" ".join(current_summary_item.pop("lines")))
        summary_items.append(current_summary_item)
        current_summary_item = None

    for line in lines:
        heading = re.match(r"^(#{2,4})\s+(.+?)\s*$", line.lstrip())
        if heading:
            level = len(heading.group(1))
            title = normalize_space(heading.group(2))
            if "引言" in title:
                flush_item()
                flush_summary_item()
                current_section = "introduction"
                continue
            if "数据信息分析" in title:
                flush_item()
                flush_summary_item()
                current_section = "data_analysis"
                continue
            if "反馈意见分析" in title:
                flush_item()
                flush_summary_item()
                continue
            if "积极反馈" in title:
                flush_item()
                flush_summary_item()
                current_section = "positive_feedback"
                continue
            if "待改进反馈" in title:
                flush_item()
                flush_summary_item()
                current_section = "negative_feedback"
                continue
            if "综合分析与建议" in title:
                flush_item()
                flush_summary_item()
                current_section = "summary"
                continue
            if "附件" in title:
                flush_item()
                flush_summary_item()
                current_section = "attachment"
                continue
            if current_section == "data_analysis" and level >= 3:
                flush_item()
                current_item = {"title": clean_dimension_title(title), "lines": []}
                continue
            if current_section == "summary" and level >= 3:
                flush_summary_item()
                current_summary_item = {"title": clean_dimension_title(title), "lines": []}
                continue

        if current_section == "introduction":
            intro_lines.append(line)
        elif current_section == "data_analysis":
            if current_item:
                current_item["lines"].append(line)
        elif current_section == "positive_feedback":
            pos_lines.append(line)
        elif current_section == "negative_feedback":
            neg_lines.append(line)
        elif current_section == "summary":
            if current_summary_item:
                current_summary_item["lines"].append(line)

    flush_item()
    flush_summary_item()

    content = {
        "introduction": parse_paragraphs(intro_lines),
        "data_analysis": data_items,
        "positive_feedback": parse_label_body_blocks(pos_lines, "positive_feedback"),
        "negative_feedback": parse_label_body_blocks(neg_lines, "negative_feedback"),
        "summary_recommendations": summary_items,
    }
    return meta, content


def parse_jsonl_content(path: Path) -> tuple[dict, dict]:
    meta: dict = {}
    content = {
        "introduction": [],
        "data_analysis": [],
        "positive_feedback": [],
        "negative_feedback": [],
        "summary_recommendations": [],
    }
    for line_no, raw in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        line = raw.strip()
        if not line:
            continue
        item = json.loads(line)
        line_type = item.get("type") or item.get("section")
        if line_type == "meta":
            meta.update({k: v for k, v in item.items() if k not in {"type", "section"}})
        elif line_type == "introduction":
            content["introduction"].append(normalize_space(item["text"]))
        elif line_type == "data_analysis":
            content["data_analysis"].append(
                {
                    "number": item.get("number"),
                    "title": clean_dimension_title(item["title"]),
                    "analysis": normalize_space(item["analysis"]),
                }
            )
        elif line_type == "positive_feedback":
            content["positive_feedback"].append({"title": normalize_space(item["title"]), "body": normalize_space(item["body"])})
        elif line_type == "negative_feedback":
            content["negative_feedback"].append({"title": normalize_space(item["title"]), "body": normalize_space(item["body"])})
        elif line_type == "summary_recommendation":
            content["summary_recommendations"].append({"title": normalize_space(item["title"]), "body": normalize_space(item["body"])})
        else:
            raise ValueError(f"Unsupported JSONL record type at line {line_no}: {line_type}")
    return meta, content


def parse_direct_json(path: Path) -> dict:
    text = path.read_text(encoding="utf-8").strip()
    fenced = re.fullmatch(r"```json\s*(.*?)\s*```", text, flags=re.DOTALL)
    if fenced:
        text = fenced.group(1)
    return json.loads(text)


def build_attachment_questions(questionnaire: dict) -> list[dict]:
    return [
        {
            "question_ref": f"q{q['number']:02d}",
            "number": q["number"],
            "question": q["question"],
            "options": [
                {
                    "code": option["label"],
                    "text": option["text"],
                    "pct": option["pct"],
                }
                for option in question_options(q)
            ],
        }
        for q in questionnaire.get("questions", [])
    ]


def first_sample_size(questions: list[dict]) -> str | None:
    totals = [str(question.get("total", "")).strip() for question in questions if str(question.get("total", "")).strip()]
    if not totals:
        return None
    if len(set(totals)) == 1:
        return f"N={totals[0]}"
    return None


def pct_to_float(text: str) -> float:
    value = str(text).strip().rstrip("%")
    if not value:
        return 0.0
    return round(float(value), 2)


def build_payload(questionnaire: dict, meta: dict, content: dict, cli_args: argparse.Namespace) -> dict:
    questions = questionnaire.get("questions", [])
    data_items = content["data_analysis"]
    if len(data_items) != len(questions):
        raise ValueError(
            f"chapter 2 item count mismatch: markdown has {len(data_items)} items, questionnaire has {len(questions)} questions"
        )

    product = cli_args.product or meta.get("product") or meta.get("品种")
    region = cli_args.region or meta.get("region") or meta.get("地区")
    time_text = cli_args.time or meta.get("time") or meta.get("时间")
    report_type = cli_args.report_type or meta.get("report_type") or meta.get("报告类型") or DEFAULT_REPORT_TYPE
    data_issue = cli_args.data_issue if cli_args.data_issue is not None else meta.get("data_issue") or meta.get("数据问题")
    unresolved = cli_args.unresolved if cli_args.unresolved is not None else meta.get("unresolved") or meta.get("未解决问题")

    if not product:
        raise ValueError("Missing product. Provide front matter `product:`/`品种:` or pass `--product`.")
    if not region:
        raise ValueError("Missing region. Provide front matter `region:`/`地区:` or pass `--region`.")
    if not content["introduction"]:
        raise ValueError("Introduction section is empty.")
    if not content["positive_feedback"]:
        raise ValueError("Positive feedback section is empty.")
    if not content["negative_feedback"]:
        raise ValueError("Negative feedback section is empty.")

    analysis_items = []
    for idx, (question, item) in enumerate(zip(questions, data_items), start=1):
        analysis_items.append(
            {
                "number": idx,
                "title": item["title"],
                "analysis": item["analysis"],
                "question": question["question"],
                "options": question_options(question),
            }
        )

    return {
        "product": product,
        "region": region,
        "time": time_text,
        "report_type": report_type,
        "introduction": content["introduction"],
        "data_analysis": analysis_items,
        "positive_feedback": content["positive_feedback"],
        "negative_feedback": content["negative_feedback"],
        "attachment_questions": build_attachment_questions(questionnaire),
        "checks": {
            "data_issue": data_issue,
            "unresolved": unresolved,
        },
    }


def build_payload_v2(questionnaire: dict, meta: dict, content: dict, cli_args: argparse.Namespace) -> dict:
    questions = questionnaire.get("questions", [])
    legacy_payload = build_payload(questionnaire, meta, content, cli_args)
    chapter2_items = []
    for item in legacy_payload["data_analysis"]:
        categories = [f"{option['label']}.{option['text']}" for option in item["options"]]
        values = [pct_to_float(option["pct"]) for option in item["options"]]
        chapter2_items.append(
            {
                "slot": f"item_{item['number']:02d}",
                "heading": f"· {item['title']}",
                "title": item["title"],
                "analysis": item["analysis"],
                "question_ref": f"q{item['number']:02d}",
                "chart_ref": f"chart_slot_{item['number']:02d}",
                "chart": {
                    "type": "bar_horizontal",
                    "title": f"图{item['number']} {item['title']}",
                    "categories": categories,
                    "series": [
                        {
                            "name": "占比",
                            "values": values,
                        }
                    ],
                },
            }
        )

    return {
        "meta": {
            "product": legacy_payload["product"],
            "region": legacy_payload["region"],
            "time": legacy_payload.get("time"),
            "report_type": legacy_payload["report_type"],
            "report_title": meta.get("report_title") or meta.get("报告标题") or legacy_payload["report_type"],
            "sample_size": first_sample_size(questions),
        },
        "template": {
            "template_id": cli_args.template_id,
            "template_file": cli_args.template_file,
            "version": "v1",
        },
        "intro": {
            "paragraphs": legacy_payload["introduction"],
        },
        "chapter2": {
            "items": chapter2_items,
        },
        "feedback": {
            "positive": [
                {"slot": f"item_{index:02d}", **item}
                for index, item in enumerate(legacy_payload["positive_feedback"], start=1)
            ],
            "negative": [
                {"slot": f"item_{index:02d}", **item}
                for index, item in enumerate(legacy_payload["negative_feedback"], start=1)
            ],
        },
        "summary": {
            "recommendations": [
                {"slot": f"item_{index:02d}", **item}
                for index, item in enumerate(content.get("summary_recommendations", []), start=1)
            ],
        },
        "attachment": {
            "questions": build_attachment_questions(questionnaire),
        },
        "checks": legacy_payload["checks"],
    }


def validate_payload(payload: dict) -> None:
    required = [
        "product",
        "region",
        "report_type",
        "introduction",
        "data_analysis",
        "positive_feedback",
        "negative_feedback",
        "attachment_questions",
    ]
    missing = [key for key in required if key not in payload]
    if missing:
        raise ValueError(f"Payload missing required keys: {', '.join(missing)}")
    if len(payload["data_analysis"]) != len(payload["attachment_questions"]):
        raise ValueError("Payload mismatch: `data_analysis` count must equal `attachment_questions` count.")
    for index, item in enumerate(payload["data_analysis"], start=1):
        if not item.get("title") or not item.get("analysis"):
            raise ValueError(f"data_analysis item {index} missing title or analysis.")
        if not item.get("options"):
            raise ValueError(f"data_analysis item {index} missing options.")


def validate_payload_v2(payload: dict) -> None:
    required = ["meta", "template", "intro", "chapter2", "feedback", "attachment", "checks"]
    missing = [key for key in required if key not in payload]
    if missing:
        raise ValueError(f"Payload v2 missing required keys: {', '.join(missing)}")
    items = payload["chapter2"].get("items", [])
    questions = payload["attachment"].get("questions", [])
    if len(items) != len(questions):
        raise ValueError("Payload v2 mismatch: chapter2 item count must equal attachment question count.")
    for index, item in enumerate(items, start=1):
        if not item.get("chart_ref") or not item.get("question_ref"):
            raise ValueError(f"chapter2 item {index} missing chart_ref or question_ref.")
        categories = item.get("chart", {}).get("categories", [])
        series = item.get("chart", {}).get("series", [])
        if not categories or not series:
            raise ValueError(f"chapter2 item {index} missing chart categories or series.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Build a valid patient survey report payload JSON from questionnaire JSON and structured report content.")
    parser.add_argument("questionnaire_json", help="Path to normalized questionnaire JSON")
    parser.add_argument("report_content", help="Path to structured report content (.md, .jsonl, or direct JSON)")
    parser.add_argument("-o", "--output", required=True, help="Output payload JSON path")
    parser.add_argument("--product", help="Override product name")
    parser.add_argument("--region", help="Override region")
    parser.add_argument("--time", help="Override time")
    parser.add_argument("--report-type", help="Override report type")
    parser.add_argument("--data-issue", help="Known data issue note")
    parser.add_argument("--unresolved", help="Known unresolved issue note")
    parser.add_argument("--schema-version", choices=["v1", "v2"], default="v1", help="Output payload schema version")
    parser.add_argument("--template-id", default="patient-unified-v1", help="Template id for payload v2")
    parser.add_argument("--template-file", default="templates/patient-unified-v1.docx", help="Template file path for payload v2")
    args = parser.parse_args()

    questionnaire = json.loads(Path(args.questionnaire_json).read_text(encoding="utf-8"))
    content_path = Path(args.report_content)

    if content_path.suffix.lower() == ".jsonl":
        meta, content = parse_jsonl_content(content_path)
        payload = build_payload_v2(questionnaire, meta, content, args) if args.schema_version == "v2" else build_payload(questionnaire, meta, content, args)
    elif content_path.suffix.lower() == ".json":
        payload = parse_direct_json(content_path)
    else:
        meta, content = parse_markdown_content(content_path)
        payload = build_payload_v2(questionnaire, meta, content, args) if args.schema_version == "v2" else build_payload(questionnaire, meta, content, args)

    if args.schema_version == "v2" or ("template" in payload and "chapter2" in payload):
        validate_payload_v2(payload)
    else:
        validate_payload(payload)
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(output)


if __name__ == "__main__":
    main()
