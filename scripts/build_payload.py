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
    if not stripped.startswith("---\n"):
        return {}, stripped
    parts = stripped.split("\n---\n", 1)
    if len(parts) != 2:
        return {}, stripped
    meta_block, body = parts
    meta = {}
    for line in meta_block.splitlines()[1:]:
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

    current_section: str | None = None
    current_item: dict | None = None

    def flush_item() -> None:
        nonlocal current_item
        if not current_item:
            return
        current_item["analysis"] = normalize_space(" ".join(current_item.pop("lines")))
        data_items.append(current_item)
        current_item = None

    for line in lines:
        heading = re.match(r"^(#{2,4})\s+(.+?)\s*$", line)
        if heading:
            level = len(heading.group(1))
            title = normalize_space(heading.group(2))
            if "引言" in title:
                flush_item()
                current_section = "introduction"
                continue
            if "数据信息分析" in title:
                flush_item()
                current_section = "data_analysis"
                continue
            if "反馈意见分析" in title:
                flush_item()
                continue
            if "积极反馈" in title:
                flush_item()
                current_section = "positive_feedback"
                continue
            if "待改进反馈" in title:
                flush_item()
                current_section = "negative_feedback"
                continue
            if "附件" in title:
                flush_item()
                current_section = "attachment"
                continue
            if current_section == "data_analysis" and level >= 3:
                flush_item()
                current_item = {"title": clean_dimension_title(title), "lines": []}
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

    flush_item()

    content = {
        "introduction": parse_paragraphs(intro_lines),
        "data_analysis": data_items,
        "positive_feedback": parse_label_body_blocks(pos_lines, "positive_feedback"),
        "negative_feedback": parse_label_body_blocks(neg_lines, "negative_feedback"),
    }
    return meta, content


def parse_jsonl_content(path: Path) -> tuple[dict, dict]:
    meta: dict = {}
    content = {
        "introduction": [],
        "data_analysis": [],
        "positive_feedback": [],
        "negative_feedback": [],
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
            "number": q["number"],
            "question": q["question"],
            "options": question_options(q),
        }
        for q in questionnaire.get("questions", [])
    ]


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
    args = parser.parse_args()

    questionnaire = json.loads(Path(args.questionnaire_json).read_text(encoding="utf-8"))
    content_path = Path(args.report_content)

    if content_path.suffix.lower() == ".jsonl":
        meta, content = parse_jsonl_content(content_path)
        payload = build_payload(questionnaire, meta, content, args)
    elif content_path.suffix.lower() == ".json":
        payload = parse_direct_json(content_path)
    else:
        meta, content = parse_markdown_content(content_path)
        payload = build_payload(questionnaire, meta, content, args)

    validate_payload(payload)
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(output)


if __name__ == "__main__":
    main()
