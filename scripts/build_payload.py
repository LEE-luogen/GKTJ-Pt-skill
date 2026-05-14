#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

DEFAULT_REPORT_TYPE = "患者端问卷调研分析报告"
CHAPTER2_TITLE_MIN_LEN = 4
CHAPTER2_TITLE_MAX_LEN = 9
ANALYSIS_MIN_LEN = 300
ANALYSIS_MAX_LEN = 350

CHAPTER2_TITLE_FALLBACK = [
    (r"用药.*原因|主要原因", "用药原因"),
    (r"血压.*监测|监测.*血压", "血压监测"),
    (r"降压.*认知|作用.*了解|长期服用", "降压认知"),
    (r"饮食|饮酒|高盐", "饮食控制"),
    (r"知识.*获取|获取.*知识|说明书|药师|医生指导", "知识获取"),
    (r"剂量|调量|加量|减量", "剂量调整"),
    (r"服药.*提醒|提醒.*服药|按时服药", "服药提醒"),
    (r"联合用药|保健品", "联合用药"),
    (r"健康教育|讲座|培训", "健教参与"),
    (r"依从性|坚持用药", "依从提升"),
]


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
    text = re.sub(r"^[((]?\d+[)).、\s]+", "", text)
    return text.strip()


def clean_question_text(raw: str) -> str:
    text = normalize_space(raw)
    text = re.sub(r"^[((]?\d+[)).、.\s]+", "", text)
    return text.strip()


def text_len(text: str) -> int:
    return len(re.sub(r"\s+", "", text))


def compact_text(text: str) -> str:
    return re.sub(r"\s+", "", normalize_space(text))


def ensure_sentence_end(text: str, max_len: int | None = None) -> str:
    stripped = text.strip()
    if not stripped:
        return stripped
    if stripped[-1] not in "。!?":
        if max_len is not None and text_len(stripped) >= max_len:
            return stripped[:-1] + "。"
        return stripped + "。"
    return stripped


def normalize_chapter2_title(ai_title: str, question_text: str) -> str:
    """
    标题生成规则：AI 标题优先，脚本只做合规校验和例外处理。
    1. AI 在 markdown 中写的标题为首选，清洗后若长度在 4-9 字范围内，直接使用。
    2. 若 AI 标题过长（>9 字），尝试去除常见后缀词后截断。
    3. 若 AI 标题过短（<4 字）或为空，从题干中提取摘要。
    4. 以上均失败时，降级使用硬编码 fallback 映射表。
    """
    clean_ai_title = clean_dimension_title(ai_title)
    clean_question = clean_question_text(question_text)

    # 1. AI 标题优先：清洗后直接校验长度
    if clean_ai_title:
        title_len = len(clean_ai_title)
        if CHAPTER2_TITLE_MIN_LEN <= title_len <= CHAPTER2_TITLE_MAX_LEN:
            return clean_ai_title

        # AI 标题过长：去后缀后截断
        if title_len > CHAPTER2_TITLE_MAX_LEN:
            trimmed = re.sub(
                r"(情况|结构|分析|问题|程度|方式|路径|调研|问卷|是否|主要|相关|了解多少|是什么|怎么样)$", "",
                clean_ai_title
            )
            if CHAPTER2_TITLE_MIN_LEN <= len(trimmed) <= CHAPTER2_TITLE_MAX_LEN:
                return trimmed
            return trimmed[:CHAPTER2_TITLE_MAX_LEN]

    # 2. AI 标题过短或为空：从题干提取标题
    if clean_question:
        simplified = re.sub(r"[，。！？：；、\u201c\u201d''（）()【】\[\]\-\u2014\s]", "", clean_question)
        simplified = re.sub(r"^(关于|患者|本题|该题|单题|您)", "", simplified)
        simplified = re.sub(
            r"(情况|结构|分析|问题|程度|方式|路径|调研|问卷|是否|主要|相关|了解多少|是什么|怎么样|的主要原因是什么|的频率是)$",
            "", simplified
        )
        if CHAPTER2_TITLE_MIN_LEN <= len(simplified) <= CHAPTER2_TITLE_MAX_LEN:
            return simplified
        if len(simplified) > CHAPTER2_TITLE_MAX_LEN:
            return simplified[:CHAPTER2_TITLE_MAX_LEN]

    # 3. 从题干提取也失败：使用 hardcoded fallback
    corpus = " ".join([clean_ai_title, clean_question])
    for pattern, candidate in CHAPTER2_TITLE_FALLBACK:
        if re.search(pattern, corpus):
            return candidate

    # 4. 最终兜底
    fallback = re.sub(r"[^\u4e00-\u9fffA-Za-z0-9]", "", clean_ai_title or clean_question)
    if len(fallback) < CHAPTER2_TITLE_MIN_LEN:
        raise ValueError(f"Chapter 2 title too short to normalize: {ai_title or question_text}")
    return fallback[:CHAPTER2_TITLE_MAX_LEN]


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


def sanitize_intro_paragraph(text: str) -> str:
    value = normalize_space(text)
    value = re.sub(r"^#{1,6}\s+", "", value)
    value = re.sub(r"^(报告背景|数据来源)\s*", "", value)
    value = re.sub(r"^[：:、\-]\s*", "", value)
    return normalize_space(value)


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

    def is_heading(title: str, patterns: list[str]) -> bool:
        for pattern in patterns:
            if re.fullmatch(pattern, title):
                return True
        return False

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
            if is_heading(title, [r"引言", r"[一1]、引言"]):
                flush_item()
                flush_summary_item()
                current_section = "introduction"
                continue
            if is_heading(title, [r"数据信息分析", r"[二2]、数据信息分析"]):
                flush_item()
                flush_summary_item()
                current_section = "data_analysis"
                continue
            if is_heading(title, [r"反馈意见分析", r"[三3]、反馈意见分析"]):
                flush_item()
                flush_summary_item()
                continue
            if is_heading(title, [r"积极反馈", r"3\.1\s+积极反馈", r"3\.1积极反馈"]):
                flush_item()
                flush_summary_item()
                current_section = "positive_feedback"
                continue
            if is_heading(title, [r"待改进反馈", r"3\.2\s+待改进反馈", r"3\.2待改进反馈"]):
                flush_item()
                flush_summary_item()
                current_section = "negative_feedback"
                continue
            if is_heading(title, [r"综合分析与建议", r"[四4]、综合分析与建议"]):
                flush_item()
                flush_summary_item()
                current_section = "summary"
                continue
            if is_heading(title, [r"附件-问卷题目内容", r"[五5]、附件-问卷题目内容", r"附件"]):
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

    intro_paragraphs = [sanitize_intro_paragraph(p) for p in parse_paragraphs(intro_lines)]
    intro_paragraphs = [p for p in intro_paragraphs if p]
    content = {
        "introduction": intro_paragraphs,
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
            value = sanitize_intro_paragraph(item["text"])
            if value:
                content["introduction"].append(value)
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
            "question": clean_question_text(q["question"]),
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


def pick_meta_value(meta: dict, *keys: str) -> str | None:
    for key in keys:
        value = meta.get(key)
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    return None


def pct_to_float(text: str) -> float:
    value = str(text).strip().rstrip("%")
    if not value:
        return 0.0
    return round(float(value), 2)


ANALYSIS_OPENINGS = [
    "从患者实际选择结构来看,本题已经呈现出较清晰的主导方向,相关比例差异能够直接反映患者在这一维度上的真实关注重点。",
    "结合本题各选项的分布情况可以看出,患者反馈并不是随机分散的,而是围绕几个核心判断形成了相对稳定的选择结构。",
    "若把本题结果放回患者日常用药场景中理解,可以发现当前比例结构背后对应的是较为典型的行为习惯、认知差异与支持需求。",
    "从本题占比高低的组合关系判断,患者在这一问题上的反馈已经不只是单点意见,而是折射出一个较完整的管理状态与体验特征。",
    "本题的结果更适合从整体行为逻辑去理解,因为主要选项之间的比例关系已经能够说明患者在现实场景中的优先考虑因素。",
    "综合本题选项占比后可以发现,患者在这一维度上的选择带有明显规律,这种规律对理解其真实需求和后续支持重点都有直接意义。",
]


def format_pct_value(text: str) -> str:
    value = normalize_pct(text)
    return value if value else "0.00%"


def split_sentences(text: str) -> list[str]:
    return [item for item in re.split(r"(?<=[。!?])", text) if item]


def trim_analysis_length(text: str) -> str:
    normalized = compact_text(text)
    if text_len(normalized) <= ANALYSIS_MAX_LEN:
        return ensure_sentence_end(normalized, ANALYSIS_MAX_LEN)

    sentences = split_sentences(normalized)
    trimmed = ""
    for sentence in sentences:
        candidate = trimmed + sentence
        if text_len(candidate) <= ANALYSIS_MAX_LEN:
            trimmed = candidate
            continue
        if text_len(trimmed) >= ANALYSIS_MIN_LEN:
            break
        remaining = ANALYSIS_MAX_LEN - text_len(trimmed)
        if remaining > 0:
            trimmed += sentence[:remaining]
        break

    if text_len(trimmed) < ANALYSIS_MIN_LEN:
        trimmed = normalized[:ANALYSIS_MAX_LEN]
    return ensure_sentence_end(trimmed[:ANALYSIS_MAX_LEN], ANALYSIS_MAX_LEN)


def ensure_analysis_length_range(text: str, question: dict, title: str, index: int) -> str:
    base = compact_text(text)
    if ANALYSIS_MIN_LEN <= text_len(base) <= ANALYSIS_MAX_LEN:
        return ensure_sentence_end(base, ANALYSIS_MAX_LEN)

    options = question_options(question)
    sorted_options = sorted(options, key=lambda item: pct_to_float(item["pct"]), reverse=True)
    top = sorted_options[0] if sorted_options else {"text": "当前主流方向", "pct": "0.00%"}
    second = sorted_options[1] if len(sorted_options) > 1 else top
    third = sorted_options[2] if len(sorted_options) > 2 else second
    opening = ANALYSIS_OPENINGS[(index - 1) % len(ANALYSIS_OPENINGS)]

    supplemental_parts = [
        opening,
        (
            f"具体来看，选择\u201c{top['text']}\u201d的患者占{format_pct_value(top['pct'])}，"
            f"选择\u201c{second['text']}\u201d的患者占{format_pct_value(second['pct'])}，"
            f"而\u201c{third['text']}\u201d也保持在{format_pct_value(third['pct'])}。"
            "这说明患者反馈已形成较明确的主次层级，并非随机分散。"
        ),
        (
            f"如果结合题目\u201c{question.get('question', title)}\u201d本身去理解，"
            "高占比选项通常代表患者最容易感知、最常遇到或最愿意表达的现实问题；"
            "次高占比则提示群体内部仍有差异，这种差异常与管理习惯、信息获取和家庭支持有关。"
        ),
        (
            "从患者管理角度看，这一分布的意义不只是比较高低，"
            "更在于提醒后续沟通要围绕主流反馈提供更具体的解释、提醒和支持。"
        ),
        (
            f"因此，围绕\u201c{title}\u201d这一维度，更可行的工作重点是把问卷中已经显现出来的主流需求转化为可执行支持动作，"
            "例如在患者教育中突出最常见疑问、在随访中优先回应高占比反馈，"
            "并通过提醒和场景化说明帮助患者把理解转化为更稳定的行为。"
        ),
    ]

    candidate_parts = [base] if base else []

    for part in supplemental_parts:
        candidate_parts.append(compact_text(part))
        candidate = "".join(candidate_parts)
        if text_len(candidate) >= ANALYSIS_MIN_LEN:
            return trim_analysis_length(candidate)

    return trim_analysis_length("".join(candidate_parts))


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
    survey_period = cli_args.survey_period or pick_meta_value(meta, "survey_period", "问卷收集时间", "调研时间", "调查时间")
    issued_count = cli_args.issued_count or pick_meta_value(meta, "issued_count", "发放问卷数", "发放数量")
    valid_count = cli_args.valid_count or pick_meta_value(meta, "valid_count", "有效问卷数", "有效样本数", "回收有效问卷数")
    sample_size = cli_args.sample_size or pick_meta_value(meta, "sample_size", "样本数", "样本量") or first_sample_size(questions)

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
        normalized_title = normalize_chapter2_title(item["title"], question["question"])
        analysis_text = ensure_analysis_length_range(item["analysis"], question, normalized_title, idx)
        analysis_items.append(
            {
                "number": idx,
                "title": normalized_title,
                "analysis": analysis_text,
                "question": clean_question_text(question["question"]),
                "options": question_options(question),
            }
        )

    return {
        "product": product,
        "region": region,
        "time": time_text,
        "report_type": report_type,
        "introduction": content["introduction"],
        "survey_period": survey_period,
        "issued_count": issued_count,
        "valid_count": valid_count,
        "sample_size": sample_size,
        "data_analysis": analysis_items,
        "positive_feedback": content["positive_feedback"],
        "negative_feedback": content["negative_feedback"],
        "summary_recommendations": content.get("summary_recommendations", []),
        "attachment_questions": build_attachment_questions(questionnaire),
        "checks": {
            "data_issue": data_issue,
            "unresolved": unresolved,
        },
    }


def build_payload_v2(questionnaire: dict, meta: dict, content: dict, cli_args: argparse.Namespace) -> dict:
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
            "report_title": meta.get("report_title") or meta.get("报告标题") or "问卷调研分析报告",
            "sample_size": legacy_payload.get("sample_size"),
            "survey_period": legacy_payload.get("survey_period"),
            "issued_count": legacy_payload.get("issued_count"),
            "valid_count": legacy_payload.get("valid_count"),
        },
        "template": {
            "template_id": cli_args.template_id,
            "template_file": cli_args.template_file,
            "version": "v2-body-only",
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
            "recommendations": [{"slot": f"item_{index:02d}", **item} for index, item in enumerate(legacy_payload["summary_recommendations"], start=1)],
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
        "summary_recommendations",
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
        title_len = len(str(item["title"]).strip())
        if not CHAPTER2_TITLE_MIN_LEN <= title_len <= CHAPTER2_TITLE_MAX_LEN:
            raise ValueError(f"data_analysis item {index} title must be 4-9 characters (got {title_len}).")
        if not item.get("options"):
            raise ValueError(f"data_analysis item {index} missing options.")


def validate_payload_v2(payload: dict) -> None:
    required = ["meta", "template", "intro", "chapter2", "feedback", "summary", "attachment", "checks"]
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
        title_len = len(str(item.get("title", "")).strip())
        if not CHAPTER2_TITLE_MIN_LEN <= title_len <= CHAPTER2_TITLE_MAX_LEN:
            raise ValueError(f"chapter2 item {index} title must be 4-9 characters (got {title_len}).")
        categories = item.get("chart", {}).get("categories", [])
        series = item.get("chart", {}).get("series", [])
        if not categories or not series:
            raise ValueError(f"chapter2 item {index} missing chart categories or series.")
        analysis_len = text_len(str(item.get("analysis", "")))
        if not ANALYSIS_MIN_LEN <= analysis_len <= ANALYSIS_MAX_LEN:
            raise ValueError(f"chapter2 item {index} analysis must be 300-350 characters.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Build a valid patient survey report payload JSON from questionnaire JSON and structured report content.")
    parser.add_argument("questionnaire_json", help="Path to normalized questionnaire JSON")
    parser.add_argument("report_content", help="Path to structured report content (.md, .jsonl, or direct JSON)")
    parser.add_argument("-o", "--output", required=True, help="Output payload JSON path")
    parser.add_argument("--product", help="Override product name")
    parser.add_argument("--region", help="Override region")
    parser.add_argument("--time", help="Override time")
    parser.add_argument("--report-type", help="Override report type")
    parser.add_argument("--survey-period", help="Override questionnaire collection period")
    parser.add_argument("--sample-size", help="Override sample size text, e.g. N=2205")
    parser.add_argument("--issued-count", help="Override issued questionnaire count")
    parser.add_argument("--valid-count", help="Override valid questionnaire count")
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
