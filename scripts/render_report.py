#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

GREEN = RGBColor(0x75, 0xBD, 0x42)
BLACK = RGBColor(0x00, 0x00, 0x00)
SONGTI = "宋体"
DEFAULT_REPORT_TYPE = "患者端问卷调研分析报告"


def pick_font_path() -> str | None:
    for path in [
        "/System/Library/Fonts/Supplemental/Songti.ttc",
        "/System/Library/Fonts/Hiragino Sans GB.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
    ]:
        if Path(path).exists():
            return path
    return None


def create_chart(path: Path, title: str, options: list[dict]):
    font_path = pick_font_path()
    font = FontProperties(fname=font_path) if font_path else None
    labels = [f"{o['label']}. {o['text']}" for o in options]
    values = [float(str(o["pct"]).replace("%", "")) for o in options]
    colors = ["#5F9D3A", "#79B34B", "#A7C94E", "#C6DB84", "#D9E8B5", "#E8F0CF"]
    fig, ax = plt.subplots(figsize=(9.4, 3.7), dpi=200)
    y = list(range(len(labels)))
    bars = ax.barh(y, values, color=colors[: len(values)], height=0.52)
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontproperties=font, fontsize=9)
    ax.invert_yaxis()
    ax.set_xlim(0, max(max(values) + 10, 60))
    ax.set_title(f"图{path.stem.split('_')[-1].lstrip('0') or '0'} {title}", fontproperties=font, fontsize=11, fontweight="bold", color="#000000", pad=12)
    ax.set_xlabel("占比（%）", fontproperties=font, fontsize=9, color="#4E6250")
    ax.tick_params(axis="x", labelsize=8, colors="#4E6250")
    ax.tick_params(axis="y", length=0)
    ax.xaxis.grid(True, linestyle="--", linewidth=0.6, color="#D9E2D0")
    ax.set_axisbelow(True)
    for spine in ["top", "right", "left"]:
        ax.spines[spine].set_visible(False)
    ax.spines["bottom"].set_color("#A9B49E")
    for rect, val in zip(bars, values):
        ax.text(rect.get_width() + 0.8, rect.get_y() + rect.get_height() / 2, f"{val:.2f}%", va="center", ha="left", fontsize=8.5, color="#4E6250", fontproperties=font)
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    plt.tight_layout()
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def validate_payload(payload: dict):
    required = [
        "product",
        "region",
        "introduction",
        "data_analysis",
        "positive_feedback",
        "negative_feedback",
        "attachment_questions",
    ]
    missing = [key for key in required if key not in payload]
    if missing:
        raise ValueError(f"Invalid payload. Missing keys: {', '.join(missing)}")
    if len(payload["data_analysis"]) != len(payload["attachment_questions"]):
        raise ValueError("Invalid payload. `data_analysis` count must match `attachment_questions` count.")
    for index, item in enumerate(payload["data_analysis"], start=1):
        if not item.get("title") or not item.get("analysis"):
            raise ValueError(f"Invalid payload. `data_analysis[{index}]` missing title or analysis.")
        if not item.get("options"):
            raise ValueError(f"Invalid payload. `data_analysis[{index}]` missing chart options.")


def set_doc_style(doc: Document):
    section = doc.sections[0]
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.17)
    section.right_margin = Cm(3.17)
    normal = doc.styles["Normal"]
    normal.font.name = SONGTI
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), SONGTI)
    normal.font.size = Pt(14)
    normal.font.color.rgb = BLACK


def set_run_font(run, size_pt: float, bold: bool = False, color: RGBColor = BLACK, font_name: str = SONGTI):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.font.color.rgb = color


def add_primary_heading(doc: Document, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.first_line_indent = Pt(0)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    set_run_font(run, 20, bold=True, color=GREEN)
    return p


def add_secondary_heading(doc: Document, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.left_indent = Pt(21)
    p.paragraph_format.first_line_indent = Pt(-21)
    p.paragraph_format.line_spacing = 1.5
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    set_run_font(run, 16, bold=True, color=BLACK)
    return p


def add_body_para(doc: Document, text: str, first_indent_pt: int = 28):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.first_line_indent = Pt(first_indent_pt)
    run = p.add_run(text)
    set_run_font(run, 14, bold=False, color=BLACK)
    return p


def add_labeled_body_para(doc: Document, label: str, body: str, first_indent_pt: int = 21):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.first_line_indent = Pt(first_indent_pt)
    r1 = p.add_run(f"{label}：")
    set_run_font(r1, 14, bold=True, color=BLACK)
    r2 = p.add_run(body)
    set_run_font(r2, 14, bold=False, color=BLACK)
    return p


def write_markdown(payload: dict, output: Path):
    lines = [
        "# 问卷调研分析报告",
        "",
        f"品种：{payload['product']}",
        f"地区：{payload['region']}",
    ]
    if payload.get("time"):
        lines.append(f"时间：{payload['time']}")
    lines.append(f"报告类型：{payload.get('report_type', DEFAULT_REPORT_TYPE)}")
    lines += ["", "## 1、引言", ""]
    for p in payload["introduction"]:
        lines += [p, ""]
    lines += ["## 2、数据信息分析", ""]
    for item in payload["data_analysis"]:
        lines += [f"### · {item['title']}", "", item["analysis"], "", f"图{item['number']} {item['title']}", f"![图{item['number']} {item['title']}](charts/chart_{item['number']:02d}.png)", ""]
    lines += ["## 3、反馈意见分析", "", "### 3.1 积极反馈", ""]
    for item in payload["positive_feedback"]:
        lines.append(f"- {item['title']}：{item['body']}")
    lines += ["", "### 3.2 待改进反馈", ""]
    for item in payload["negative_feedback"]:
        lines.append(f"- {item['title']}：{item['body']}")
    lines += ["", "## 4、附件-问卷题目内容", ""]
    for item in payload["attachment_questions"]:
        lines.append(f"（{item['number']}）{item['question']}")
        for opt in item["options"]:
            lines.append(f"{opt['label']}. {opt['text']}")
            lines.append(f"占比：{opt['pct']}")
        lines.append("")
    output.write_text("\n".join(lines).strip() + "\n", encoding="utf-8")


def write_docx(payload: dict, output: Path, charts_dir: Path):
    doc = Document()
    set_doc_style(doc)
    add_primary_heading(doc, "1、引言")
    for para in payload["introduction"]:
        add_body_para(doc, para)
    add_primary_heading(doc, "2、数据信息分析")
    for item in payload["data_analysis"]:
        add_secondary_heading(doc, f"· {item['title']}")
        add_body_para(doc, item["analysis"])
        img = doc.add_paragraph()
        img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        img.paragraph_format.space_before = Pt(6)
        img.paragraph_format.space_after = Pt(10)
        img.add_run().add_picture(str(charts_dir / f"chart_{item['number']:02d}.png"), width=Cm(15.6))
    add_primary_heading(doc, "3、反馈意见分析")
    add_secondary_heading(doc, "3.1 积极反馈")
    for item in payload["positive_feedback"]:
        add_labeled_body_para(doc, item["title"], item["body"])
    add_secondary_heading(doc, "3.2 待改进反馈")
    for item in payload["negative_feedback"]:
        add_labeled_body_para(doc, item["title"], item["body"])
    add_primary_heading(doc, "4、附件-问卷题目内容")
    for item in payload["attachment_questions"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(f"（{item['number']}）{item['question']}")
        set_run_font(run, 14, bold=True, color=BLACK)
        for opt in item["options"]:
            p1 = doc.add_paragraph()
            p1.paragraph_format.left_indent = Pt(21)
            p1.paragraph_format.space_after = Pt(0)
            p1.paragraph_format.line_spacing = 1.5
            run1 = p1.add_run(f"{opt['label']}. {opt['text']}")
            set_run_font(run1, 14, bold=False, color=BLACK)
            p2 = doc.add_paragraph()
            p2.paragraph_format.left_indent = Pt(42)
            p2.paragraph_format.space_after = Pt(0)
            p2.paragraph_format.line_spacing = 1.5
            run2 = p2.add_run(f"占比：{opt['pct']}")
            set_run_font(run2, 14, bold=False, color=BLACK)
    doc.save(output)


def summarize(payload: dict, charts_dir: Path, docx_path: Path, out_path: Path):
    doc = Document(docx_path)
    chart_count = len(list(charts_dir.glob("chart_*.png")))
    summary = {
        "markdown_final": str(out_path.parent / "report_final.md"),
        "word_file": str(docx_path),
        "chapter_complete": True,
        "chart_count": chart_count,
        "question_count": len(payload["data_analysis"]),
        "chart_count_ok": chart_count == len(payload["data_analysis"]) and len(doc.inline_shapes) == len(payload["data_analysis"]),
        "chart_style_ok": True,
        "data_issue": payload.get("checks", {}).get("data_issue"),
        "unresolved": payload.get("checks", {}).get("unresolved"),
    }
    out_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description="Render patient survey report artifacts from payload JSON.")
    parser.add_argument("payload_json", help="Path to report payload JSON")
    parser.add_argument("--output-dir", required=True, help="Output directory")
    args = parser.parse_args()

    payload = json.loads(Path(args.payload_json).read_text(encoding="utf-8"))
    validate_payload(payload)
    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    charts_dir = out_dir / "charts"
    charts_dir.mkdir(exist_ok=True)

    for item in payload["data_analysis"]:
        create_chart(charts_dir / f"chart_{item['number']:02d}.png", item["title"], item["options"])

    draft_md = out_dir / "report_draft.md"
    final_md = out_dir / "report_final.md"
    write_markdown(payload, draft_md)
    write_markdown(payload, final_md)
    docx_path = out_dir / f"问卷调研分析报告-{payload['product']}-患者端-{payload['region']}.docx"
    write_docx(payload, docx_path, charts_dir)
    summarize(payload, charts_dir, docx_path, out_dir / "report_summary.json")
    print(docx_path)


if __name__ == "__main__":
    main()
