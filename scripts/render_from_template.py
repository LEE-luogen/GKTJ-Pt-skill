#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.text.paragraph import Paragraph


def load_payload(path: Path) -> dict:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if "template" not in payload or "chapter2" not in payload:
        raise ValueError("Expected template-driven payload v2.")
    return payload


def load_manifest(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def clear_paragraph(paragraph: Paragraph) -> None:
    p = paragraph._element
    for child in list(p):
        p.remove(child)


def set_paragraph_text(paragraph: Paragraph, text: str) -> None:
    clear_paragraph(paragraph)
    if text:
        paragraph.add_run(text)


def split_into_paragraphs(text: str, target_count: int) -> list[str]:
    raw_parts = [part.strip() for part in str(text).splitlines() if part.strip()]
    if not raw_parts:
        raw_parts = [str(text).strip()] if str(text).strip() else []
    if target_count <= 1:
        return [raw_parts[0] if raw_parts else ""]
    if len(raw_parts) >= target_count:
        parts = raw_parts[: target_count - 1]
        parts.append(" ".join(raw_parts[target_count - 1 :]).strip())
        return parts

    sentences: list[str] = []
    current = ""
    for char in " ".join(raw_parts):
        current += char
        if char in "。！？":
            sentences.append(current.strip())
            current = ""
    if current.strip():
        sentences.append(current.strip())
    if not sentences:
        sentences = [" ".join(raw_parts).strip()]

    if len(sentences) <= target_count:
        return sentences + [""] * (target_count - len(sentences))

    bucket_size = (len(sentences) + target_count - 1) // target_count
    grouped = []
    for index in range(0, len(sentences), bucket_size):
        grouped.append("".join(sentences[index : index + bucket_size]).strip())
    if len(grouped) < target_count:
        grouped.extend([""] * (target_count - len(grouped)))
    return grouped[:target_count]


def build_feedback_lines(items: Iterable[dict], slots: int) -> list[str]:
    lines = [f"{item['title']}：{item['body']}".strip("：") for item in items]
    if len(lines) < slots:
        lines.extend([""] * (slots - len(lines)))
    return lines[:slots]


def trim_document_after_paragraph(document: Document, paragraph_index: int) -> None:
    paragraphs = document.paragraphs
    for paragraph in reversed(paragraphs[paragraph_index + 1 :]):
        parent = paragraph._element.getparent()
        parent.remove(paragraph._element)


def render_attachment(document: Document, questions: list[dict]) -> None:
    for item in questions:
        document.add_paragraph(f"{item['number']}.{item['question']}")
        for option in item.get("options", []):
            document.add_paragraph(f"{option['code']}.{option['text']}")
            document.add_paragraph(f"占比：{option['pct']}")


def render_from_template(payload: dict, template_path: Path, manifest_path: Path, output_path: Path) -> Path:
    shutil.copyfile(template_path, output_path)
    manifest = load_manifest(manifest_path)
    document = Document(output_path)

    intro_paragraphs = payload.get("intro", {}).get("paragraphs", [])
    intro = manifest["intro"]
    set_paragraph_text(document.paragraphs[intro["report_background_paragraph"]], intro_paragraphs[0] if len(intro_paragraphs) > 0 else "")
    set_paragraph_text(document.paragraphs[intro["data_source_paragraph"]], intro_paragraphs[1] if len(intro_paragraphs) > 1 else "")

    items = payload["chapter2"]["items"]
    blocks = manifest["chapter2"]["blocks"]
    if len(items) > len(blocks):
        raise ValueError(f"Template supports {len(blocks)} chapter2 blocks, got {len(items)}.")

    for index, block in enumerate(blocks):
        title_paragraph = document.paragraphs[block["title_paragraph"]]
        analysis_paragraphs = [document.paragraphs[position] for position in block["analysis_paragraphs"]]
        if index < len(items):
            item = items[index]
            set_paragraph_text(title_paragraph, item.get("title") or item.get("heading", ""))
            chunks = split_into_paragraphs(item.get("analysis", ""), len(analysis_paragraphs))
            for paragraph, chunk in zip(analysis_paragraphs, chunks):
                set_paragraph_text(paragraph, chunk)
        else:
            set_paragraph_text(title_paragraph, "")
            for paragraph in analysis_paragraphs:
                set_paragraph_text(paragraph, "")

    positive_lines = build_feedback_lines(payload.get("feedback", {}).get("positive", []), len(manifest["feedback"]["positive_paragraphs"]))
    for line, paragraph_index in zip(positive_lines, manifest["feedback"]["positive_paragraphs"]):
        set_paragraph_text(document.paragraphs[paragraph_index], line)

    negative_lines = build_feedback_lines(payload.get("feedback", {}).get("negative", []), len(manifest["feedback"]["negative_paragraphs"]))
    for line, paragraph_index in zip(negative_lines, manifest["feedback"]["negative_paragraphs"]):
        set_paragraph_text(document.paragraphs[paragraph_index], line)

    recommendations = payload.get("summary", {}).get("recommendations", [])
    title_slots = manifest["summary"]["recommendation_title_paragraphs"]
    body_slots = manifest["summary"]["recommendation_body_paragraphs"]
    for slot_index, (title_paragraph, body_paragraph) in enumerate(zip(title_slots, body_slots)):
        item = recommendations[slot_index] if slot_index < len(recommendations) else {}
        set_paragraph_text(document.paragraphs[title_paragraph], item.get("title", ""))
        set_paragraph_text(document.paragraphs[body_paragraph], item.get("body", ""))

    attachment_heading = manifest["attachment"]["section_heading_paragraph"]
    trim_document_after_paragraph(document, attachment_heading)
    render_attachment(document, payload.get("attachment", {}).get("questions", []))
    document.save(output_path)
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Render a patient report from the fixed docx template and payload v2.")
    parser.add_argument("payload_json", help="Path to report payload v2 json")
    parser.add_argument("-o", "--output", required=True, help="Output docx path")
    parser.add_argument("--template", help="Override template docx path")
    parser.add_argument("--manifest", help="Override template manifest path")
    args = parser.parse_args()

    payload = load_payload(Path(args.payload_json))
    base_dir = Path(__file__).resolve().parents[1]
    template_path = Path(args.template) if args.template else base_dir / payload["template"]["template_file"]
    manifest_path = Path(args.manifest) if args.manifest else template_path.with_suffix(".manifest.json")
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    render_from_template(payload, template_path, manifest_path, output_path)
    print(output_path)


if __name__ == "__main__":
    main()
