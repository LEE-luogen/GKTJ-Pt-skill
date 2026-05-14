#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

CURRENT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = CURRENT_DIR.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from scripts.render_report import create_chart, validate_payload, write_docx


def load_payload(path: Path) -> dict:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if "meta" not in payload or "chapter2" not in payload or "attachment" not in payload:
        raise ValueError("Expected patient report payload v2.")
    return payload


def v2_to_legacy_payload(payload: dict) -> dict:
    attachment_questions = payload.get("attachment", {}).get("questions", [])
    attachment_map = {item.get("question_ref"): item for item in attachment_questions}

    data_analysis = []
    for index, item in enumerate(payload.get("chapter2", {}).get("items", []), start=1):
        question_ref = item.get("question_ref")
        attachment = attachment_map.get(question_ref, {})
        data_analysis.append(
            {
                "number": index,
                "title": item.get("title") or item.get("heading", ""),
                "analysis": item.get("analysis", ""),
                "question": attachment.get("question", ""),
                "options": [
                    {
                        "label": option.get("code", ""),
                        "text": option.get("text", ""),
                        "pct": option.get("pct", ""),
                    }
                    for option in attachment.get("options", [])
                ],
            }
        )

    return {
        "product": payload.get("meta", {}).get("product", ""),
        "region": payload.get("meta", {}).get("region", ""),
        "time": payload.get("meta", {}).get("time"),
        "report_type": payload.get("meta", {}).get("report_type", ""),
        "introduction": payload.get("intro", {}).get("paragraphs", []),
        "survey_period": payload.get("meta", {}).get("survey_period"),
        "issued_count": payload.get("meta", {}).get("issued_count"),
        "valid_count": payload.get("meta", {}).get("valid_count"),
        "sample_size": payload.get("meta", {}).get("sample_size"),
        "data_analysis": data_analysis,
        "positive_feedback": payload.get("feedback", {}).get("positive", []),
        "negative_feedback": payload.get("feedback", {}).get("negative", []),
        "summary_recommendations": payload.get("summary", {}).get("recommendations", []),
        "attachment_questions": attachment_questions,
        "checks": payload.get("checks", {}),
    }


def chart_output_dir(output_path: Path) -> Path:
    return output_path.parent / f"{output_path.stem}_charts"


def render_from_template(payload: dict, template_path: Path, manifest_path: Path, output_path: Path) -> Path:
    del template_path
    del manifest_path

    legacy_payload = v2_to_legacy_payload(payload)
    validate_payload(legacy_payload)

    charts_dir = chart_output_dir(output_path)
    charts_dir.mkdir(parents=True, exist_ok=True)
    for item in legacy_payload["data_analysis"]:
        create_chart(charts_dir / f"chart_{item['number']:02d}.png", item["title"], item["options"])

    write_docx(legacy_payload, output_path, charts_dir)
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Render a patient report body-only docx from payload v2.")
    parser.add_argument("payload_json", help="Path to report payload v2 json")
    parser.add_argument("-o", "--output", required=True, help="Output docx path")
    parser.add_argument("--template", help="Reserved legacy arg; ignored in body-only workflow")
    parser.add_argument("--manifest", help="Reserved legacy arg; ignored in body-only workflow")
    args = parser.parse_args()

    payload = load_payload(Path(args.payload_json))
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    render_from_template(payload, Path(args.template or "."), Path(args.manifest or "."), output_path)
    print(output_path)


if __name__ == "__main__":
    main()
