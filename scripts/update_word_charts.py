#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import shutil
import string
import tempfile
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree
from openpyxl import load_workbook

CHART_NS = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
REL_NS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
CHART_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
PACKAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"


def load_payload(path: Path) -> dict:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if "chapter2" not in payload or "template" not in payload:
        raise ValueError("Expected template-driven payload v2.")
    return payload


def chart_targets(zip_file: ZipFile) -> list[str]:
    rels_root = etree.fromstring(zip_file.read("word/_rels/document.xml.rels"))
    targets = []
    for relationship in rels_root.findall("r:Relationship", REL_NS):
        if relationship.get("Type") == CHART_REL_TYPE:
            targets.append("word/" + relationship.get("Target").lstrip("/"))
    return sorted(targets, key=lambda value: int(Path(value).stem.replace("chart", "")))


def workbook_target(zip_file: ZipFile, chart_path: str) -> str:
    rel_path = f"word/charts/_rels/{Path(chart_path).name}.rels"
    root = etree.fromstring(zip_file.read(rel_path))
    for relationship in root.findall("r:Relationship", REL_NS):
        if relationship.get("Type") == PACKAGE_REL_TYPE:
            target = relationship.get("Target")
            return str(Path("word/charts") / target).replace("word/charts/../", "word/")
    raise ValueError(f"Workbook relationship missing for {chart_path}.")


def excel_column(index: int) -> str:
    return string.ascii_uppercase[index]


def update_workbook_bytes(source: bytes, categories: list[str], values: list[float]) -> bytes:
    workbook = load_workbook(BytesIO(source))
    worksheet = workbook[workbook.sheetnames[0]]
    worksheet.delete_rows(1, worksheet.max_row)
    worksheet["A1"] = "选项"
    worksheet["A2"] = "占比"
    for offset, (category, value) in enumerate(zip(categories, values), start=2):
        worksheet.cell(row=1, column=offset, value=category)
        worksheet.cell(row=2, column=offset, value=value / 100 if value > 1 else value)
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def set_cache_points(cache_node, values: list[str]) -> None:
    for point in list(cache_node.findall("c:pt", CHART_NS)):
        cache_node.remove(point)
    point_count = cache_node.find("c:ptCount", CHART_NS)
    if point_count is None:
        point_count = etree.SubElement(cache_node, f"{{{CHART_NS['c']}}}ptCount")
    point_count.set("val", str(len(values)))
    for index, value in enumerate(values):
        point = etree.SubElement(cache_node, f"{{{CHART_NS['c']}}}pt")
        point.set("idx", str(index))
        v = etree.SubElement(point, f"{{{CHART_NS['c']}}}v")
        v.text = value


def update_chart_xml(xml_bytes: bytes, categories: list[str], values: list[float]) -> bytes:
    root = etree.fromstring(xml_bytes)
    bar_chart = root.find(".//c:barChart", CHART_NS)
    if bar_chart is None:
        raise ValueError("Only bar charts are supported in v1.")
    series_nodes = bar_chart.findall("c:ser", CHART_NS)
    if not series_nodes:
        raise ValueError("Chart has no series nodes.")
    template_series = series_nodes[0]
    for node in series_nodes[1:]:
        bar_chart.remove(node)

    for idx, (category, value) in enumerate(zip(categories, values)):
        if idx == 0:
            series = template_series
        else:
            series = etree.fromstring(etree.tostring(template_series))
            bar_chart.append(series)
        series.find("c:idx", CHART_NS).set("val", str(idx))
        series.find("c:order", CHART_NS).set("val", str(idx))

        tx = series.find("c:tx/c:strRef", CHART_NS)
        tx.find("c:f", CHART_NS).text = f"Sheet1!${excel_column(idx + 1)}$1"
        set_cache_points(tx.find("c:strCache", CHART_NS), [category])

        cat = series.find("c:cat/c:strRef", CHART_NS)
        cat.find("c:f", CHART_NS).text = "Sheet1!$A$2"
        set_cache_points(cat.find("c:strCache", CHART_NS), ["占比"])

        val_ref = series.find("c:val/c:numRef", CHART_NS)
        val_ref.find("c:f", CHART_NS).text = f"Sheet1!${excel_column(idx + 1)}$2"
        set_cache_points(val_ref.find("c:numCache", CHART_NS), [str(value / 100 if value > 1 else value)])

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def update_docx_charts(docx_path: Path, payload: dict, output_path: Path | None = None) -> Path:
    final_path = output_path or docx_path
    if output_path and output_path != docx_path:
        shutil.copyfile(docx_path, output_path)
    working_path = final_path

    with ZipFile(working_path) as source_zip:
        chart_paths = chart_targets(source_zip)
        replacements: dict[str, bytes] = {}
        items = payload["chapter2"]["items"]
        if len(items) > len(chart_paths):
            raise ValueError(f"Template provides {len(chart_paths)} charts, got {len(items)} chapter2 items.")
        for chart_path, item in zip(chart_paths, items):
            categories = item["chart"]["categories"]
            values = item["chart"]["series"][0]["values"]
            workbook_path = workbook_target(source_zip, chart_path)
            replacements[workbook_path] = update_workbook_bytes(source_zip.read(workbook_path), categories, values)
            replacements[chart_path] = update_chart_xml(source_zip.read(chart_path), categories, values)

        temp_fd, temp_name = tempfile.mkstemp(suffix=".docx")
        Path(temp_name).unlink(missing_ok=True)
        with ZipFile(temp_name, "w", compression=ZIP_DEFLATED) as target_zip:
            for entry in source_zip.infolist():
                data = replacements.get(entry.filename, source_zip.read(entry.filename))
                target_zip.writestr(entry, data)
    shutil.move(temp_name, working_path)
    return working_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Update Word native editable charts in a rendered docx report.")
    parser.add_argument("docx_file", help="Path to rendered docx file")
    parser.add_argument("payload_json", help="Path to report payload v2 json")
    parser.add_argument("-o", "--output", help="Optional output docx path")
    args = parser.parse_args()

    payload = load_payload(Path(args.payload_json))
    result = update_docx_charts(Path(args.docx_file), payload, Path(args.output) if args.output else None)
    print(result)


if __name__ == "__main__":
    main()
