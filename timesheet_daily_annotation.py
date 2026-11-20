"""Create a human-annotation scaffold for Timesheet Daily Hours cells."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

import config
import pypdfium2 as pdfium
from document_processing import TesseractProvider, remove_form_lines
from document_processing.timesheet_table import extract_daily_hours_from_blocks


SAMPLE_TOTALS = {
    "Chen Yakun-202608-IAV.pdf": 156.0,
    "Liu Xing-202608-IAV.pdf": 168.0,
    "何博-2608-IAV.pdf": 156.0,
    "黄德佳-2608-IAV.pdf": 177.0,
    "轩亮亮-2608-IAV.pdf": 178.0,
}


def _page_size(path: Path, dpi: int) -> tuple[int, int]:
    document = pdfium.PdfDocument(str(path))
    try:
        page = document[0]
        bitmap = page.render(scale=float(dpi) / 72.0)
        image = bitmap.to_pil()
        size = image.size
        image.close()
        bitmap.close()
        page.close()
        return size
    finally:
        document.close()


def build_manifest(root: Path, *, dpi: int = 300, limit: int = 5) -> dict[str, Any]:
    by_name = {path.name: path for path in root.rglob("*.pdf")}
    paths = [by_name[name] for name in SAMPLE_TOTALS if name in by_name][:limit]
    records = []
    for path in paths:
        width, height = _page_size(path, dpi)
        records.append({
            "source_file": str(path),
            "source_file_name": path.name,
            "page": 1,
            "image_size": {"width": width, "height": height, "dpi": dpi},
            "expected_total_hours": SAMPLE_TOTALS[path.name],
            "daily_hours": [],
            "annotation_status": "UNLABELED",
            "notes": "Fill one row per calendar day; preserve bbox in rendered-image coordinates.",
        })
    return {
        "schema_version": 1,
        "field": "timesheet.daily_hours",
        "instructions": [
            "Annotate only numeric cells in the Daily Hours column.",
            "Include zero/leave days when the form has an explicit cell.",
            "Use rendered-image pixel coordinates and keep source page immutable.",
            "Do not infer missing values from the monthly total.",
            "Set annotation_status to REVIEWED only after a second-person check.",
        ],
        "records": records,
    }


def validate_manifest(manifest: dict[str, Any]) -> dict[str, Any]:
    """Validate annotation shape without deciding business values."""

    issues: list[str] = []
    records = manifest.get("records")
    if not isinstance(records, list) or not records:
        issues.append("records must be a non-empty list")
        records = []
    reviewed = 0
    for index, record in enumerate(records):
        if not isinstance(record, dict):
            issues.append(f"record[{index}] must be an object")
            continue
        daily = record.get("daily_hours")
        if not isinstance(daily, list):
            issues.append(f"record[{index}].daily_hours must be a list")
            continue
        dates: set[str] = set()
        for row_index, row in enumerate(daily):
            if not isinstance(row, dict):
                issues.append(f"record[{index}].daily_hours[{row_index}] must be an object")
                continue
            date = str(row.get("date", "")).strip()
            if not date:
                issues.append(f"record[{index}].daily_hours[{row_index}] missing date")
            elif date in dates:
                issues.append(f"record[{index}] duplicate date: {date}")
            dates.add(date)
            try:
                hours = float(row.get("hours"))
            except (TypeError, ValueError):
                issues.append(f"record[{index}].daily_hours[{row_index}] invalid hours")
            else:
                if not 0 <= hours <= 24:
                    issues.append(f"record[{index}].daily_hours[{row_index}] hours outside 0-24")
            bbox = row.get("bbox")
            if not isinstance(bbox, list) or len(bbox) != 4:
                issues.append(f"record[{index}].daily_hours[{row_index}] bbox must have four values")
        if str(record.get("annotation_status", "")).upper() == "REVIEWED":
            reviewed += 1
    return {
        "status": "READY" if records and reviewed == len(records) and not issues else "INCOMPLETE",
        "record_count": len(records),
        "reviewed_records": reviewed,
        "issues": issues,
    }


def apply_ocr_proposals(manifest: dict[str, Any], root: Path, *, dpi: int = 300) -> dict[str, Any]:
    """Attach OCR suggestions without promoting them to reviewed annotations."""

    by_name = {path.name: path for path in root.rglob("*.pdf")}
    provider = TesseractProvider(config.TESSERACT_EXE, languages=config.OCR_LANGUAGE, timeout=90)
    for record in manifest.get("records", []):
        source_name = str(record.get("source_file_name", ""))
        if not source_name and record.get("source_file"):
            source_name = Path(str(record["source_file"])).name
        path = by_name.get(source_name)
        if path is None:
            record["suggestion_diagnostics"] = {"status": "SOURCE_NOT_FOUND"}
            record["suggested_daily_hours"] = []
            continue
        document = pdfium.PdfDocument(str(path))
        try:
            page = document[0]
            bitmap = page.render(scale=float(dpi) / 72.0)
            image = bitmap.to_pil()
            cleaned = remove_form_lines(image)
            image.close()
            bitmap.close()
            page.close()
        finally:
            document.close()
        try:
            result = provider.recognize(cleaned, page=1, psm="6", output_mode="tsv")
            blocks = []
            for block in result.blocks:
                if block.bbox is None:
                    continue
                left, top, right, bottom = block.bbox
                blocks.append({
                    "text": block.text, "left": left, "top": top,
                    "width": right - left, "height": bottom - top,
                    "line": list(block.line_id or ()),
                })
            cells, diagnostics = extract_daily_hours_from_blocks(
                blocks, page_height=cleaned.height,
            )
            record["suggested_daily_hours"] = cells
            record["suggestion_diagnostics"] = diagnostics
            record["suggestion_ocr_confidence"] = result.confidence
        finally:
            cleaned.close()
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(description="Create a Timesheet Daily Hours annotation manifest")
    parser.add_argument("--root", default=str(Path(config.ROOT) / "CNTS-25165" / "01_IAV TMS & WP"))
    parser.add_argument("--dpi", type=int, default=300)
    parser.add_argument("--limit", type=int, default=5)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--validate", type=Path, help="validate an existing annotation manifest")
    parser.add_argument("--propose", action="store_true", help="attach OCR suggestions to a new manifest")
    args = parser.parse_args()
    if args.validate:
        manifest = json.loads(args.validate.read_text(encoding="utf-8"))
        result = validate_manifest(manifest)
        print(json.dumps(result, ensure_ascii=True, indent=2))
        return 0 if result["status"] == "READY" else 2
    manifest = build_manifest(Path(args.root), dpi=args.dpi, limit=args.limit)
    if args.propose:
        manifest = apply_ocr_proposals(manifest, Path(args.root), dpi=args.dpi)
    payload = json.dumps(manifest, ensure_ascii=True, indent=2)
    if args.output:
        args.output.write_text(payload + "\n", encoding="utf-8")
    else:
        print(payload)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
