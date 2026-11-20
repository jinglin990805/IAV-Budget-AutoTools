"""Benchmark conservative Daily Hours row/column association on real scans."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

import config
import pypdfium2 as pdfium

from document_processing import TesseractProvider, remove_form_lines
from document_processing.timesheet_table import extract_daily_hours_from_blocks


SAMPLE_NAMES = (
    "Chen Yakun-202608-IAV.pdf",
    "Liu Xing-202608-IAV.pdf",
    "何博-2608-IAV.pdf",
    "黄德佳-2608-IAV.pdf",
    "轩亮亮-2608-IAV.pdf",
)

SAMPLE_TOTALS = {
    "Chen Yakun-202608-IAV.pdf": 156.0,
    "Liu Xing-202608-IAV.pdf": 168.0,
    "何博-2608-IAV.pdf": 156.0,
    "黄德佳-2608-IAV.pdf": 177.0,
    "轩亮亮-2608-IAV.pdf": 178.0,
}


def _first_page(path: Path, dpi: int) -> tuple[Any, tuple[int, int]]:
    document = pdfium.PdfDocument(str(path))
    try:
        page = document[0]
        bitmap = page.render(scale=float(dpi) / 72.0)
        image = bitmap.to_pil()
        size = image.size
        bitmap.close()
        page.close()
        return image, size
    finally:
        document.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="Benchmark Timesheet Daily Hours association")
    parser.add_argument("--root", default=str(Path(config.ROOT) / "CNTS-25165" / "01_IAV TMS & WP"))
    parser.add_argument("--dpi", type=int, default=300)
    parser.add_argument("--limit", type=int, default=5)
    parser.add_argument("--clean-lines", action="store_true", help="remove long form lines before TSV OCR")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    by_name = {path.name: path for path in Path(args.root).rglob("*.pdf")}
    paths = [by_name[name] for name in SAMPLE_NAMES if name in by_name][:args.limit]
    provider = TesseractProvider(config.TESSERACT_EXE, languages=config.OCR_LANGUAGE, timeout=90)
    records: list[dict[str, Any]] = []
    for path in paths:
        image, size = _first_page(path, args.dpi)
        try:
            if args.clean_lines:
                cleaned = remove_form_lines(image)
                image.close()
                image = cleaned
            result = provider.recognize(image, page=1, psm="6", output_mode="tsv")
            blocks = []
            for block in result.blocks:
                if block.bbox is None:
                    continue
                left, top, right, bottom = block.bbox
                blocks.append({
                    "text": block.text,
                    "left": left,
                    "top": top,
                    "width": right - left,
                    "height": bottom - top,
                    "line": list(block.line_id or ()),
                })
            cells, diagnostics = extract_daily_hours_from_blocks(blocks, page_height=size[1])
            expected_total = SAMPLE_TOTALS.get(path.name)
            daily_sum = round(sum(float(cell["hours"]) for cell in cells), 2)
            diagnostics["daily_sum"] = daily_sum
            diagnostics["expected_total_hours"] = expected_total
            if (
                diagnostics["status"] == "OK"
                and expected_total is not None
                and abs(daily_sum - expected_total) > 0.01
            ):
                diagnostics["status"] = "SUM_MISMATCH"
                diagnostics["reason"] = "daily cell sum differs from annotated Total Hours"
            records.append({
                "source_file": str(path),
                "status": diagnostics["status"],
                "cells": cells,
                "diagnostics": diagnostics,
                "ocr_characters": len(result.text),
                "ocr_confidence": result.confidence,
                "warnings": list(result.warnings),
            })
        finally:
            image.close()

    summary = {
        "sample_count": len(records),
        "stable_pages": sum(item["status"] == "OK" for item in records),
        "records": records,
    }
    if args.json:
        print(json.dumps(summary, ensure_ascii=True, indent=2))
    else:
        print(f"stable_pages={summary['stable_pages']}/{summary['sample_count']}")
        for item in records:
            print(Path(item["source_file"]).name, item["status"], item["diagnostics"].get("cells", 0))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
