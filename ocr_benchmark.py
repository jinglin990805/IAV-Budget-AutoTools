"""Small, repeatable OCR provider benchmark.

The harness intentionally does not treat an unavailable provider as a score of
zero.  It reports availability separately so a missing optional dependency cannot
make a routing decision look like an accuracy result.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import time
from pathlib import Path
from typing import Any, Iterable

import config
from document_processing import PaddleOCRProvider, ProviderUnavailable, TesseractProvider


def _pdf_paths(inputs: Iterable[str], limit: int | None = None) -> list[Path]:
    paths: list[Path] = []
    for raw in inputs:
        path = Path(raw)
        if path.is_file() and path.suffix.casefold() == ".pdf":
            paths.append(path)
        elif path.is_dir():
            paths.extend(sorted(path.rglob("*.pdf"), key=lambda item: str(item).casefold()))
    unique: list[Path] = []
    seen: set[str] = set()
    for path in paths:
        key = str(path.resolve()).casefold()
        if key not in seen:
            seen.add(key)
            unique.append(path)
    return unique[:limit] if limit is not None else unique


def _render_pages(path: Path, dpi: int, max_pages: int | None) -> Iterable[tuple[int, Any]]:
    import pypdfium2 as pdfium

    document = pdfium.PdfDocument(str(path))
    try:
        count = len(document) if max_pages is None else min(len(document), max_pages)
        for index in range(count):
            page = document[index]
            bitmap = page.render(scale=float(dpi) / 72.0)
            image = bitmap.to_pil()
            bitmap.close()
            page.close()
            yield index + 1, image
    finally:
        document.close()


def benchmark_tesseract(
    paths: Iterable[Path], *, dpi: int = 150, max_pages: int | None = 1,
) -> dict[str, Any]:
    executable = Path(getattr(config, "TESSERACT_EXE", ""))
    provider = TesseractProvider(executable, languages=config.OCR_LANGUAGE, timeout=90)
    started = time.perf_counter()
    pages: list[dict[str, Any]] = []
    if not provider.available:
        return {
            "status": "UNAVAILABLE",
            "engine": provider.name,
            "reason": f"tesseract_not_found:{executable}",
            "pages": [],
            "elapsed_ms": round((time.perf_counter() - started) * 1000, 1),
        }
    for path in paths:
        for page_number, image in _render_pages(path, dpi, max_pages):
            page_started = time.perf_counter()
            try:
                result = provider.recognize(image, page=page_number, output_mode="text", psm="6")
                pages.append({
                    "source_file": str(path),
                    "page": page_number,
                    "ok": result.ok,
                    "character_count": len(result.text.strip()),
                    "confidence": result.confidence,
                    "warnings": list(result.warnings),
                    "elapsed_ms": round((time.perf_counter() - page_started) * 1000, 1),
                })
            finally:
                image.close()
    return {
        "status": "PASS",
        "engine": provider.name,
        "dpi": dpi,
        "pages": pages,
        "page_recall": round(sum(item["ok"] for item in pages) / len(pages), 4) if pages else None,
        "elapsed_ms": round((time.perf_counter() - started) * 1000, 1),
    }


def benchmark_paddle(
    paths: Iterable[Path], *, dpi: int = 150, max_pages: int | None = 1,
) -> dict[str, Any]:
    """Run Paddle only when both packages and the optional adapter are usable."""
    started = time.perf_counter()
    try:
        provider = PaddleOCRProvider(
            lang=getattr(config, "PADDLE_LANG", "ch"),
            device=getattr(config, "PADDLE_DEVICE", "cpu"),
        )
    except ProviderUnavailable as exc:
        return {
            "status": "UNAVAILABLE",
            "engine": "paddleocr",
            "reason": str(exc),
            "pages": [],
            "elapsed_ms": round((time.perf_counter() - started) * 1000, 1),
        }
    pages: list[dict[str, Any]] = []
    for path in paths:
        for page_number, image in _render_pages(path, dpi, max_pages):
            page_started = time.perf_counter()
            try:
                result = provider.recognize(image, page=page_number)
                pages.append({
                    "source_file": str(path),
                    "page": page_number,
                    "ok": result.ok,
                    "character_count": len(result.text.strip()),
                    "confidence": result.confidence,
                    "warnings": list(result.warnings),
                    "elapsed_ms": round((time.perf_counter() - page_started) * 1000, 1),
                })
            finally:
                image.close()
    return {
        "status": "PASS",
        "engine": provider.name,
        "dpi": dpi,
        "pages": pages,
        "page_recall": round(sum(item["ok"] for item in pages) / len(pages), 4) if pages else None,
        "elapsed_ms": round((time.perf_counter() - started) * 1000, 1),
    }


def provider_availability() -> dict[str, Any]:
    result: dict[str, Any] = {}
    for name in ("paddle", "paddleocr", "cv2"):
        result[name] = bool(importlib.util.find_spec(name))
    if not result["paddle"] or not result["paddleocr"]:
        result["paddle_status"] = "UNAVAILABLE"
        result["paddle_reason"] = "paddle and/or paddleocr is not installed in the active environment"
    else:
        result["paddle_status"] = "AVAILABLE_FOR_ADAPTER_CHECK"
    return result


def run_benchmark(inputs: Iterable[str], *, limit: int | None = None, dpi: int = 150, max_pages: int | None = 1) -> dict[str, Any]:
    paths = _pdf_paths(inputs, limit)
    availability = provider_availability()
    paddle = benchmark_paddle(paths, dpi=dpi, max_pages=max_pages) if availability["paddle_status"] != "UNAVAILABLE" else {
        "status": "UNAVAILABLE",
        "engine": "paddleocr",
        "reason": availability.get("paddle_reason", ""),
        "pages": [],
    }
    comparison_status = "COMPARABLE" if paddle.get("status") == "PASS" else "PADDLE_UNAVAILABLE"
    return {
        "inputs": list(inputs),
        "files": [str(path) for path in paths],
        "provider_availability": availability,
        "tesseract": benchmark_tesseract(paths, dpi=dpi, max_pages=max_pages),
        "paddle": paddle,
        "comparison_status": comparison_status,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Benchmark installed OCR providers without changing production routing")
    parser.add_argument("inputs", nargs="*", default=[".playwright-mcp/ocr-fixtures"])
    parser.add_argument("--limit", type=int, default=4)
    parser.add_argument("--dpi", type=int, default=150)
    parser.add_argument("--max-pages", type=int, default=1)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    result = run_benchmark(args.inputs, limit=args.limit, dpi=args.dpi, max_pages=args.max_pages)
    if args.json:
        # Windows console code pages may not be UTF-8; escaped JSON remains
        # lossless and can be redirected safely to a report file.
        print(json.dumps(result, ensure_ascii=True, indent=2))
    else:
        print(f"files={len(result['files'])} paddle={result['paddle']['status']} tesseract={result['tesseract']['status']}")
        print(f"tesseract_page_recall={result['tesseract'].get('page_recall')} elapsed_ms={result['tesseract'].get('elapsed_ms')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
