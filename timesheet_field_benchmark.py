"""Field-level Timesheet OCR benchmark for the annotated 2026-08 sample."""

from __future__ import annotations

import argparse
import json
import re
import time
from pathlib import Path
from typing import Any, Callable

import config
import pypdfium2 as pdfium
from document_processing import PaddleOCRProvider, ProviderUnavailable, TesseractProvider
try:
    from timesheet_ingest import _extract_hours, _extract_period, _extract_po
except ModuleNotFoundError:
    # The isolated Paddle environment intentionally contains only OCR
    # dependencies. Keep this benchmark independent of the production PDF stack.
    def _extract_period(text: str, path: Path) -> tuple[int | None, int | None]:
        match = re.search(
            r"(?:Month\s*(?:[/|lI1]\s*)?Year)\s*[:：]?\s*"
            r"(0?[1-9]|1[0-2])\s*[/|]\s*(20\d{2})",
            text or "",
            re.I,
        )
        if match:
            return int(match.group(2)), int(match.group(1))
        match = re.search(r"\b(20\d{2})[-_/ ](0?[1-9]|1[0-2])\b", text or "")
        return (int(match.group(1)), int(match.group(2))) if match else (None, None)

    def _extract_po(text: str, path: Path) -> str:
        match = re.search(r"(?:Purchase\s*Order|PO)[^0-9]{0,40}(\d{8,})", text or "", re.I)
        return match.group(1) if match else ""

    def _extract_hours(text: str) -> float | None:
        values = []
        for match in re.finditer(r"(?:ACTUAL_WORK_HOUR_VALUE|Actual\s+Work[- ]?hour[^\d]{0,80})(\d{1,4}(?:\.\d+)?)", text or "", re.I):
            try:
                value = float(match.group(1))
                if 0 <= value <= 400:
                    values.append(value)
            except ValueError:
                pass
        return max(values) if values else None


FALLBACK_GROUND_TRUTH = {
    "Chen Yakun-202608-IAV.pdf": {"period": "2026-08", "po_id": "2607000312", "hours": 156.0},
    "Liu Xing-202608-IAV.pdf": {"period": "2026-08", "po_id": "2607000312", "hours": 168.0},
    "何博-2608-IAV.pdf": {"period": "2026-08", "po_id": "2607000312", "hours": 156.0},
    "黄德佳-2608-IAV.pdf": {"period": "2026-08", "po_id": "2607000312", "hours": 177.0},
    "轩亮亮-2608-IAV.pdf": {"period": "2026-08", "po_id": "2607000312", "hours": 178.0},
}

# Populated by ``main`` from the fact store; the literal above is only the
# offline fallback for an environment without a database.
GROUND_TRUTH: dict[str, dict[str, Any]] = FALLBACK_GROUND_TRUTH


def load_ground_truth(
    month: str | None = None,
    project: str | None = None,
    db_path: str | Path | None = None,
) -> dict[str, dict[str, Any]]:
    """Rebuild the expected-field map from approved facts in the fact store.

    Scoring the parser against a hand-maintained literal meant every reviewer
    correction had to be copied in by hand, so the corrections that actually
    mattered never became regression cases.  Each approved Timesheet fact is
    joined back to the PDF it came from, so what a reviewer confirmed *is* the
    expectation from then on.
    """
    import sqlite3

    candidates = [
        Path(db_path) if db_path else Path(__file__).resolve().parent / "data" / "iav_facts.sqlite3",
        Path(config.ROOT) / "data" / "iav_facts.sqlite3",
    ]
    path = next((item for item in candidates if item.exists()), None)
    if path is None:
        return {}

    query = """
        SELECT s.file_name, f.month, f.po_id, f.hours, f.version
          FROM timesheet_facts f
          JOIN source_files s ON s.source_id = f.source_id
         WHERE f.approval_status = 'APPROVED'
           AND f.is_current = 1
           AND f.source_name = 'timesheet_pdf'
    """
    params: list[Any] = []
    if month:
        query += " AND f.month = ?"
        params.append(month)
    if project:
        query += " AND f.project_id = ?"
        params.append(project)

    try:
        with sqlite3.connect(path) as connection:
            connection.row_factory = sqlite3.Row
            rows = [dict(row) for row in connection.execute(query, params)]
    except sqlite3.Error:
        return {}

    truth: dict[str, dict[str, Any]] = {}
    for row in rows:
        name = str(row["file_name"] or "")
        if not name:
            continue
        current = truth.get(name)
        if current is not None and int(current.get("version", 0)) >= int(row["version"] or 0):
            continue
        truth[name] = {
            "period": str(row["month"] or ""),
            "po_id": str(row["po_id"] or ""),
            "hours": float(row["hours"]),
            "version": int(row["version"] or 0),
        }
    return truth


def _sample_paths(root: Path) -> list[Path]:
    by_name = {path.name: path for path in root.rglob("*.pdf")}
    return [by_name[name] for name in GROUND_TRUTH if name in by_name]


def _first_page(path: Path, dpi: int) -> Any:
    document = pdfium.PdfDocument(str(path))
    try:
        page = document[0]
        bitmap = page.render(scale=float(dpi) / 72.0)
        image = bitmap.to_pil()
        bitmap.close()
        page.close()
        return image
    finally:
        document.close()


def _extract(text: str) -> dict[str, Any]:
    year, month = _extract_period(text, Path("scan.pdf"))
    hours = _extract_hours(text)
    return {
        "period": f"{year:04d}-{month:02d}" if year and month else "",
        "po_id": _extract_po(text, Path("scan.pdf")),
        "hours": hours,
    }


def _run_production(paths: list[Path]) -> dict[str, Any]:
    """Measure the real Timesheet parser separately from raw OCR text.

    The raw provider benchmark intentionally bypasses filename recovery and the
    layout-aware footer extractor.  This route measures the production contract
    that feeds candidates/review, so the two results cannot be mistaken for one
    another.
    """
    try:
        from timesheet_ingest import parse_pdf
    except (ImportError, ModuleNotFoundError) as exc:
        return {
            "engine": "production_parser",
            "status": "UNAVAILABLE",
            "reason": str(exc),
            "records": [],
        }

    started = time.perf_counter()
    records: list[dict[str, Any]] = []
    for path in paths:
        page_started = time.perf_counter()
        parsed = parse_pdf(path, use_ocr=True)
        actual = {
            "period": (
                f"{parsed.year:04d}-{parsed.month:02d}"
                if parsed.year and parsed.month else ""
            ),
            "po_id": parsed.po_id,
            "hours": parsed.total_hours,
        }
        expected = GROUND_TRUTH[path.name]
        matches = {
            field: actual[field] == expected[field]
            for field in ("period", "po_id", "hours")
        }
        hours_evidence = next(
            (item for item in parsed.evidence if item.get("field") == "hours"),
            {},
        )
        records.append({
            "source_file": str(path),
            "actual": actual,
            "expected": expected,
            "matches": matches,
            "all_fields_exact": all(matches.values()),
            "parse_method": parsed.parse_method,
            "validation_status": parsed.validation.get("status"),
            "monthly_total_diagnostics": parsed.ocr_diagnostics.get("monthly_total", {}),
            "hours_evidence_bbox": hours_evidence.get("bbox"),
            "elapsed_ms": round((time.perf_counter() - page_started) * 1000, 1),
        })

    totals = {
        field: sum(record["matches"][field] for record in records)
        for field in ("period", "po_id", "hours")
    }
    return {
        "engine": "production_parser",
        "status": "PASS",
        "records": records,
        "field_exact_counts": totals,
        "all_fields_exact": sum(record["all_fields_exact"] for record in records),
        "sample_count": len(records),
        "elapsed_ms": round((time.perf_counter() - started) * 1000, 1),
    }


def _run_engine(
    name: str,
    paths: list[Path],
    provider_factory: Callable[[], Any],
    dpi: int,
) -> dict[str, Any]:
    started = time.perf_counter()
    try:
        provider = provider_factory()
    except ProviderUnavailable as exc:
        return {"engine": name, "status": "UNAVAILABLE", "reason": str(exc), "records": []}
    records: list[dict[str, Any]] = []
    for path in paths:
        image = _first_page(path, dpi)
        page_started = time.perf_counter()
        try:
            result = provider.recognize(image, page=1, psm="6")
            actual = _extract(result.text)
            expected = GROUND_TRUTH[path.name]
            matches = {
                "period": actual["period"] == expected["period"],
                "po_id": actual["po_id"] == expected["po_id"],
                "hours": actual["hours"] == expected["hours"],
            }
            records.append({
                "source_file": str(path),
                "actual": actual,
                "expected": expected,
                "matches": matches,
                "all_fields_exact": all(matches.values()),
                "character_count": len(result.text.strip()),
                "confidence": result.confidence,
                "warnings": list(result.warnings),
                "elapsed_ms": round((time.perf_counter() - page_started) * 1000, 1),
            })
        finally:
            image.close()
    totals = {
        field: sum(record["matches"][field] for record in records)
        for field in ("period", "po_id", "hours")
    }
    return {
        "engine": name,
        "status": "PASS",
        "records": records,
        "field_exact_counts": totals,
        "all_fields_exact": sum(record["all_fields_exact"] for record in records),
        "sample_count": len(records),
        "elapsed_ms": round((time.perf_counter() - started) * 1000, 1),
    }


def main() -> int:
    global GROUND_TRUTH

    parser = argparse.ArgumentParser(description="Benchmark Timesheet field exact match")
    parser.add_argument("--root", default=str(Path(config.ROOT) / "CNTS-25165" / "01_IAV TMS & WP"))
    parser.add_argument("--dpi", type=int, default=150)
    parser.add_argument("--limit", type=int, default=5, help="0 runs the whole corpus")
    parser.add_argument("--engine", choices=("tesseract", "paddle", "both"), default="both")
    parser.add_argument(
        "--production",
        action="store_true",
        help="also run the real parse_pdf production path, including footer/layout recovery",
    )
    parser.add_argument("--month", default=None, help="score only this period, e.g. 2026-09")
    parser.add_argument("--project", default=None, help="restrict facts to one project id")
    parser.add_argument("--facts-db", default=None, help="path to the SQLite fact store")
    parser.add_argument(
        "--fallback-literal",
        action="store_true",
        help="score against the built-in five-file literal instead of approved facts",
    )
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    if args.fallback_literal:
        GROUND_TRUTH = FALLBACK_GROUND_TRUTH
        expectations = "builtin_literal"
    else:
        loaded = load_ground_truth(args.month, args.project, args.facts_db)
        if loaded:
            GROUND_TRUTH = loaded
            expectations = f"approved_facts({len(loaded)} files)"
        else:
            GROUND_TRUTH = FALLBACK_GROUND_TRUTH
            expectations = "builtin_literal (no facts found)"

    paths = _sample_paths(Path(args.root))
    if args.limit > 0:
        paths = paths[: args.limit]
    result: dict[str, Any] = {
        "root": args.root,
        "expectations": expectations,
        "sample_files": [str(path) for path in paths],
        "sample_expected": GROUND_TRUTH,
    }
    if not args.json:
        print(f"expectations: {expectations}")
        print(f"sample files: {len(paths)}")
    if args.engine in {"tesseract", "both"}:
        result["tesseract"] = _run_engine(
            "tesseract", paths,
            lambda: TesseractProvider(config.TESSERACT_EXE, languages=config.OCR_LANGUAGE, timeout=90),
            args.dpi,
        )
    if args.engine in {"paddle", "both"}:
        result["paddle"] = _run_engine(
            "paddleocr", paths,
            lambda: PaddleOCRProvider(lang=config.PADDLE_LANG, device=config.PADDLE_DEVICE),
            args.dpi,
        )
    if args.production:
        result["production"] = _run_production(paths)
    if args.json:
        print(json.dumps(result, ensure_ascii=True, indent=2))
    else:
        for engine in ("tesseract", "paddle"):
            if engine not in result:
                continue
            item = result[engine]
            print(engine, item["status"], item.get("field_exact_counts"), item.get("all_fields_exact"), item.get("elapsed_ms"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
