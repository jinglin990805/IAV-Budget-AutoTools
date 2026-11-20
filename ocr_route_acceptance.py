"""Read-only acceptance scan for the OCR production routes.

This command exercises existing Timesheet and Travel OCR paths on real files.
It never writes candidates, facts, Excel files, or business decisions.
"""

from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path
from typing import Any, Iterable

import config


# These are OCR field-level regression anchors for the two historical Travel
# packages whose old notes were one item short.  They describe source-page
# evidence only; they are not travel-policy or payable-value rules.
_TRAVEL_FIELD_BASELINES: dict[str, dict[str, Any]] = {
    "Chengdu": {
        "expense_count": 18,
        "required_page_amounts": [
            (1, 1360.00), (2, 1555.00), (4, 379.00), (6, 449.00),
            (8, 1569.00), (10, 5.00), (11, 43.20), (13, 119.40),
            (15, 51.70), (17, 19.10), (18, 14.00), (18, 21.00),
            (19, 56.70), (21, 167.50), (22, 100.00), (23, 131.00),
            (24, 131.00), (25, 13.05),
        ],
        "forbidden_amounts": [99.88],
    },
    "Suzhou": {
        "expense_count": 11,
        "required_page_amounts": [
            (1, 532.00), (2, 117.00), (3, 1077.00), (5, 399.00),
            (9, 123.70), (11, 23.90), (13, 61.70), (15, 34.20),
            (17, 135.40), (18, 104.04), (19, 14.84),
        ],
        "forbidden_amounts": [],
    },
}


def _pdf_paths(
    root: Path,
    limit: int | None = None,
    sample_per_directory: int | None = None,
    offset: int = 0,
) -> list[Path]:
    paths = sorted(root.rglob("*.pdf"), key=lambda item: str(item).casefold())
    if sample_per_directory is not None:
        sample_size = max(1, int(sample_per_directory))
        selected: list[Path] = []
        counts: Counter[str] = Counter()
        for path in paths:
            directory = str(path.parent.resolve()).casefold()
            if counts[directory] >= sample_size:
                continue
            counts[directory] += 1
            selected.append(path)
        paths = selected
    start = max(0, int(offset))
    if limit is None:
        return paths[start:]
    return paths[start:start + max(0, int(limit))]


def _timesheet_acceptance(
    root: Path,
    limit: int | None = None,
    sample_per_directory: int | None = None,
    offset: int = 0,
) -> dict[str, Any]:
    from timesheet_ingest import parse_pdf

    paths = _pdf_paths(root, limit, sample_per_directory, offset)
    methods: Counter[str] = Counter()
    validation: Counter[str] = Counter()
    missing_fields: Counter[str] = Counter()
    retry_count = 0
    selected_retry_count = 0
    errors: list[dict[str, str]] = []
    review_records: list[dict[str, Any]] = []

    for path in paths:
        try:
            parsed = parse_pdf(path, use_ocr=True)
            methods[parsed.parse_method] += 1
            validation[str(parsed.validation.get("status", "UNKNOWN"))] += 1
            for field, value in (parsed.fields or {}).items():
                if not value:
                    missing_fields[field] += 1
            page_rows = parsed.ocr_diagnostics.get("pages", [])
            attempts = [item for page in page_rows for item in page.get("attempts", [])]
            if len(attempts) > 1:
                retry_count += 1
            enhanced_selected = any(
                item.get("selected") == "tesseract_enhanced_psm11" for item in page_rows
            )
            if enhanced_selected:
                selected_retry_count += 1
            if parsed.ocr_error:
                errors.append({"file": str(path), "error": parsed.ocr_error})
            missing = [field for field, value in (parsed.fields or {}).items() if not value]
            status = str(parsed.validation.get("status", "UNKNOWN"))
            issue_codes = [
                str(issue.get("code", ""))
                for issue in (parsed.validation.get("issues") or [])
                if issue.get("code")
            ]
            ocr_issue_codes = [code for code in issue_codes if code.startswith("OCR_")]
            business_issue_codes = [code for code in issue_codes if not code.startswith("OCR_")]
            ocr_completeness = parsed.validation.get("checks", {}).get("ocr_completeness", {})
            if status != "PASS" or missing or parsed.ocr_error:
                review_records.append({
                    "file": str(path),
                    "status": status,
                    "parse_method": parsed.parse_method,
                    "missing_fields": missing,
                    "field_complete": not missing,
                    "ocr_completeness_status": ocr_completeness.get("status", "NOT_APPLICABLE"),
                    "ocr_issue_codes": ocr_issue_codes,
                    "business_issue_codes": business_issue_codes,
                    "retry_triggered": len(attempts) > 1,
                    "enhanced_retry_selected": enhanced_selected,
                    "ocr_error": parsed.ocr_error or "",
                })
        except Exception as exc:  # keep the acceptance scan going
            errors.append({"file": str(path), "error": f"{type(exc).__name__}: {exc}"})

    return {
        "root": str(root),
        "file_count": len(paths),
        "offset": max(0, int(offset)),
        "sample_per_directory": sample_per_directory,
        "parse_methods": dict(methods),
        "validation_status": dict(validation),
        "missing_fields": dict(missing_fields),
        "retry_triggered_files": retry_count,
        "enhanced_retry_selected_files": selected_retry_count,
        "error_count": len(errors),
        "errors": errors[:20],
        "review_records": review_records,
    }


def _travel_acceptance(paths: Iterable[Path]) -> dict[str, Any]:
    from travel_invoice_reviewer import _read_pdf

    records: list[dict[str, Any]] = []
    for path in paths:
        try:
            text, warnings = _read_pdf(str(path))
            pages = [page for page in text.split("\n\f\n") if page.strip()]
            records.append({
                "file": str(path),
                "non_empty_pages": len(pages),
                "character_count": len(text.strip()),
                "warnings": warnings,
                "status": "PASS" if text.strip() else "REVIEW_REQUIRED",
            })
        except Exception as exc:
            records.append({
                "file": str(path),
                "non_empty_pages": 0,
                "character_count": 0,
                "warnings": [f"{type(exc).__name__}: {exc}"],
                "status": "REVIEW_REQUIRED",
            })
    return {
        "file_count": len(records),
        "pass_count": sum(item["status"] == "PASS" for item in records),
        "review_required_count": sum(item["status"] != "PASS" for item in records),
        "records": records,
    }


def _travel_field_assertions(path: Path, trip: Any) -> dict[str, Any]:
    """Check stable page/amount OCR anchors without applying business rules."""

    baseline_name = next(
        (name for name in _TRAVEL_FIELD_BASELINES if name.casefold() in path.name.casefold()),
        None,
    )
    if baseline_name is None:
        return {"status": "NOT_APPLICABLE", "baseline": None}

    baseline = _TRAVEL_FIELD_BASELINES[baseline_name]
    actual = Counter(
        (int(getattr(expense, "source_page", 0) or 0), round(float(expense.amount), 2))
        for expense in (getattr(trip, "expenses", []) or [])
    )
    required = Counter((int(page), round(float(amount), 2)) for page, amount in baseline["required_page_amounts"])
    missing = list((required - actual).elements())
    unexpected_forbidden = [
        round(float(amount), 2)
        for expense in (getattr(trip, "expenses", []) or [])
        for amount in baseline["forbidden_amounts"]
        if abs(float(expense.amount) - float(amount)) <= 0.01
    ]
    count_ok = len(getattr(trip, "expenses", []) or []) == int(baseline["expense_count"])
    status = "PASS" if count_ok and not missing and not unexpected_forbidden else "REVIEW_REQUIRED"
    return {
        "status": status,
        "baseline": baseline_name,
        "expected_expense_count": baseline["expense_count"],
        "actual_expense_count": len(getattr(trip, "expenses", []) or []),
        "missing_page_amounts": [list(item) for item in missing],
        "forbidden_amounts_found": unexpected_forbidden,
    }


def _travel_structured_acceptance(paths: Iterable[Path]) -> dict[str, Any]:
    """Run the existing Travel parser without policy or Excel writes."""
    from travel_invoice_reviewer import parse_trip_file

    records: list[dict[str, Any]] = []
    for path in paths:
        try:
            trip = parse_trip_file(str(path))
            structured = list(getattr(trip, "structured_invoices", []) or [])
            field_assertions = _travel_field_assertions(path, trip)
            records.append({
                "file": str(path),
                "expense_count": len(trip.expenses),
                "structured_invoice_count": len(structured),
                "warning_count": len(trip.warnings),
                "field_assertions": field_assertions,
                "status": "PASS" if trip.expenses and field_assertions["status"] in {"PASS", "NOT_APPLICABLE"} else "REVIEW_REQUIRED",
            })
        except Exception as exc:
            records.append({
                "file": str(path),
                "expense_count": 0,
                "structured_invoice_count": 0,
                "warning_count": 1,
                "status": "REVIEW_REQUIRED",
                "error": f"{type(exc).__name__}: {exc}",
            })
    return {
        "file_count": len(records),
        "pass_count": sum(item["status"] == "PASS" for item in records),
        "review_required_count": sum(item["status"] != "PASS" for item in records),
        "field_assertion_pass_count": sum(
            item.get("field_assertions", {}).get("status") == "PASS" for item in records
        ),
        "records": records,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only OCR route acceptance scan")
    parser.add_argument(
        "--timesheet-root",
        default=str(Path(config.ROOT) / "CNTS-25165" / "01_IAV TMS & WP"),
    )
    parser.add_argument("--timesheet-limit", type=int, default=None)
    parser.add_argument(
        "--timesheet-offset",
        type=int,
        default=0,
        help="skip this many sorted Timesheet PDFs before selecting the batch",
    )
    parser.add_argument(
        "--timesheet-sample-per-directory",
        type=int,
        default=None,
        help="select up to N PDFs from each source/format directory",
    )
    parser.add_argument("--travel-file", action="append", default=[])
    parser.add_argument("--skip-travel-ocr", action="store_true")
    parser.add_argument("--travel-structured", action="store_true")
    parser.add_argument("--skip-travel", action="store_true")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    result: dict[str, Any] = {
        "timesheet": _timesheet_acceptance(
            Path(args.timesheet_root),
            args.timesheet_limit,
            args.timesheet_sample_per_directory,
            args.timesheet_offset,
        ),
    }
    if not args.skip_travel:
        travel_paths = [Path(item) for item in args.travel_file if Path(item).is_file()]
        if not args.skip_travel_ocr:
            result["travel"] = _travel_acceptance(travel_paths)
        if args.travel_structured:
            result["travel_structured"] = _travel_structured_acceptance(travel_paths)
    if args.json:
        print(json.dumps(result, ensure_ascii=True, indent=2))
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
