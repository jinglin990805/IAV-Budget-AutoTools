"""VS Code entry point for stepping through one travel PDF parse."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

from travel_invoice_reviewer import allocate_trip_expenses, parse_trip_file


def _iso(value):
    return value.isoformat() if hasattr(value, "isoformat") else value


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser(
        description="Parse one travel PDF and print hotel-date/policy evidence."
    )
    parser.add_argument("pdf", type=Path, help="Path to the travel PDF")
    parser.add_argument("--year", type=int, default=None, help="Fallback document year")
    args = parser.parse_args()
    if not args.pdf.is_file():
        parser.error(f"PDF does not exist: {args.pdf}")

    trip = parse_trip_file(str(args.pdf), default_year=args.year)
    allocate_trip_expenses(trip)

    accommodation = []
    for expense in trip.expenses:
        if expense.category != "accommodation":
            continue
        accommodation.append({
            "amount": expense.amount,
            "source_page": expense.source_page,
            "start": _iso(expense.start),
            "end": _iso(expense.end),
            "nights": expense.accommodation_nights,
            "nights_source": expense.accommodation_nights_source,
            "date_source": expense.service_date_source,
            "folio_page": expense.service_date_evidence_page,
            "folio_total": expense.accommodation_folio_total,
            "folio_text": expense.service_date_evidence_text,
        })

    result = {
        "source_file": trip.source_file,
        "trip": {
            "name": trip.name,
            "city": trip.city,
            "start": _iso(trip.start),
            "end": _iso(trip.end),
        },
        "accommodation": accommodation,
        "hotel_policy_trace": [
            item for item in trip.policy_trace
            if item.get("rule_id") == "TRAVEL.HOTEL_CAP"
        ],
        "page_coverage": trip.page_coverage,
        "warnings": trip.warnings,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
