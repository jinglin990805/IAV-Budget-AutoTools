"""Shared document-processing primitives.

Business modules remain responsible for parsing and validation.  This package
only normalises OCR provider calls and evidence so Timesheet and Travel do not
depend on a provider's native return shape.
"""

from .models import Evidence, OCRBlock, OCRResult, SourceDocument
from .ocr_base import OCRProvider, ProviderUnavailable
from .ocr_router import OCRRouter, build_tesseract_router
from .paddle_provider import PaddleOCRProvider
from .preprocessing import enhance_scan, needs_ocr_retry, remove_form_lines
from .quality import OCRQualityReport, assess_ocr_text, ocr_candidate_rank
from .timesheet_table import extract_daily_hours_from_blocks
from .tesseract_provider import TesseractProvider
from .travel import (
    TravelInvoice,
    TravelLineItem,
    detect_duplicate_invoices,
    extract_travel_invoice,
    extract_travel_invoices,
    parse_locale_number,
    parse_travel_date,
    summarise_travel_validation,
    validate_travel_invoice,
)
from .travel_policy import TravelPolicyEngine
from .reimbursement_claim import (
    Claim,
    ClaimIndex,
    ClaimLine,
    ClaimWorkbookError,
    load_claim_index,
)
from .travel_claim_reconciliation import (
    ReconciliationRow,
    TripReconciliation,
    add_reconciliation_sheets,
    match_claim,
    reconcile_trip,
    reconcile_trips,
    write_reconciliation_workbook,
)
from .excel_adapters import AcceptanceTemplateAdapter, TemplateValidation, TravelExpenseTemplateAdapter
from .validation import (
    ValidationIssue,
    ValidationReport,
    build_timesheet_evidence,
    validate_ocr_completeness,
    validate_timesheet_fields,
)

__all__ = [
    "Evidence",
    "OCRBlock",
    "OCRProvider",
    "OCRResult",
    "OCRRouter",
    "PaddleOCRProvider",
    "ProviderUnavailable",
    "SourceDocument",
    "TesseractProvider",
    "ValidationIssue",
    "ValidationReport",
    "build_timesheet_evidence",
    "build_tesseract_router",
    "enhance_scan",
    "needs_ocr_retry",
    "remove_form_lines",
    "OCRQualityReport",
    "assess_ocr_text",
    "ocr_candidate_rank",
    "extract_daily_hours_from_blocks",
    "TravelInvoice",
    "TravelLineItem",
    "detect_duplicate_invoices",
    "extract_travel_invoice",
    "extract_travel_invoices",
    "parse_locale_number",
    "parse_travel_date",
    "summarise_travel_validation",
    "validate_travel_invoice",
    "TravelPolicyEngine",
    "Claim",
    "ClaimIndex",
    "ClaimLine",
    "ClaimWorkbookError",
    "load_claim_index",
    "ReconciliationRow",
    "TripReconciliation",
    "add_reconciliation_sheets",
    "match_claim",
    "reconcile_trip",
    "reconcile_trips",
    "write_reconciliation_workbook",
    "AcceptanceTemplateAdapter",
    "TemplateValidation",
    "TravelExpenseTemplateAdapter",
    "validate_ocr_completeness",
    "validate_timesheet_fields",
]
