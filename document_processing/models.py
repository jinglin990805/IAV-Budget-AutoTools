"""Provider-neutral OCR and source evidence models."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass(frozen=True)
class SourceDocument:
    """A source reference carried alongside extraction evidence."""

    source_file: str
    sha256: str = ""
    page: int = 1
    document_type: str = ""
    extraction_method: str = ""
    run_id: str = ""
    project_id: str = ""


@dataclass(frozen=True)
class OCRBlock:
    """One positioned OCR word/block in provider-neutral form."""

    text: str
    confidence: float | None = None
    bbox: tuple[float, float, float, float] | None = None
    page: int = 1
    line_id: tuple[int, ...] | None = None
    metadata: dict[str, Any] = field(default_factory=dict)

    def as_dict(self) -> dict[str, Any]:
        value = asdict(self)
        if self.bbox is not None:
            value["bbox"] = list(self.bbox)
        if self.line_id is not None:
            value["line_id"] = list(self.line_id)
        return value


@dataclass
class OCRResult:
    """Normalised OCR output shared by all providers."""

    text: str = ""
    blocks: list[OCRBlock] = field(default_factory=list)
    engine: str = "unknown"
    confidence: float | None = None
    page: int = 1
    warnings: list[str] = field(default_factory=list)
    elapsed_ms: float | None = None

    @property
    def ok(self) -> bool:
        return bool(self.text.strip())

    @property
    def error(self) -> str:
        """Compatibility helper for callers that expect one error string."""

        return self.warnings[-1] if self.warnings else ""

    def as_dict(self) -> dict[str, Any]:
        return {
            "text": self.text,
            "blocks": [block.as_dict() for block in self.blocks],
            "engine": self.engine,
            "confidence": self.confidence,
            "page": self.page,
            "warnings": list(self.warnings),
            "elapsed_ms": self.elapsed_ms,
        }


@dataclass(frozen=True)
class Evidence:
    """Traceable evidence for a parsed business field."""

    field: str
    value: Any
    source_file: str
    page: int = 1
    bbox: tuple[float, float, float, float] | None = None
    raw_text: str = ""
    ocr_engine: str = ""
    ocr_confidence: float | None = None
    validation_status: str = "PENDING"

    def as_dict(self) -> dict[str, Any]:
        value = asdict(self)
        if self.bbox is not None:
            value["bbox"] = list(self.bbox)
        return value
