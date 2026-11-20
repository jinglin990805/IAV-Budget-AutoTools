"""OCR provider contract and provider-independent errors."""

from __future__ import annotations

from typing import Any, Protocol

from .models import OCRResult


class ProviderUnavailable(RuntimeError):
    """Raised when an optional OCR provider is not installed or configured."""


class OCRProvider(Protocol):
    name: str

    def recognize(
        self,
        image: Any,
        *,
        languages: str | None = None,
        psm: str = "6",
        whitelist: str = "",
        page: int = 1,
        output_mode: str = "text",
    ) -> OCRResult:
        ...
