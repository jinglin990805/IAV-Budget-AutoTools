"""Provider routing with a stable, provider-neutral return type."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from pathlib import Path
from typing import Any

from .models import OCRResult
from .ocr_base import OCRProvider, ProviderUnavailable
from .tesseract_provider import TesseractProvider


class OCRRouter:
    def __init__(
        self,
        primary: OCRProvider,
        fallbacks: Sequence[OCRProvider] = (),
    ) -> None:
        self.primary = primary
        self.fallbacks = tuple(fallbacks)

    def recognize(
        self,
        image: Any,
        *,
        languages: str | None = None,
        psm: str = "6",
        whitelist: str = "",
        page: int = 1,
        output_mode: str = "text",
        accept: Callable[[OCRResult], bool] | None = None,
    ) -> OCRResult:
        providers = (self.primary, *self.fallbacks)
        warnings: list[str] = []
        for index, provider in enumerate(providers):
            try:
                result = provider.recognize(
                    image,
                    languages=languages,
                    psm=psm,
                    whitelist=whitelist,
                    page=page,
                    output_mode=output_mode,
                )
            except ProviderUnavailable as exc:
                warnings.append(f"{getattr(provider, 'name', 'ocr')}: {exc}")
                continue
            warnings.extend(result.warnings)
            if result.ok and (accept is None or accept(result)):
                if warnings and warnings != result.warnings:
                    result.warnings = list(dict.fromkeys(warnings))
                return result
            if result.ok and index == len(providers) - 1:
                result.warnings = list(dict.fromkeys(warnings))
                return result
        return OCRResult(
            engine="none",
            page=page,
            warnings=warnings or ["No OCR provider returned text"],
        )


def build_tesseract_router(
    executable: str | Path,
    languages: str = "eng",
    timeout: int = 90,
) -> OCRRouter:
    return OCRRouter(TesseractProvider(executable, languages=languages, timeout=timeout))
