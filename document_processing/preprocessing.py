"""Small, deterministic image enhancements for difficult scanned pages."""

from __future__ import annotations

from typing import Any


def enhance_scan(image: Any) -> Any:
    """Return a contrast-normalised RGB image suitable for an OCR retry."""

    from PIL import ImageEnhance, ImageOps

    gray = ImageOps.grayscale(image)
    normalised = ImageOps.autocontrast(gray, cutoff=1)
    contrasted = ImageEnhance.Contrast(normalised).enhance(1.6)
    return contrasted.convert("RGB")


def needs_ocr_retry(text: str, *, minimum_chars: int = 180) -> bool:
    """Identify pages whose first OCR pass is too weak for safe extraction."""

    return len((text or "").strip()) < max(1, int(minimum_chars))


def remove_form_lines(image: Any, *, dark_threshold: int = 170) -> Any:
    """Erase long horizontal/vertical form lines while preserving text pixels."""

    from PIL import ImageDraw, ImageOps

    gray = ImageOps.grayscale(image)
    width, height = gray.size
    binary = gray.point(lambda pixel: 0 if pixel < dark_threshold else 255)
    pixels = binary.load()
    horizontal = [
        y for y in range(height)
        if sum(pixels[x, y] == 0 for x in range(width)) >= width * 0.45
    ]
    vertical = [
        x for x in range(width)
        if sum(pixels[x, y] == 0 for y in range(height)) >= height * 0.55
    ]
    cleaned = gray.copy()
    draw = ImageDraw.Draw(cleaned)
    line_width = max(2, round(min(width, height) / 900))
    for y in horizontal:
        draw.line((0, y, width, y), fill=255, width=line_width)
    for x in vertical:
        draw.line((x, 0, x, height), fill=255, width=line_width)
    return cleaned.convert("RGB")
