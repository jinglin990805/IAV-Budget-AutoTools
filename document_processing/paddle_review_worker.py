"""One-shot PaddleOCR worker for optional travel page review.

This script runs only inside the isolated ``.paddle-venv``. It writes OCR text
and crop-relative block geometry to a caller-provided JSON file; it never
changes a workbook or expense record.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
import types
from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
PADDLE_HOME = ROOT / ".paddle-home"
PADDLE_CACHE = ROOT / ".paddle-cache"
PADDLE_TEMP = ROOT / ".paddle-tmp"
MODEL_ROOT = PADDLE_CACHE / "paddlex" / "official_models"
DET_MODEL = MODEL_ROOT / "PP-OCRv6_medium_det"
REC_MODEL = MODEL_ROOT / "PP-OCRv6_medium_rec"


def configure_environment() -> None:
    """Keep Paddle, model caches, and temporary writes inside the workspace."""
    PADDLE_HOME.mkdir(parents=True, exist_ok=True)
    (PADDLE_CACHE / "paddlex").mkdir(parents=True, exist_ok=True)
    (PADDLE_CACHE / "modelscope").mkdir(parents=True, exist_ok=True)
    (PADDLE_CACHE / "huggingface").mkdir(parents=True, exist_ok=True)
    PADDLE_TEMP.mkdir(parents=True, exist_ok=True)
    os.environ["USERPROFILE"] = str(PADDLE_HOME)
    os.environ["PADDLE_PDX_CACHE_HOME"] = str(PADDLE_CACHE / "paddlex")
    os.environ["MODELSCOPE_CACHE"] = str(PADDLE_CACHE / "modelscope")
    os.environ["HF_HOME"] = str(PADDLE_CACHE / "huggingface")
    os.environ["HUGGINGFACE_HUB_CACHE"] = str(PADDLE_CACHE / "huggingface" / "hub")
    os.environ["HF_HUB_CACHE"] = os.environ["HUGGINGFACE_HUB_CACHE"]
    os.environ["TEMP"] = str(PADDLE_TEMP)
    os.environ["TMP"] = str(PADDLE_TEMP)
    os.environ["PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK"] = "True"


def import_provider():
    # The isolated environment intentionally lacks workbench packages such as
    # openpyxl. Skip document_processing.__init__ and load only the small,
    # provider-neutral adapter and its dataclasses.
    package = types.ModuleType("document_processing")
    package.__path__ = [str(ROOT / "document_processing")]
    package.__package__ = "document_processing"
    sys.modules["document_processing"] = package
    from document_processing.paddle_provider import PaddleOCRProvider

    return PaddleOCRProvider


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf", required=True)
    parser.add_argument("--pages", required=True, help="Comma-separated 1-based pages")
    parser.add_argument("--dpi", type=int, default=600)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    configure_environment()
    if not DET_MODEL.is_dir() or not REC_MODEL.is_dir():
        raise RuntimeError("Cached PP-OCRv6 medium detection/recognition models are unavailable")

    import pypdfium2 as pdfium

    PaddleOCRProvider = import_provider()
    provider = PaddleOCRProvider(
        lang="ch",
        device="cpu",
        text_detection_model_dir=str(DET_MODEL),
        text_recognition_model_dir=str(REC_MODEL),
        text_det_limit_side_len=2560,
        text_det_limit_type="max",
        text_rec_score_thresh=0.0,
    )
    pages = list(dict.fromkeys(int(value) for value in args.pages.split(",")))
    document = pdfium.PdfDocument(args.pdf)
    results = []
    try:
        if any(page < 1 or page > len(document) for page in pages):
            raise ValueError(f"Requested page is outside PDF range (1-{len(document)})")
        for page_number in pages:
            started = time.perf_counter()
            page = document.get_page(page_number - 1)
            image = None
            crop = None
            red_channel_crop = None
            try:
                image = page.render(scale=args.dpi / 72, prefer_bgrx=True).to_pil().convert("RGB")
                width, height = image.size
                left = int(width * 0.20)
                top = int(height * 0.45)
                right = int(width * 0.85)
                bottom = height
                crop = image.crop((left, top, right, bottom))
                result = provider.recognize(crop, page=page_number)

                red_channel_crop = crop.getchannel("R").resize(
                    (int(crop.width * 1.5), int(crop.height * 1.5)),
                    resample=Image.Resampling.LANCZOS,
                )
                red_channel_result = provider.recognize(red_channel_crop, page=page_number)

                def serialise_blocks(ocr_result):
                    blocks = []
                    for block in ocr_result.blocks:
                        if not block.text.strip():
                            continue
                        blocks.append({
                            "text": block.text,
                            "confidence": block.confidence,
                            "bbox_crop_pixels": list(block.bbox) if block.bbox else None,
                        })
                    return blocks

                blocks = serialise_blocks(result)
                red_channel_blocks = serialise_blocks(red_channel_result)
                results.append({
                    "page": page_number,
                    "render_size": [width, height],
                    "region_bbox_pixels": [left, top, right, bottom],
                    "region_bbox_normalized": [left / width, top / height, right / width, bottom / height],
                    "text": result.text,
                    "blocks": blocks,
                    "confidence": result.confidence,
                    "elapsed_ms": round((time.perf_counter() - started) * 1000, 1),
                    "warnings": result.warnings,
                    "scale_factor": 1.0,
                    "preprocessing": "Original RGB crop",
                    "red_channel": {
                        "text": red_channel_result.text,
                        "blocks": red_channel_blocks,
                        "confidence": red_channel_result.confidence,
                        "elapsed_ms": red_channel_result.elapsed_ms,
                        "warnings": red_channel_result.warnings,
                        "scale_factor": 1.5,
                        "preprocessing": "R-channel, 1.5x LANCZOS",
                    },
                })
            finally:
                if red_channel_crop is not None:
                    red_channel_crop.close()
                if crop is not None:
                    crop.close()
                if image is not None:
                    image.close()
                if hasattr(page, "close"):
                    page.close()
    finally:
        if hasattr(document, "close"):
            document.close()

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8", newline="\n") as output:
        json.dump({"pages": results}, output, ensure_ascii=False, separators=(",", ":"))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"PaddleOCR worker failed: {type(exc).__name__}: {exc}", file=sys.stderr)
        raise SystemExit(1)
