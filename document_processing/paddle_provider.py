"""Optional PaddleOCR provider with a provider-neutral result adapter.

PaddleOCR remains lazy and optional. Importing the workbench must continue to
work on the current Tesseract-only deployment; model construction is attempted
only when a benchmark explicitly creates this provider.
"""

from __future__ import annotations

import os
import tempfile
import time
from pathlib import Path
from typing import Any

from .models import OCRBlock, OCRResult
from .ocr_base import ProviderUnavailable


class PaddleOCRProvider:
    name = "paddleocr"

    def __init__(self, **options: Any) -> None:
        try:
            from paddleocr import PaddleOCR  # type: ignore
        except Exception as exc:
            raise ProviderUnavailable(
                "PaddleOCR/PaddlePaddle is not installed; install it only in an isolated benchmark environment"
            ) from exc

        lang = options.pop("lang", "ch")
        device = options.pop("device", "cpu")
        defaults = {
            "lang": lang,
            "device": device,
            "use_doc_orientation_classify": False,
            "use_doc_unwarping": False,
            "use_textline_orientation": False,
        }
        defaults.update(options)
        try:
            self._ocr = PaddleOCR(**defaults)
        except TypeError:
            try:
                self._ocr = PaddleOCR(lang=lang, use_angle_cls=False)
            except Exception as exc:
                raise ProviderUnavailable(f"PaddleOCR initialisation failed: {exc}") from exc
        except Exception as exc:
            raise ProviderUnavailable(f"PaddleOCR initialisation failed: {exc}") from exc

    @staticmethod
    def _as_list(value: Any) -> list[Any]:
        if value is None:
            return []
        if hasattr(value, "tolist"):
            value = value.tolist()
        if isinstance(value, (list, tuple)):
            return list(value)
        return [value]

    @staticmethod
    def _box(value: Any) -> tuple[float, float, float, float] | None:
        if value is None:
            return None
        if hasattr(value, "tolist"):
            value = value.tolist()
        try:
            if len(value) == 4 and all(isinstance(item, (int, float)) for item in value):
                numbers = [float(item) for item in value]
                return (numbers[0], numbers[1], numbers[2], numbers[3])
            points = [[float(coord) for coord in point] for point in value]
            xs = [point[0] for point in points]
            ys = [point[1] for point in points]
            return (min(xs), min(ys), max(xs), max(ys))
        except (TypeError, ValueError, IndexError):
            return None

    @classmethod
    def _mapping(cls, result: Any) -> dict[str, Any]:
        data: Any = result
        if hasattr(data, "json"):
            data = data.json
            if callable(data):
                data = data()
        elif hasattr(data, "res"):
            data = data.res
        if isinstance(data, dict) and isinstance(data.get("res"), dict):
            data = data["res"]
        return data if isinstance(data, dict) else {}

    @classmethod
    def _normalise_result(
        cls, raw: Any, *, page: int = 1,
    ) -> tuple[str, list[OCRBlock], float | None]:
        """Normalise PaddleOCR 3.x result objects and legacy 2.x lists."""
        if isinstance(raw, list) and raw and isinstance(raw[0], (list, tuple)):
            blocks: list[OCRBlock] = []
            for item in raw:
                if len(item) < 2:
                    continue
                polygon, recognised = item[0], item[1]
                if not isinstance(recognised, (list, tuple)) or not recognised:
                    continue
                text = str(recognised[0] or "").strip()
                if not text:
                    continue
                try:
                    score = float(recognised[1]) if len(recognised) > 1 else None
                except (TypeError, ValueError):
                    score = None
                blocks.append(OCRBlock(text=text, confidence=score, bbox=cls._box(polygon), page=page))
            values = [item.confidence for item in blocks if item.confidence is not None]
            return "\n".join(item.text for item in blocks), blocks, sum(values) / len(values) if values else None

        data = cls._mapping(raw)
        texts = [str(item or "").strip() for item in cls._as_list(data.get("rec_texts"))]
        scores = cls._as_list(data.get("rec_scores"))
        raw_boxes = data.get("rec_boxes")
        if raw_boxes is None:
            raw_boxes = data.get("dt_polys")
        boxes = cls._as_list(raw_boxes)
        blocks = []
        for index, text in enumerate(texts):
            if not text:
                continue
            try:
                score = float(scores[index]) if index < len(scores) else None
            except (TypeError, ValueError):
                score = None
            bbox = cls._box(boxes[index]) if index < len(boxes) else None
            blocks.append(OCRBlock(text=text, confidence=score, bbox=bbox, page=page))
        values = [item.confidence for item in blocks if item.confidence is not None]
        return "\n".join(item.text for item in blocks), blocks, sum(values) / len(values) if values else None

    def recognize(self, image: Any, **kwargs: Any) -> OCRResult:
        started = time.perf_counter()
        page = int(kwargs.get("page", 1))
        descriptor, image_name = tempfile.mkstemp(prefix="iav_paddle_", suffix=".png")
        os.close(descriptor)
        image_path = Path(image_name)
        try:
            image.save(str(image_path), format="PNG")
            if hasattr(self._ocr, "predict"):
                raw_results = self._ocr.predict(str(image_path))
            elif hasattr(self._ocr, "ocr"):
                raw_results = self._ocr.ocr(str(image_path), cls=True)
            else:
                raise ProviderUnavailable("PaddleOCR object has no predict/ocr method")
            raw_list = list(raw_results) if not isinstance(raw_results, list) else raw_results
            raw = raw_list[0] if len(raw_list) == 1 else raw_list
            text, blocks, confidence = self._normalise_result(raw, page=page)
            return OCRResult(
                text=text,
                blocks=blocks,
                engine=self.name,
                confidence=confidence,
                page=page,
                warnings=[] if text else ["PaddleOCR returned no text"],
                elapsed_ms=(time.perf_counter() - started) * 1000,
            )
        except ProviderUnavailable:
            raise
        except Exception as exc:
            return OCRResult(
                engine=self.name,
                page=page,
                warnings=[f"PaddleOCR inference failed: {exc}"],
                elapsed_ms=(time.perf_counter() - started) * 1000,
            )
        finally:
            try:
                image_path.unlink()
            except OSError:
                pass
