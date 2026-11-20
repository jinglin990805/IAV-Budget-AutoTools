"""Tesseract OCR provider.

The provider owns temporary-file handling, subprocess invocation, TSV parsing
and confidence normalisation.  Existing business parsers can continue to use
plain text while gradually adopting positioned blocks/evidence.
"""

from __future__ import annotations

import os
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Any

from .models import OCRBlock, OCRResult
from .ocr_base import ProviderUnavailable


class TesseractProvider:
    name = "tesseract"

    def __init__(
        self,
        executable: str | os.PathLike[str],
        languages: str = "eng",
        timeout: int = 90,
    ) -> None:
        self.executable = Path(executable)
        self.languages = languages or "eng"
        self.timeout = max(1, int(timeout))

    @property
    def available(self) -> bool:
        return self.executable.is_file()

    @classmethod
    def list_languages(
        cls,
        executable: str | os.PathLike[str],
        timeout: int = 15,
    ) -> list[str]:
        executable_path = Path(executable)
        if not executable_path.is_file():
            return []
        try:
            result = subprocess.run(
                [str(executable_path), "--list-langs"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=max(1, int(timeout)),
                check=False,
                # The workbench may run under pythonw. Without this flag,
                # Windows creates a visible console for every Tesseract call.
                creationflags=(getattr(subprocess, "CREATE_NO_WINDOW", 0)
                               if os.name == "nt" else 0),
            )
        except (OSError, subprocess.SubprocessError):
            return []
        return [
            line.strip()
            for line in (result.stdout or "").splitlines()
            if line.strip() and not line.lower().startswith("list of available")
        ]

    @staticmethod
    def _parse_tsv(output: str, page: int) -> list[OCRBlock]:
        blocks: list[OCRBlock] = []
        for raw_line in (output or "").splitlines()[1:]:
            columns = raw_line.split("\t")
            if len(columns) < 12 or not columns[11].strip():
                continue
            try:
                confidence = float(columns[10])
                if confidence < 0:
                    confidence = None
                left, top = int(columns[6]), int(columns[7])
                width, height = int(columns[8]), int(columns[9])
                line_id = tuple(int(columns[index]) for index in range(1, 5))
            except (TypeError, ValueError):
                continue
            blocks.append(
                OCRBlock(
                    text=columns[11].strip(),
                    confidence=confidence,
                    bbox=(left, top, left + width, top + height),
                    page=page,
                    line_id=line_id,
                    metadata={
                        "left": left,
                        "top": top,
                        "width": width,
                        "height": height,
                    },
                )
            )
        return blocks

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
        started = time.perf_counter()
        if not self.available:
            return OCRResult(
                engine=self.name,
                page=page,
                warnings=[f"Tesseract executable not found: {self.executable}"],
            )
        if output_mode not in {"text", "tsv"}:
            return OCRResult(
                engine=self.name,
                page=page,
                warnings=[f"Unsupported Tesseract output mode: {output_mode}"],
            )

        descriptor, image_name = tempfile.mkstemp(
            prefix="iav_ocr_provider_", suffix=".png",
        )
        os.close(descriptor)
        image_path = Path(image_name)
        warnings: list[str] = []
        try:
            image.save(str(image_path), format="PNG", dpi=(300, 300))
            command = [
                str(self.executable), str(image_path), "stdout",
                "-l", languages or self.languages,
                "--psm", str(psm),
            ]
            if output_mode == "tsv":
                command.append("tsv")
            if whitelist:
                command.extend(["-c", f"tessedit_char_whitelist={whitelist}"])
            try:
                result = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    timeout=self.timeout,
                    check=False,
                    creationflags=(getattr(subprocess, "CREATE_NO_WINDOW", 0)
                                   if os.name == "nt" else 0),
                )
            except subprocess.TimeoutExpired:
                warnings.append(f"Tesseract OCR timeout after {self.timeout}s")
                return OCRResult(
                    engine=self.name,
                    page=page,
                    warnings=warnings,
                    elapsed_ms=(time.perf_counter() - started) * 1000,
                )
            except OSError as exc:
                warnings.append(f"Tesseract OCR failed to start: {exc}")
                return OCRResult(
                    engine=self.name,
                    page=page,
                    warnings=warnings,
                    elapsed_ms=(time.perf_counter() - started) * 1000,
                )

            output = result.stdout or ""
            if result.returncode != 0:
                warnings.append((result.stderr or "").strip()[:500]
                                or f"Tesseract OCR returned code {result.returncode}")
            blocks = self._parse_tsv(output, page) if output_mode == "tsv" else []
            text = "\n".join(block.text for block in blocks) if output_mode == "tsv" else output.strip()
            confidences = [
                block.confidence for block in blocks
                if block.confidence is not None
            ]
            confidence = sum(confidences) / len(confidences) / 100 if confidences else None
            return OCRResult(
                text=text,
                blocks=blocks,
                engine=self.name,
                confidence=confidence,
                page=page,
                warnings=warnings,
                elapsed_ms=(time.perf_counter() - started) * 1000,
            )
        finally:
            try:
                image_path.unlink()
            except OSError:
                pass


def require_tesseract(
    executable: str | os.PathLike[str],
    languages: str = "eng",
    timeout: int = 90,
) -> TesseractProvider:
    provider = TesseractProvider(executable, languages=languages, timeout=timeout)
    if not provider.available:
        raise ProviderUnavailable(f"Tesseract executable not found: {provider.executable}")
    return provider
