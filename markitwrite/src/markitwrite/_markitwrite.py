"""Main dispatcher for MarkItWrite - the AI virtual clipboard."""

from __future__ import annotations

import io
import mimetypes
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO, Optional, Union

from markitwrite._base_writer import DocumentWriter, WriteResult


@dataclass
class _WriterRegistration:
    writer: DocumentWriter
    priority: float


class MarkItWrite:
    """AI virtual clipboard - paste images into any document format.

    Usage:
        writer = MarkItWrite()
        writer.paste("screenshot.png", target="report.docx")
        writer.paste("chart.png", target="slides.pptx", slide=3)
        writer.paste("diagram.png", target="notes.md", embed=True)
    """

    def __init__(self, enable_builtins: bool = True):
        self._writers: list[_WriterRegistration] = []

        if enable_builtins:
            self._register_builtins()

    def _register_builtins(self) -> None:
        """Register all built-in writers."""
        # Each import is guarded so missing optional deps don't break everything.
        writers_to_try = [
            ("markitwrite.writers._docx_writer", "DocxWriter"),
            ("markitwrite.writers._pptx_writer", "PptxWriter"),
            ("markitwrite.writers._markdown_writer", "MarkdownWriter"),
        ]
        for module_path, class_name in writers_to_try:
            try:
                mod = __import__(module_path, fromlist=[class_name])
                writer_cls = getattr(mod, class_name)
                self.register_writer(writer_cls())
            except ImportError:
                # Optional dependency not installed - skip silently
                pass

    def register_writer(self, writer: DocumentWriter, priority: float = 0.0) -> None:
        """Register a document writer.

        Args:
            writer: The writer instance.
            priority: Lower values = tried first (default 0.0).
        """
        self._writers.append(_WriterRegistration(writer=writer, priority=priority))

    def _resolve_image(
        self, image_source: Union[str, bytes, Path, BinaryIO]
    ) -> tuple[BinaryIO, str]:
        """Resolve image source to a (stream, mime_type) pair."""
        if isinstance(image_source, (str, Path)):
            path = str(image_source)
            mime, _ = mimetypes.guess_type(path)
            with open(path, "rb") as f:
                data = f.read()
            return io.BytesIO(data), mime or "image/png"
        elif isinstance(image_source, bytes):
            return io.BytesIO(image_source), "image/png"
        else:
            # Already a stream
            pos = image_source.tell()
            data = image_source.read()
            image_source.seek(pos)
            return io.BytesIO(data), "image/png"

    def _resolve_target_format(
        self, target: Optional[str], target_format: Optional[str]
    ) -> str:
        """Determine the target format from target path or explicit format."""
        if target_format:
            fmt = target_format.lower()
            return fmt if fmt.startswith(".") else f".{fmt}"
        if target:
            _, ext = os.path.splitext(target)
            if ext:
                return ext.lower()
        return ".docx"  # default

    def _find_writer(self, target_format: str, **kwargs) -> DocumentWriter:
        """Find a writer that accepts the target format."""
        sorted_writers = sorted(self._writers, key=lambda r: r.priority)
        for reg in sorted_writers:
            if reg.writer.accepts(target_format, **kwargs):
                return reg.writer
        available = []
        for reg in self._writers:
            available.extend(reg.writer.ACCEPTED_EXTENSIONS)
        raise ValueError(
            f"No writer found for format '{target_format}'. "
            f"Available formats: {', '.join(available) or 'none (install optional deps)'}"
        )

    def paste(
        self,
        image_source: Union[str, bytes, Path, BinaryIO],
        target: Optional[str] = None,
        target_format: Optional[str] = None,
        position: Optional[dict] = None,
        size: Optional[dict] = None,
        **kwargs,
    ) -> WriteResult:
        """Paste an image into a document - the core API.

        Args:
            image_source: Image path, bytes, or stream.
            target: Target file path. If the file exists, the image is inserted
                    into it; otherwise a new document is created then saved here.
                    If None, a new document is created in memory only.
            target_format: Explicit format override (e.g. ".docx"). Inferred from
                          ``target`` extension when omitted.
            position: Placement hints (format-specific). Examples:
                      {"slide": 3} for PPTX, {"paragraph": 5} for DOCX.
            size: Sizing hints. Examples:
                  {"width": 6.0} (inches), {"width": 4, "height": 3}.
            **kwargs: Passed through to the writer (e.g. embed=True for Markdown).

        Returns:
            WriteResult with the document bytes.
        """
        fmt = self._resolve_target_format(target, target_format)
        writer = self._find_writer(fmt, **kwargs)

        image_stream, mime_type = self._resolve_image(image_source)
        kwargs.setdefault("mime_type", mime_type)

        # Load existing document if target file exists
        document_stream: Optional[BinaryIO] = None
        if target and os.path.isfile(target):
            with open(target, "rb") as f:
                document_stream = io.BytesIO(f.read())

        result = writer.insert_image(
            image_stream=image_stream,
            document_stream=document_stream,
            target_format=fmt,
            position=position,
            size=size,
            **kwargs,
        )

        # Save to target path if provided
        if target:
            result.save(target)

        return result

    def supported_formats(self) -> list[str]:
        """Return list of supported target format extensions."""
        formats = []
        for reg in self._writers:
            formats.extend(reg.writer.ACCEPTED_EXTENSIONS)
        return sorted(set(formats))
