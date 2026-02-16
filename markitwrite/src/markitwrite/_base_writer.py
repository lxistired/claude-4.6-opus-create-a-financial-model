"""Base class for all document writers."""

from __future__ import annotations

import io
from dataclasses import dataclass, field
from typing import BinaryIO, Optional


@dataclass
class WriteResult:
    """Result of a write/paste operation."""

    output: bytes
    target_format: str
    images_inserted: int = 1
    metadata: dict = field(default_factory=dict)

    def to_stream(self) -> io.BytesIO:
        """Return the output as a seekable BytesIO stream."""
        stream = io.BytesIO(self.output)
        stream.seek(0)
        return stream

    def save(self, path: str) -> None:
        """Save the output to a file."""
        with open(path, "wb") as f:
            f.write(self.output)


class DocumentWriter:
    """Base class for all format-specific document writers.

    Mirrors the markitdown DocumentConverter pattern but for the write direction.
    Each writer handles inserting images into a specific document format.
    """

    ACCEPTED_EXTENSIONS: list[str] = []

    def accepts(self, target_format: str, **kwargs) -> bool:
        """Determine if this writer can handle the target format.

        Args:
            target_format: File extension (e.g. ".docx", ".pptx", ".md")

        Returns:
            True if this writer can handle the format.
        """
        normalized = target_format.lower()
        if not normalized.startswith("."):
            normalized = f".{normalized}"
        return normalized in self.ACCEPTED_EXTENSIONS

    def insert_image(
        self,
        image_stream: BinaryIO,
        document_stream: Optional[BinaryIO] = None,
        target_format: str = "",
        position: Optional[dict] = None,
        size: Optional[dict] = None,
        **kwargs,
    ) -> WriteResult:
        """Insert an image into a document.

        Args:
            image_stream: The image data to insert.
            document_stream: Existing document to modify (None = create new).
            target_format: Target format extension.
            position: Placement hints (e.g. {"slide": 3}, {"paragraph": 5}).
            size: Sizing hints (e.g. {"width": 6.0}, {"width": 4, "height": 3}).
            **kwargs: Format-specific options.

        Returns:
            WriteResult containing the modified document bytes.
        """
        raise NotImplementedError(
            f"{self.__class__.__name__} must implement insert_image()"
        )
