"""Writer for Markdown (.md) files with image embedding."""

from __future__ import annotations

import base64
import io
import mimetypes
import os
import shutil
from typing import BinaryIO, Optional

from markitwrite._base_writer import DocumentWriter, WriteResult


class MarkdownWriter(DocumentWriter):
    """Insert images into Markdown (.md) files.

    Supports two modes:
        - **embed=True** (default): Base64-encode the image inline.
          Produces a self-contained Markdown file.
        - **embed=False**: Save the image to disk next to the Markdown file
          and insert a relative file reference ``![alt](images/name.png)``.

    Kwargs:
        embed (bool): Embed as base64 (True) or file reference (False).
        alt_text (str): Alt text for the image (default "image").
        image_dir (str): Subdirectory name for saved images (default "images").
        image_filename (str): Override the saved image filename.
        mime_type (str): MIME type of the image (auto-detected when possible).
    """

    ACCEPTED_EXTENSIONS = [".md", ".markdown", ".mdown", ".mkd"]

    def insert_image(
        self,
        image_stream: BinaryIO,
        document_stream: Optional[BinaryIO] = None,
        target_format: str = ".md",
        position: Optional[dict] = None,
        size: Optional[dict] = None,
        **kwargs,
    ) -> WriteResult:
        image_stream.seek(0)
        image_data = image_stream.read()

        alt_text = kwargs.get("alt_text", "image")
        embed = kwargs.get("embed", True)
        mime_type = kwargs.get("mime_type", "image/png")

        if embed:
            b64 = base64.b64encode(image_data).decode("utf-8")
            img_tag = f"![{alt_text}](data:{mime_type};base64,{b64})"
        else:
            # File reference mode - determine path
            image_dir = kwargs.get("image_dir", "images")
            image_filename = kwargs.get("image_filename")
            if not image_filename:
                ext = mimetypes.guess_extension(mime_type) or ".png"
                image_filename = f"image{ext}"
            img_tag = f"![{alt_text}]({image_dir}/{image_filename})"

            # If a target path is known, save the image file
            target_path = kwargs.get("_target_path")
            if target_path:
                target_dir = os.path.dirname(target_path)
                full_image_dir = os.path.join(target_dir, image_dir)
                os.makedirs(full_image_dir, exist_ok=True)
                img_path = os.path.join(full_image_dir, image_filename)
                with open(img_path, "wb") as f:
                    f.write(image_data)

        # Build output content
        if document_stream:
            document_stream.seek(0)
            existing = document_stream.read().decode("utf-8")
            # Insert at position if specified
            if position and "after_heading" in position:
                heading = position["after_heading"]
                lines = existing.split("\n")
                insert_idx = None
                for i, line in enumerate(lines):
                    if line.lstrip().startswith("#") and heading in line:
                        insert_idx = i + 1
                        break
                if insert_idx is not None:
                    lines.insert(insert_idx, "")
                    lines.insert(insert_idx + 1, img_tag)
                    content = "\n".join(lines)
                else:
                    content = existing.rstrip() + "\n\n" + img_tag + "\n"
            else:
                content = existing.rstrip() + "\n\n" + img_tag + "\n"
        else:
            content = img_tag + "\n"

        output_bytes = content.encode("utf-8")

        return WriteResult(
            output=output_bytes,
            target_format=".md",
            images_inserted=1,
            metadata={
                "embed": embed,
                "alt_text": alt_text,
            },
        )
