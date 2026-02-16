"""CLI entry point for markitwrite.

Usage:
    markitwrite paste screenshot.png --to output.docx
    markitwrite paste chart.png --to slides.pptx --slide 3
    markitwrite paste diagram.png --to notes.md --embed
    markitwrite paste image.png --to report.docx --width 4 --height 3
    markitwrite formats
"""

from __future__ import annotations

import argparse
import sys


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="markitwrite",
        description="AI virtual clipboard - paste images into any document format.",
    )
    subparsers = parser.add_subparsers(dest="command")

    # -- paste command --
    paste_parser = subparsers.add_parser(
        "paste", help="Paste an image into a document."
    )
    paste_parser.add_argument("image", help="Path to the image file.")
    paste_parser.add_argument(
        "--to", dest="target", required=True, help="Target document path."
    )
    paste_parser.add_argument(
        "--format",
        dest="target_format",
        default=None,
        help="Explicit target format (e.g. .docx). Inferred from --to if omitted.",
    )
    paste_parser.add_argument(
        "--width", type=float, default=None, help="Image width in inches."
    )
    paste_parser.add_argument(
        "--height", type=float, default=None, help="Image height in inches."
    )
    paste_parser.add_argument(
        "--slide", type=int, default=None, help="Slide number for PPTX (1-indexed)."
    )
    paste_parser.add_argument(
        "--paragraph",
        type=int,
        default=None,
        help="Paragraph index for DOCX (0-indexed).",
    )
    paste_parser.add_argument(
        "--embed",
        action="store_true",
        default=True,
        help="Embed image as base64 in Markdown (default).",
    )
    paste_parser.add_argument(
        "--no-embed",
        action="store_true",
        help="Use file reference instead of base64 for Markdown.",
    )
    paste_parser.add_argument(
        "--alt", default="image", help="Alt text for the image."
    )

    # -- formats command --
    subparsers.add_parser("formats", help="List supported document formats.")

    args = parser.parse_args(argv)

    if args.command is None:
        parser.print_help()
        return 1

    # Lazy import to avoid loading deps for --help
    from markitwrite import MarkItWrite

    writer = MarkItWrite()

    if args.command == "formats":
        formats = writer.supported_formats()
        if formats:
            print("Supported formats:")
            for fmt in formats:
                print(f"  {fmt}")
        else:
            print("No formats available. Install optional dependencies:")
            print("  pip install 'markitwrite[all]'")
        return 0

    if args.command == "paste":
        # Build position dict
        position = {}
        if args.slide is not None:
            position["slide"] = args.slide
        if args.paragraph is not None:
            position["paragraph"] = args.paragraph

        # Build size dict
        size = {}
        if args.width is not None:
            size["width"] = args.width
        if args.height is not None:
            size["height"] = args.height

        embed = not args.no_embed

        try:
            result = writer.paste(
                image_source=args.image,
                target=args.target,
                target_format=args.target_format,
                position=position or None,
                size=size or None,
                embed=embed,
                alt_text=args.alt,
            )
            print(
                f"Done: pasted image into {args.target} "
                f"({result.target_format}, {len(result.output)} bytes)"
            )
            return 0
        except Exception as e:
            print(f"Error: {e}", file=sys.stderr)
            return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
