"""Document preview CLI - Render docx to images using OnlyOffice AppImage.

Standalone CLI tool for visual preview. No MCP dependency.

Usage:
    docx-preview input.docx                    # Preview page 1
    docx-preview input.docx --pages 1-3        # Preview pages 1 to 3
    docx-preview input.docx --pages 2,5,8      # Preview specific pages
    docx-preview input.docx --pages all        # Preview all pages
    docx-preview input.docx --dpi 300          # High resolution
    docx-preview --cleanup                     # Remove temp preview files

Dependencies:
    - OnlyOffice DesktopEditors AppImage (docx -> pdf)
    - ImageMagick convert (pdf -> png per page)

Environment:
    ONLYOFFICE_APPIMAGE  Path to OnlyOffice AppImage (auto-detected if not set)
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import List, Optional

# Preview output directory
PREVIEW_DIR = "/tmp/docx_preview"

# Default paths to search for OnlyOffice AppImage
ONLYOFFICE_SEARCH_PATHS = [
    os.path.expanduser("~/DesktopEditors.AppImage"),
    os.path.expanduser("~/Applications/DesktopEditors.AppImage"),
    "/opt/onlyoffice/DesktopEditors.AppImage",
    "/usr/local/bin/DesktopEditors",
]


def find_onlyoffice() -> Optional[str]:
    """Find OnlyOffice AppImage on the system."""
    env_path = os.environ.get("ONLYOFFICE_APPIMAGE")
    if env_path and os.path.isfile(env_path) and os.access(env_path, os.X_OK):
        return env_path

    for path in ONLYOFFICE_SEARCH_PATHS:
        if path and os.path.isfile(path) and os.access(path, os.X_OK):
            return path

    return shutil.which("DesktopEditors")


def find_imagemagick() -> Optional[str]:
    """Find ImageMagick convert command."""
    for cmd in ["magick", "convert"]:
        result = shutil.which(cmd)
        if result:
            return result
    return None


def parse_pages(pages: str) -> List[int]:
    """Parse page specification into list of 0-based page indices.

    Examples:
        "1"       -> [0]
        "1-3"     -> [0, 1, 2]
        "2,5,8"   -> [1, 4, 7]
        "all"     -> handled separately by caller
    """
    if pages.lower() == "all":
        return []  # Empty means all pages

    result = []
    for part in pages.split(","):
        part = part.strip()
        if "-" in part:
            start, end = part.split("-", 1)
            result.extend(range(int(start.strip()) - 1, int(end.strip())))
        else:
            result.append(int(part.strip()) - 1)

    return sorted(set(result))


def convert_docx_to_pdf(onlyoffice_path: str, docx_path: str, output_dir: str) -> str:
    """Convert docx to pdf using OnlyOffice."""
    cmd = [
        onlyoffice_path,
        "--convert-to", "pdf",
        "--output-dir", output_dir,
        docx_path
    ]

    env = {**os.environ, "DISPLAY": os.environ.get("DISPLAY", ":0")}

    result = subprocess.run(cmd, capture_output=True, text=True, timeout=60, env=env)

    if result.returncode != 0:
        print(f"ERROR: OnlyOffice conversion failed (exit {result.returncode})", file=sys.stderr)
        if result.stderr:
            print(result.stderr, file=sys.stderr)
        sys.exit(1)

    # Find generated PDF
    pdf_name = Path(docx_path).stem + ".pdf"
    pdf_path = os.path.join(output_dir, pdf_name)
    if os.path.exists(pdf_path):
        return pdf_path

    for f in Path(output_dir).glob("*.pdf"):
        return str(f)

    print("ERROR: PDF not found after conversion", file=sys.stderr)
    sys.exit(1)


def convert_pdf_to_images(
    imagemagick_path: str,
    pdf_path: str,
    output_dir: str,
    pages: List[int],
    dpi: int = 200
) -> List[str]:
    """Convert specific PDF pages to PNG images."""
    image_paths = []

    # If pages is empty, convert all pages
    if not pages:
        # Get page count via identify
        try:
            result = subprocess.run(
                [imagemagick_path, "identify", pdf_path] if os.path.basename(imagemagick_path) == "magick"
                else ["identify", pdf_path],
                capture_output=True, text=True, timeout=30
            )
            page_count = len(result.stdout.strip().split("\n")) if result.stdout.strip() else 1
            pages = list(range(page_count))
        except Exception:
            pages = list(range(20))  # Fallback: try up to 20 pages

    for page_idx in pages:
        output_path = os.path.join(output_dir, f"page-{page_idx + 1:03d}.png")

        cmd = [imagemagick_path]
        if os.path.basename(imagemagick_path) == "magick":
            cmd.append("convert")

        cmd.extend([
            "-density", str(dpi),
            "-background", "white",
            "-alpha", "remove",
            "-alpha", "off",
            f"{pdf_path}[{page_idx}]",
            "-quality", "95",
            output_path
        ])

        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            if result.returncode == 0 and os.path.exists(output_path):
                image_paths.append(output_path)
            else:
                # Page doesn't exist, stop
                break
        except subprocess.TimeoutExpired:
            break

    return image_paths


def cleanup():
    """Remove all preview files."""
    if not os.path.exists(PREVIEW_DIR):
        print("No preview files to clean up.")
        return

    removed = 0
    for f in Path(PREVIEW_DIR).glob("*.png"):
        f.unlink()
        removed += 1

    print(f"Removed {removed} preview files from {PREVIEW_DIR}")


def preview(docx_path: str, pages: str = "1", dpi: int = 200) -> List[str]:
    """Main preview function. Returns list of image paths."""
    # Validate input
    if not os.path.isfile(docx_path):
        print(f"ERROR: File not found: {docx_path}", file=sys.stderr)
        sys.exit(1)

    # Find dependencies
    onlyoffice = find_onlyoffice()
    if not onlyoffice:
        print(
            "ERROR: OnlyOffice not found.\n"
            "Install OnlyOffice DesktopEditors AppImage and either:\n"
            "  - Set ONLYOFFICE_APPIMAGE environment variable\n"
            "  - Place at ~/DesktopEditors.AppImage",
            file=sys.stderr
        )
        sys.exit(1)

    imagemagick = find_imagemagick()
    if not imagemagick:
        print("ERROR: ImageMagick not found. Install: sudo apt install imagemagick", file=sys.stderr)
        sys.exit(1)

    os.makedirs(PREVIEW_DIR, exist_ok=True)
    page_indices = parse_pages(pages)

    # Convert docx -> pdf in temp dir
    with tempfile.TemporaryDirectory() as pdf_dir:
        pdf_path = convert_docx_to_pdf(onlyoffice, docx_path, pdf_dir)
        image_paths = convert_pdf_to_images(imagemagick, pdf_path, PREVIEW_DIR, page_indices, dpi)

    return image_paths


def main():
    parser = argparse.ArgumentParser(
        description="Preview Word documents as images using OnlyOffice",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  docx-preview document.docx                  # Preview page 1
  docx-preview document.docx --pages 1-3      # Pages 1 to 3
  docx-preview document.docx --pages 2,5,8    # Specific pages
  docx-preview document.docx --pages all      # All pages
  docx-preview document.docx --dpi 300        # High resolution
  docx-preview --cleanup                      # Clean temp files
"""
    )

    parser.add_argument("docx_path", nargs="?", help="Path to .docx file")
    parser.add_argument("--pages", "-p", default="1", help="Pages to preview: '1', '1-3', '2,5,8', 'all' (default: 1)")
    parser.add_argument("--dpi", "-d", type=int, default=200, help="Resolution in DPI (default: 200)")
    parser.add_argument("--cleanup", action="store_true", help="Remove all preview files")

    args = parser.parse_args()

    if args.cleanup:
        cleanup()
        return

    if not args.docx_path:
        parser.error("docx_path is required (unless using --cleanup)")

    image_paths = preview(args.docx_path, args.pages, args.dpi)

    if not image_paths:
        print("No pages rendered. Document may be empty or page numbers out of range.", file=sys.stderr)
        sys.exit(1)

    # Output image paths (one per line, for easy scripting)
    for path in image_paths:
        print(path)


if __name__ == "__main__":
    main()
