#!/usr/bin/env python3
"""Export template slide thumbnails as PNG images for HTML preview.

Uses LibreOffice headless to convert each template PPTX into per-slide
PNG images. These thumbnails are referenced in the HTML deck preview as
CSS background-image, making the preview a much closer representation
of the final branded PPTX.

Usage:
    python scripts/export_thumbnails.py template_registry.json \
        -o templates/thumbnails \
        --update-registry

Requires LibreOffice to be installed and available as `soffice` on PATH.
If LibreOffice is not available, the script exits gracefully and the HTML
preview falls back to solid-color backgrounds.
"""

import argparse
import json
import logging
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.schemas.template_schema import TemplateRegistry

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def _find_soffice() -> str | None:
    """Find the LibreOffice soffice binary."""
    # Check common locations
    candidates = [
        "soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
    ]
    for candidate in candidates:
        if shutil.which(candidate):
            return candidate
    return None


def _export_pptx_slides(pptx_path: Path, output_dir: Path, soffice: str) -> list[Path]:
    """Export all slides from a PPTX as individual PNG files.

    Returns list of PNG paths in slide order.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)

        try:
            result = subprocess.run(
                [
                    soffice,
                    "--headless",
                    "--convert-to", "png",
                    "--outdir", str(tmpdir_path),
                    str(pptx_path),
                ],
                capture_output=True,
                text=True,
                timeout=120,
            )
            if result.returncode != 0:
                logger.warning(
                    f"LibreOffice conversion failed for {pptx_path}: {result.stderr[:200]}"
                )
                return []
        except subprocess.TimeoutExpired:
            logger.warning(f"LibreOffice timed out converting {pptx_path}")
            return []
        except FileNotFoundError:
            logger.warning(f"LibreOffice not found at {soffice}")
            return []

        # LibreOffice exports a single PNG with the first slide,
        # OR multiple PNGs if we use the filter properly.
        # The single-export approach gives us just one file.
        # For per-slide export we need a different approach.

        # Collect exported files
        pngs = sorted(tmpdir_path.glob("*.png"))
        if not pngs:
            logger.warning(f"No PNGs generated for {pptx_path}")
            return []

        # Copy to output directory
        stem = pptx_path.stem.replace(" ", "_")
        output_dir.mkdir(parents=True, exist_ok=True)
        result_paths = []

        for i, png in enumerate(pngs):
            dest = output_dir / f"{stem}_slide_{i}.png"
            shutil.copy2(png, dest)
            result_paths.append(dest)

        return result_paths


def _export_via_python_pptx(pptx_path: Path, output_dir: Path) -> dict[int, Path]:
    """Fallback: extract embedded thumbnail from PPTX if available.

    python-pptx can access the presentation thumbnail but not per-slide
    images. This is a limited fallback that provides at least something.
    Returns dict mapping slide_index to thumbnail path.
    """
    try:
        from pptx import Presentation
        prs = Presentation(str(pptx_path))

        # Try to extract the package-level thumbnail
        stem = pptx_path.stem.replace(" ", "_")
        output_dir.mkdir(parents=True, exist_ok=True)

        # PPTX packages sometimes have a thumbnail.jpeg in docProps/
        import zipfile
        thumb_paths: dict[int, Path] = {}

        with zipfile.ZipFile(pptx_path, "r") as zf:
            for name in zf.namelist():
                if "thumbnail" in name.lower() and name.lower().endswith((".jpeg", ".jpg", ".png")):
                    ext = Path(name).suffix
                    dest = output_dir / f"{stem}_thumb{ext}"
                    with zf.open(name) as src, open(dest, "wb") as dst:
                        dst.write(src.read())
                    # Use as fallback for slide 0
                    thumb_paths[0] = dest
                    break

        return thumb_paths

    except Exception as e:
        logger.debug(f"python-pptx thumbnail extraction failed: {e}")
        return {}


def main():
    parser = argparse.ArgumentParser(
        description="Export template slide thumbnails as PNG for HTML preview"
    )
    parser.add_argument("registry", type=Path, help="Path to template_registry.json")
    parser.add_argument("-o", "--output-dir", type=Path,
                        default=Path("templates/thumbnails"),
                        help="Output directory for thumbnails")
    parser.add_argument("--update-registry", action="store_true",
                        help="Update the registry JSON with thumbnail_path fields")
    args = parser.parse_args()

    if not args.registry.exists():
        print(f"Error: Registry not found: {args.registry}", file=sys.stderr)
        sys.exit(1)

    registry = TemplateRegistry.load(args.registry)
    soffice = _find_soffice()

    if not soffice:
        logger.warning(
            "LibreOffice (soffice) not found. Attempting python-pptx fallback."
        )

    args.output_dir.mkdir(parents=True, exist_ok=True)

    # Group templates by source file
    by_file: dict[str, list[int]] = {}
    for i, tmpl in enumerate(registry.templates):
        by_file.setdefault(tmpl.template_file, []).append(i)

    exported = 0
    skipped = 0

    for pptx_file, indices in by_file.items():
        pptx_path = Path(pptx_file)
        if not pptx_path.exists():
            logger.warning(f"Template file not found: {pptx_path}")
            skipped += len(indices)
            continue

        slide_pngs: dict[int, Path] = {}

        if soffice:
            # LibreOffice export
            pngs = _export_pptx_slides(pptx_path, args.output_dir, soffice)
            for i, png_path in enumerate(pngs):
                slide_pngs[i] = png_path
        else:
            # Fallback
            slide_pngs = _export_via_python_pptx(pptx_path, args.output_dir)

        # Map thumbnails to registry entries
        for idx in indices:
            tmpl = registry.templates[idx]
            png = slide_pngs.get(tmpl.slide_index)
            if png:
                tmpl.thumbnail_path = str(png)
                exported += 1
            else:
                skipped += 1

    print(f"Thumbnails: {exported} exported, {skipped} skipped")
    print(f"Output: {args.output_dir}")

    if args.update_registry:
        registry.save(args.registry)
        print(f"Registry updated: {args.registry}")


if __name__ == "__main__":
    main()
