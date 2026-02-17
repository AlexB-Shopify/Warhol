#!/usr/bin/env python3
"""Compact a PPTX file for cross-platform compatibility (Google Slides, Keynote).

Reduces file size by stripping embedded fonts and compressing oversized images,
while preserving the original template's layout/master structure intact.

The key insight: Google Slides is strict about the OPC package structure.
The original template structure (from Google Slides exports) is valid, but
clone-imported masters/layouts from our build pipeline are not. This script
avoids modifying the layout/master chain entirely — it only touches media
and fonts, which are safe to modify.

Usage:
    python scripts/repair_pptx.py input.pptx [-o output.pptx]

If no -o is given, the input file is overwritten in place.
"""

import argparse
import hashlib
import logging
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def repair_pptx(input_path: Path, output_path: Path, compress_images: bool = True) -> dict:
    """Compact a PPTX for cross-platform compatibility.

    Operations (all safe — no layout/master changes):
    1. Strip embedded fonts (.fntdata) and embedTrueTypeFonts attribute
    2. Convert animated GIFs to static PNG first frames
    3. Downscale oversized images
    4. Deduplicate identical media files
    5. Inject missing docProps (required by Google Slides)
    6. Ensure notesMasterIdLst exists (required by Keynote)

    Returns a dict of stats.
    """
    stats = {
        "original_size_mb": input_path.stat().st_size / (1024 * 1024),
        "fonts_removed": 0,
        "media_compressed": 0,
        "media_deduplicated": 0,
        "media_saved_mb": 0.0,
    }

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir) / "repaired.pptx"

        with zipfile.ZipFile(input_path, "r") as zf_in:
            all_names = set(zf_in.namelist())

            # -------------------------------------------------------
            # Phase 1: Identify fonts to strip
            # -------------------------------------------------------
            font_files = {n for n in all_names if n.endswith(".fntdata")}
            stats["fonts_removed"] = len(font_files)

            # -------------------------------------------------------
            # Phase 2: Compress images
            # -------------------------------------------------------
            media_replacements: dict[str, tuple[bytes, str]] = {}
            media_ext_changes: dict[str, str] = {}

            if compress_images:
                try:
                    media_replacements, media_ext_changes, saved_mb = (
                        _compress_media(zf_in, all_names)
                    )
                    stats["media_compressed"] = len(media_replacements)
                    stats["media_saved_mb"] = saved_mb
                except ImportError:
                    logger.warning("Pillow not installed — skipping image compression")
                except Exception as e:
                    logger.warning(f"Image compression failed: {e}")

            # -------------------------------------------------------
            # Phase 3: Deduplicate media by content hash
            # -------------------------------------------------------
            kept_media = sorted(
                n for n in all_names if n.startswith("ppt/media/")
            )
            hash_to_canonical: dict[str, str] = {}
            media_dedup_remap: dict[str, str] = {}  # old_filename → canonical_filename
            media_to_skip: set[str] = set()

            for media_name in kept_media:
                if media_name in media_replacements:
                    blob = media_replacements[media_name][0]
                else:
                    blob = zf_in.read(media_name)
                h = hashlib.sha256(blob).hexdigest()
                if h in hash_to_canonical:
                    canonical = hash_to_canonical[h]
                    old_fn = media_name.split("/")[-1]
                    can_fn = canonical.split("/")[-1]
                    if old_fn != can_fn:
                        media_dedup_remap[old_fn] = can_fn
                    media_to_skip.add(media_name)
                    stats["media_deduplicated"] += 1
                else:
                    hash_to_canonical[h] = media_name

            # Merge extension changes and dedup remaps for rels patching
            all_filename_remap = {}
            all_filename_remap.update(media_ext_changes)
            all_filename_remap.update(media_dedup_remap)

            # -------------------------------------------------------
            # Phase 4: Detect missing parts
            # -------------------------------------------------------
            needs_docprops = "docProps/core.xml" not in all_names
            
            # Check for notesMaster in rels but not in presentation.xml
            pres_xml_text = zf_in.read("ppt/presentation.xml").decode("utf-8")
            pres_rels_text = zf_in.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
            has_nm_rel = "notesMaster" in pres_rels_text
            has_nm_id = "notesMasterIdLst" in pres_xml_text
            needs_notes_fix = has_nm_rel and not has_nm_id

            # -------------------------------------------------------
            # Phase 5: Rewrite the PPTX
            # -------------------------------------------------------
            with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
                for name in all_names:
                    # Skip fonts
                    if name in font_files:
                        continue
                    # Skip deduplicated media
                    if name in media_to_skip:
                        continue

                    data = zf_in.read(name)

                    # --- Strip font references from presentation.xml ---
                    if name == "ppt/presentation.xml":
                        text = data.decode("utf-8")
                        text = text.replace(' embedTrueTypeFonts="1"', "")
                        # Fix notesMasterIdLst if needed
                        if needs_notes_fix:
                            nm_rid = re.search(
                                r'Id="([^"]+)"[^>]*notesMaster', pres_rels_text
                            )
                            if nm_rid:
                                rid = nm_rid.group(1)
                                notes_el = (
                                    f"<p:notesMasterIdLst>"
                                    f'<p:notesMasterId r:id="{rid}"/>'
                                    f"</p:notesMasterIdLst>"
                                )
                                text = text.replace(
                                    "</p:sldMasterIdLst>",
                                    f"</p:sldMasterIdLst>{notes_el}",
                                )
                                logger.info(f"Fixed notesMasterIdLst (r:id={rid})")
                        data = text.encode("utf-8")

                    # --- Strip font rels from presentation.xml.rels ---
                    if name == "ppt/_rels/presentation.xml.rels":
                        text = data.decode("utf-8")
                        text = re.sub(r"<Relationship[^>]*/font[^>]*/>", "", text)
                        data = text.encode("utf-8")

                    # --- Strip fntdata from [Content_Types].xml ---
                    if name == "[Content_Types].xml":
                        text = data.decode("utf-8")
                        text = re.sub(
                            r'<Default Extension="fntdata"[^/]*/>', "", text
                        )
                        # Add docProps content types if needed
                        if needs_docprops and "docProps/core.xml" not in text:
                            text = text.replace(
                                "</Types>",
                                '<Override PartName="/docProps/core.xml"'
                                ' ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                                '<Override PartName="/docProps/app.xml"'
                                ' ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
                                "</Types>",
                            )
                        # Update content types for GIF→PNG conversions
                        for old_fn, new_fn in media_ext_changes.items():
                            text = text.replace(f"media/{old_fn}", f"media/{new_fn}")
                            # Fix content type if extension changed
                            old_ext = old_fn.rsplit(".", 1)[-1]
                            new_ext = new_fn.rsplit(".", 1)[-1]
                            if old_ext != new_ext:
                                ct_map = {"gif": "image/gif", "png": "image/png"}
                                old_ct = ct_map.get(old_ext, f"image/{old_ext}")
                                new_ct = ct_map.get(new_ext, f"image/{new_ext}")
                                # Only replace for this specific file's Override
                                pattern = (
                                    r'(<Override[^>]*' + re.escape(new_fn)
                                    + r'[^>]*ContentType=")' + re.escape(old_ct) + '"'
                                )
                                text = re.sub(pattern, r"\g<1>" + new_ct + '"', text)
                        data = text.encode("utf-8")

                    # --- Inject docProps into _rels/.rels ---
                    if name == "_rels/.rels" and needs_docprops:
                        data = (
                            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                            '<Relationship Id="rId1"'
                            ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
                            ' Target="ppt/presentation.xml"/>'
                            '<Relationship Id="rId2"'
                            ' Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"'
                            ' Target="docProps/core.xml"/>'
                            '<Relationship Id="rId3"'
                            ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"'
                            ' Target="docProps/app.xml"/>'
                            "</Relationships>"
                        ).encode("utf-8")

                    # --- Remap media filenames in .rels files ---
                    if all_filename_remap and name.endswith(".rels"):
                        text = data.decode("utf-8")
                        changed = False
                        for old_fn, new_fn in all_filename_remap.items():
                            old_ref = f"media/{old_fn}"
                            new_ref = f"media/{new_fn}"
                            if old_ref in text:
                                text = text.replace(old_ref, new_ref)
                                changed = True
                        if changed:
                            data = text.encode("utf-8")

                    # --- Write (compressed media or original) ---
                    if name in media_replacements:
                        new_data, new_ext = media_replacements[name]
                        old_fn = name.split("/")[-1]
                        new_fn = media_ext_changes.get(old_fn, old_fn)
                        new_name = name.rsplit("/", 1)[0] + "/" + new_fn
                        zf_out.writestr(new_name, new_data)
                    else:
                        zf_out.writestr(name, data)

                # --- Write docProps if missing ---
                if needs_docprops:
                    zf_out.writestr(
                        "docProps/core.xml",
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                        '<cp:coreProperties'
                        ' xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
                        ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
                        ' xmlns:dcterms="http://purl.org/dc/terms/"'
                        ' xmlns:dcmitype="http://purl.org/dc/dcmitype/"'
                        ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                        "<dc:title>Presentation</dc:title>"
                        '<dcterms:created xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:created>'
                        '<dcterms:modified xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:modified>'
                        "</cp:coreProperties>",
                    )
                    zf_out.writestr(
                        "docProps/app.xml",
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                        "<Properties"
                        ' xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
                        ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
                        "<Application>Warhol</Application>"
                        "</Properties>",
                    )
                    logger.info("Injected missing docProps")

        shutil.copy2(tmp_path, output_path)

    stats["final_size_mb"] = output_path.stat().st_size / (1024 * 1024)
    return stats


# ---------------------------------------------------------------------------
# Image compression
# ---------------------------------------------------------------------------

_MAX_IMAGE_DIM = 1920


def _compress_media(
    zf: zipfile.ZipFile,
    all_names: set[str],
) -> tuple[dict[str, tuple[bytes, str]], dict[str, str], float]:
    """Compress media: animated GIFs → static PNG, downscale oversized images."""
    from io import BytesIO
    from PIL import Image

    media_replacements: dict[str, tuple[bytes, str]] = {}
    filename_changes: dict[str, str] = {}
    total_saved = 0

    # Track all media names to avoid collisions when converting GIF→PNG
    existing_media_names = {
        n.split("/")[-1] for n in all_names if n.startswith("ppt/media/")
    }

    for media_path in sorted(n for n in all_names if n.startswith("ppt/media/")):
        info = zf.getinfo(media_path)
        original_size = info.file_size
        ext = media_path.rsplit(".", 1)[-1].lower()
        filename = media_path.split("/")[-1]
        blob = zf.read(media_path)

        if ext == "gif" and original_size > 100_000:
            try:
                img = Image.open(BytesIO(blob)).convert("RGBA")
                w, h = img.size
                if max(w, h) > _MAX_IMAGE_DIM:
                    ratio = _MAX_IMAGE_DIM / max(w, h)
                    img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)
                buf = BytesIO()
                img.save(buf, format="PNG", optimize=True)
                new_data = buf.getvalue()
                if len(new_data) < original_size:
                    # Pick a unique .png name (avoid collisions)
                    base_name = filename.rsplit(".", 1)[0]
                    new_fn = base_name + ".png"
                    suffix = 1
                    while new_fn in existing_media_names:
                        new_fn = f"{base_name}_c{suffix}.png"
                        suffix += 1
                    existing_media_names.add(new_fn)
                    media_replacements[media_path] = (new_data, "png")
                    filename_changes[filename] = new_fn
                    total_saved += original_size - len(new_data)
            except Exception:
                pass
            continue

        if ext == "png" and original_size > 500_000:
            try:
                img = Image.open(BytesIO(blob))
                w, h = img.size
                if max(w, h) > _MAX_IMAGE_DIM:
                    ratio = _MAX_IMAGE_DIM / max(w, h)
                    img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)
                    buf = BytesIO()
                    img.save(buf, format="PNG", optimize=True)
                    new_data = buf.getvalue()
                    if len(new_data) < original_size * 0.8:
                        media_replacements[media_path] = (new_data, "png")
                        total_saved += original_size - len(new_data)
            except Exception:
                pass

    return media_replacements, filename_changes, total_saved / (1024 * 1024)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Compact a PPTX for Google Slides / Keynote compatibility"
    )
    parser.add_argument("input_file", type=Path, help="Input PPTX file")
    parser.add_argument(
        "-o", "--output", type=Path, default=None,
        help="Output PPTX path (default: overwrite input)",
    )
    args = parser.parse_args()

    if not args.input_file.exists():
        print(f"Error: file not found: {args.input_file}", file=sys.stderr)
        sys.exit(1)

    output_path = args.output or args.input_file

    print(f"Repairing: {args.input_file}")
    stats = repair_pptx(args.input_file, output_path)

    savings_pct = (
        (1 - stats["final_size_mb"] / stats["original_size_mb"]) * 100
        if stats["original_size_mb"] > 0 else 0
    )
    print(f"  Original:       {stats['original_size_mb']:.1f} MB")
    print(f"  Final:          {stats['final_size_mb']:.1f} MB  ({savings_pct:.0f}% smaller)")
    print(f"  Fonts stripped:  {stats['fonts_removed']}")
    print(f"  Images compressed: {stats['media_compressed']}  ({stats['media_saved_mb']:.1f} MB)")
    print(f"  Images deduped:  {stats['media_deduplicated']}")
    print(f"")
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    main()
