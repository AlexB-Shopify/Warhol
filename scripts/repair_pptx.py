#!/usr/bin/env python3
"""Repair and compact a PPTX file for cross-platform compatibility.

Strips unused slide layouts, masters, themes, and media to produce a clean
OPC package that opens reliably in Google Slides, Keynote, and LibreOffice.

The primary problem: when Warhol builds a PPTX from a large template bank,
the base template's layouts/masters/media carry over even when no slide
references them.  A 30-slide deck can balloon to 400+ MB with 400+ unused
layouts.  This script trims it down to only the parts actually used.

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
    """Repair a PPTX file by stripping unused parts, deduplicating and compressing media.

    Args:
        input_path: Source PPTX file.
        output_path: Where to write the repaired file.
        compress_images: If True, convert animated GIFs to static PNGs and
            downscale oversized images.  Requires Pillow.

    Returns a dict of stats about what was removed.
    """
    stats = {
        "original_size_mb": input_path.stat().st_size / (1024 * 1024),
        "layouts_removed": 0,
        "masters_removed": 0,
        "themes_removed": 0,
        "media_removed": 0,
        "media_deduplicated": 0,
        "media_compressed": 0,
        "media_saved_mb": 0.0,
        "content_types_cleaned": 0,
    }

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir) / "repaired.pptx"

        with zipfile.ZipFile(input_path, "r") as zf_in:
            all_names = set(zf_in.namelist())

            # ===========================================================
            # Phase 1: Walk the full reference chain to find USED parts
            # ===========================================================
            # Strategy: start from slides, walk outward through rels to
            # discover every layout, master, theme, and media that is
            # transitively referenced.

            used_files: set[str] = set()

            # --- Always keep structural / root files ---
            structural = {
                "[Content_Types].xml",
                "_rels/.rels",
                "ppt/presentation.xml",
                "ppt/_rels/presentation.xml.rels",
                "ppt/presProps.xml",
                "ppt/tableStyles.xml",
                "ppt/viewProps.xml",
                "docProps/app.xml",
                "docProps/core.xml",
                "docProps/thumbnail.jpeg",
            }
            for s in structural:
                if s in all_names:
                    used_files.add(s)

            # --- Find slides from presentation.xml.rels ---
            pres_rels_content = _read(zf_in, "ppt/_rels/presentation.xml.rels")
            slide_parts = set()
            for target, reltype in _parse_rels(pres_rels_content):
                resolved = _resolve("ppt/presentation.xml", target)
                if "slides/slide" in resolved:
                    slide_parts.add(resolved)

            # --- Walk from each slide through the full tree ---
            # Collect: layouts used by slides
            used_layouts: set[str] = set()
            for slide in sorted(slide_parts):
                used_files.add(slide)
                _add_rels_file(slide, all_names, used_files)
                for target, reltype in _get_rels(zf_in, slide, all_names):
                    resolved = _resolve(slide, target)
                    if "slideLayouts/" in resolved:
                        used_layouts.add(resolved)
                    else:
                        _mark_used(resolved, all_names, used_files)

            # Collect: masters used by those layouts
            used_masters: set[str] = set()
            for layout in sorted(used_layouts):
                used_files.add(layout)
                _add_rels_file(layout, all_names, used_files)
                for target, reltype in _get_rels(zf_in, layout, all_names):
                    resolved = _resolve(layout, target)
                    if "slideMasters/" in resolved:
                        used_masters.add(resolved)
                    else:
                        _mark_used(resolved, all_names, used_files)

            # Collect: themes used by those masters
            used_themes: set[str] = set()
            for master in sorted(used_masters):
                used_files.add(master)
                _add_rels_file(master, all_names, used_files)
                for target, reltype in _get_rels(zf_in, master, all_names):
                    resolved = _resolve(master, target)
                    if "theme/" in resolved:
                        used_themes.add(resolved)
                    else:
                        _mark_used(resolved, all_names, used_files)

            # Walk themes for their media
            for theme in sorted(used_themes):
                used_files.add(theme)
                _add_rels_file(theme, all_names, used_files)
                for target, reltype in _get_rels(zf_in, theme, all_names):
                    resolved = _resolve(theme, target)
                    _mark_used(resolved, all_names, used_files)

            # --- Notes slides/masters (keep all that exist) ---
            for name in all_names:
                if name.startswith("ppt/notesSlides/") or name.startswith("ppt/notesMasters/"):
                    used_files.add(name)
                    # Walk their rels too
                    for target, reltype in _get_rels(zf_in, name, all_names):
                        resolved = _resolve(name, target)
                        _mark_used(resolved, all_names, used_files)
                        # If it's a theme ref, keep it
                        if "theme/" in resolved:
                            used_themes.add(resolved)
                            used_files.add(resolved)
                            _add_rels_file(resolved, all_names, used_files)

            # ===========================================================
            # Phase 2: Determine files to REMOVE
            # ===========================================================
            to_remove: set[str] = set()

            for name in all_names:
                if name in used_files:
                    continue

                # Only remove files in specific ppt/ subdirectories
                if any(name.startswith(prefix) for prefix in (
                    "ppt/slideLayouts/",
                    "ppt/slideMasters/",
                    "ppt/theme/",
                    "ppt/media/",
                )):
                    to_remove.add(name)

            # Count by category
            for name in to_remove:
                if "slideLayouts/slideLayout" in name and name.endswith(".xml"):
                    stats["layouts_removed"] += 1
                elif "slideMasters/slideMaster" in name and name.endswith(".xml"):
                    stats["masters_removed"] += 1
                elif name.startswith("ppt/theme/") and name.endswith(".xml"):
                    stats["themes_removed"] += 1
                elif name.startswith("ppt/media/"):
                    stats["media_removed"] += 1

            logger.info(
                f"Removing {len(to_remove)} unused files: "
                f"{stats['layouts_removed']} layouts, "
                f"{stats['masters_removed']} masters, "
                f"{stats['themes_removed']} themes, "
                f"{stats['media_removed']} media"
            )

            # ===========================================================
            # Phase 3: Deduplicate media (same content, different names)
            # ===========================================================
            kept_media = sorted(
                n for n in all_names - to_remove
                if n.startswith("ppt/media/")
            )

            hash_to_canonical: dict[str, str] = {}
            media_remap: dict[str, str] = {}  # old_path → canonical_path

            for media_name in kept_media:
                blob = zf_in.read(media_name)
                h = hashlib.sha256(blob).hexdigest()
                if h in hash_to_canonical:
                    canonical = hash_to_canonical[h]
                    media_remap[media_name] = canonical
                    to_remove.add(media_name)
                    stats["media_deduplicated"] += 1
                    stats["media_removed"] += 1
                else:
                    hash_to_canonical[h] = media_name

            if media_remap:
                logger.info(f"Deduplicated {len(media_remap)} media files")

            # Build filename-level remap for rels patching
            # e.g. "image664.gif" → "image29.gif"
            filename_remap: dict[str, str] = {}
            for dup_path, canonical_path in media_remap.items():
                dup_name = dup_path.split("/")[-1]
                canonical_name = canonical_path.split("/")[-1]
                if dup_name != canonical_name:
                    filename_remap[dup_name] = canonical_name

            # ===========================================================
            # Phase 3b: Compress images (animated GIF → static PNG, downscale)
            # ===========================================================
            # media_replacements: maps original media path → (new_data, new_ext)
            media_replacements: dict[str, tuple[bytes, str]] = {}
            media_ext_changes: dict[str, str] = {}  # old_filename → new_filename

            if compress_images:
                try:
                    media_replacements, media_ext_changes, saved_mb = (
                        _compress_media(zf_in, all_names - to_remove)
                    )
                    stats["media_compressed"] = len(media_replacements)
                    stats["media_saved_mb"] = saved_mb
                    if media_replacements:
                        logger.info(
                            f"Compressed {len(media_replacements)} images, "
                            f"saved {saved_mb:.1f} MB"
                        )
                except ImportError:
                    logger.warning(
                        "Pillow not installed — skipping image compression. "
                        "Install with: pip install Pillow"
                    )
                except Exception as e:
                    logger.warning(f"Image compression failed: {e}")

            # Merge ext changes into the filename remap for rels patching
            filename_remap.update(media_ext_changes)

            # ===========================================================
            # Phase 3c: Determine if docProps need to be injected
            # ===========================================================
            needs_docprops = (
                "docProps/core.xml" not in (all_names - to_remove)
                or "docProps/app.xml" not in (all_names - to_remove)
            )
            if needs_docprops:
                logger.info("Injecting missing docProps (required by Google Slides)")

            # ===========================================================
            # Phase 4: Rewrite the PPTX
            # ===========================================================
            removed_set = to_remove  # alias for clarity

            with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
                for name in sorted(all_names):
                    if name in removed_set:
                        continue

                    data = zf_in.read(name)

                    # --- Patch [Content_Types].xml ---
                    if name == "[Content_Types].xml":
                        data = _patch_content_types(data, removed_set)
                        # Count how many were cleaned
                        orig_ct = zf_in.read(name).decode()
                        new_ct = data.decode()
                        orig_count = orig_ct.count("<Override")
                        new_count = new_ct.count("<Override")
                        stats["content_types_cleaned"] = orig_count - new_count

                    # --- Patch presentation.xml ---
                    if name == "ppt/presentation.xml":
                        data = _patch_presentation_xml(
                            data, zf_in, all_names, used_masters
                        )

                    # --- Patch presentation.xml.rels ---
                    if name == "ppt/_rels/presentation.xml.rels":
                        data = _patch_rels_remove_targets(
                            data, removed_set, "ppt/"
                        )

                    # --- Patch master rels (remove refs to deleted layouts) ---
                    if (name.startswith("ppt/slideMasters/_rels/")
                            and name.endswith(".rels")):
                        data = _patch_rels_remove_targets(
                            data, removed_set, "ppt/slideMasters/"
                        )
                        # Also patch master XML sldLayoutIdLst
                        master_xml_name = name.replace("/_rels/", "/").replace(".rels", "")
                        if master_xml_name in all_names and master_xml_name not in removed_set:
                            # We'll patch the master XML below
                            pass

                    # --- Patch master XMLs (remove sldLayoutIdLst entries) ---
                    if (name.startswith("ppt/slideMasters/slideMaster")
                            and name.endswith(".xml")
                            and name in used_masters):
                        data = _patch_master_layout_ids(
                            data, zf_in, name, all_names, removed_set
                        )

                    # --- Remap media refs in ALL .rels files ---
                    if filename_remap and name.endswith(".rels"):
                        data = _remap_media_in_rels(data, filename_remap)

                    # --- Patch [Content_Types] for GIF→PNG conversions ---
                    if name == "[Content_Types].xml" and media_ext_changes:
                        data = _patch_content_types_extensions(
                            data, media_ext_changes
                        )

                    # --- Inject docProps into _rels/.rels ---
                    if name == "_rels/.rels" and needs_docprops:
                        data = _inject_docprops_rels(data)

                    # --- Inject docProps into [Content_Types].xml ---
                    if name == "[Content_Types].xml" and needs_docprops:
                        data = _inject_docprops_content_types(data)

                    # --- Write compressed media or original ---
                    if name in media_replacements:
                        new_data, new_ext = media_replacements[name]
                        old_filename = name.split("/")[-1]
                        new_filename = media_ext_changes.get(
                            old_filename, old_filename
                        )
                        new_name = name.rsplit("/", 1)[0] + "/" + new_filename
                        zf_out.writestr(new_name, new_data)
                    else:
                        zf_out.writestr(name, data)

                # --- Write injected docProps files ---
                if needs_docprops:
                    _write_docprops(zf_out)

        # Copy result to output
        shutil.copy2(tmp_path, output_path)

    stats["final_size_mb"] = output_path.stat().st_size / (1024 * 1024)
    return stats


# ---------------------------------------------------------------------------
# Reference chain helpers
# ---------------------------------------------------------------------------

def _read(zf: zipfile.ZipFile, path: str) -> str:
    """Read a text file from the ZIP, returning '' if missing."""
    try:
        return zf.read(path).decode("utf-8")
    except KeyError:
        return ""


def _rels_path_for(part_path: str) -> str:
    """Get the .rels file path for a given part.

    e.g. ppt/slides/slide1.xml → ppt/slides/_rels/slide1.xml.rels
    """
    parts = part_path.rsplit("/", 1)
    if len(parts) == 2:
        return f"{parts[0]}/_rels/{parts[1]}.rels"
    return f"_rels/{parts[0]}.rels"


def _add_rels_file(part_path: str, all_names: set[str], used: set[str]) -> None:
    """Add the .rels file for a part to the used set."""
    rp = _rels_path_for(part_path)
    if rp in all_names:
        used.add(rp)


def _get_rels(zf: zipfile.ZipFile, part_path: str, all_names: set[str]) -> list[tuple[str, str]]:
    """Get all (target, reltype) pairs from a part's .rels file."""
    rp = _rels_path_for(part_path)
    if rp not in all_names:
        return []
    content = _read(zf, rp)
    return _parse_rels(content)


def _parse_rels(content: str) -> list[tuple[str, str]]:
    """Parse a .rels XML string into (Target, Type) pairs.

    Only returns internal targets (skips external/http).
    """
    results = []
    for m in re.finditer(r"<Relationship\s([^>]+?)/>", content, re.DOTALL):
        attrs = m.group(1)
        target_m = re.search(r'Target="([^"]+)"', attrs)
        type_m = re.search(r'Type="([^"]+)"', attrs)
        if not target_m:
            continue
        target = target_m.group(1)
        reltype = type_m.group(1) if type_m else ""

        # Skip external
        if target.startswith("http://") or target.startswith("https://"):
            continue
        if 'TargetMode="External"' in attrs:
            continue

        results.append((target, reltype))
    return results


def _resolve(source_part: str, rel_target: str) -> str:
    """Resolve a relative Target from a .rels file to an absolute package path.

    The source_part is the part that OWNS the .rels file (not the rels file itself).
    e.g. source=ppt/slides/slide1.xml, target=../slideLayouts/slideLayout5.xml
    → ppt/slideLayouts/slideLayout5.xml
    """
    # Start from the directory containing the source part
    parts = source_part.split("/")[:-1]

    for seg in rel_target.split("/"):
        if seg == "..":
            if parts:
                parts.pop()
        elif seg and seg != ".":
            parts.append(seg)

    return "/".join(parts)


def _mark_used(resolved: str, all_names: set[str], used: set[str]) -> None:
    """Mark a resolved path as used (if it exists in the package)."""
    if resolved in all_names:
        used.add(resolved)


# ---------------------------------------------------------------------------
# Patching functions
# ---------------------------------------------------------------------------

def _patch_content_types(data: bytes, removed: set[str]) -> bytes:
    """Remove Override entries from [Content_Types].xml for deleted parts."""
    text = data.decode("utf-8")

    def _filter_override(m: re.Match) -> str:
        partname_m = re.search(r'PartName="([^"]+)"', m.group(0))
        if partname_m:
            pn = partname_m.group(1).lstrip("/")
            if pn in removed:
                return ""
        return m.group(0)

    text = re.sub(r"<Override\s[^>]+/>", _filter_override, text)

    # Clean up whitespace from removed entries
    text = re.sub(r"\n\s*\n", "\n", text)
    return text.encode("utf-8")


def _patch_presentation_xml(
    data: bytes,
    zf: zipfile.ZipFile,
    all_names: set[str],
    used_masters: set[str],
) -> bytes:
    """Remove sldMasterId entries for masters that were removed."""
    text = data.decode("utf-8")

    # Map rId → master path from presentation.xml.rels
    pres_rels = _read(zf, "ppt/_rels/presentation.xml.rels")
    rid_to_master: dict[str, str] = {}
    for m in re.finditer(r"<Relationship\s([^>]+?)/>", pres_rels, re.DOTALL):
        attrs = m.group(1)
        id_m = re.search(r'Id="([^"]+)"', attrs)
        target_m = re.search(r'Target="([^"]+)"', attrs)
        type_m = re.search(r'Type="([^"]+)"', attrs)
        if id_m and target_m and type_m and "slideMaster" in (type_m.group(1)):
            rid = id_m.group(1)
            target = f"ppt/{target_m.group(1)}"
            rid_to_master[rid] = target

    # Remove sldMasterId entries for removed masters
    for rid, master_path in rid_to_master.items():
        if master_path not in used_masters:
            # Match the exact rId with trailing quote to avoid prefix collisions
            # e.g. rId10 must not match rId107
            text = re.sub(
                r'<p:sldMasterId\s[^>]*?"' + re.escape(rid) + r'"[^>]*?/>',
                "",
                text,
            )

    text = re.sub(r"\n\s*\n", "\n", text)
    return text.encode("utf-8")


def _patch_rels_remove_targets(
    data: bytes,
    removed: set[str],
    base_dir: str,
) -> bytes:
    """Remove Relationship entries whose resolved Target was removed."""
    text = data.decode("utf-8")

    def _filter_rel(m: re.Match) -> str:
        attrs = m.group(0)
        target_m = re.search(r'Target="([^"]+)"', attrs)
        if not target_m:
            return attrs

        target = target_m.group(1)
        # Skip external
        if target.startswith("http") or 'TargetMode="External"' in attrs:
            return attrs

        # Resolve relative to base_dir
        if target.startswith("../"):
            parts = base_dir.rstrip("/").split("/")
            for seg in target.split("/"):
                if seg == "..":
                    if parts:
                        parts.pop()
                elif seg and seg != ".":
                    parts.append(seg)
            resolved = "/".join(parts)
        else:
            resolved = base_dir.rstrip("/") + "/" + target if base_dir else target

        if resolved in removed:
            return ""
        return attrs

    text = re.sub(r"<Relationship\s[^>]+?/>", _filter_rel, text)
    text = re.sub(r"\n\s*\n", "\n", text)
    return text.encode("utf-8")


def _patch_master_layout_ids(
    data: bytes,
    zf: zipfile.ZipFile,
    master_path: str,
    all_names: set[str],
    removed: set[str],
) -> bytes:
    """Remove sldLayoutId entries for layouts that were removed."""
    text = data.decode("utf-8")

    # Read the master's rels to map rId → layout path
    rels_path = _rels_path_for(master_path)
    if rels_path not in all_names:
        return data

    master_rels = _read(zf, rels_path)
    rid_to_layout: dict[str, str] = {}
    for m in re.finditer(r"<Relationship\s([^>]+?)/>", master_rels, re.DOTALL):
        attrs = m.group(1)
        id_m = re.search(r'Id="([^"]+)"', attrs)
        target_m = re.search(r'Target="([^"]+)"', attrs)
        if id_m and target_m:
            target = target_m.group(1)
            resolved = _resolve(master_path, target)
            rid_to_layout[id_m.group(1)] = resolved

    # Remove sldLayoutId entries for removed layouts
    for rid, layout_path in rid_to_layout.items():
        if layout_path in removed:
            # Use trailing quote to avoid prefix collisions (rId10 vs rId107)
            text = re.sub(
                r'<p:sldLayoutId\s[^>]*?"' + re.escape(rid) + r'"[^>]*?/>',
                "",
                text,
            )

    text = re.sub(r"\n\s*\n", "\n", text)
    return text.encode("utf-8")


def _inject_docprops_rels(data: bytes) -> bytes:
    """Rewrite _rels/.rels to include docProps relationships."""
    return (
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
        '</Relationships>'
    ).encode("utf-8")


def _inject_docprops_content_types(data: bytes) -> bytes:
    """Add docProps Override entries to [Content_Types].xml."""
    text = data.decode("utf-8")
    if "docProps/core.xml" not in text:
        overrides = (
            '<Override PartName="/docProps/core.xml"'
            ' ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/docProps/app.xml"'
            ' ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        )
        text = text.replace("</Types>", f"{overrides}</Types>")
    return text.encode("utf-8")


def _write_docprops(zf_out: zipfile.ZipFile) -> None:
    """Write docProps/core.xml and docProps/app.xml into the ZIP."""
    core_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<cp:coreProperties'
        ' xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
        ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
        ' xmlns:dcterms="http://purl.org/dc/terms/"'
        ' xmlns:dcmitype="http://purl.org/dc/dcmitype/"'
        ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>Presentation</dc:title>'
        '<dcterms:created xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:modified>'
        '</cp:coreProperties>'
    )
    zf_out.writestr("docProps/core.xml", core_xml)

    app_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Properties'
        ' xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
        ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>Warhol</Application>'
        '<Slides>0</Slides>'
        '</Properties>'
    )
    zf_out.writestr("docProps/app.xml", app_xml)


def _remap_media_in_rels(data: bytes, filename_remap: dict[str, str]) -> bytes:
    """Remap media filenames in .rels Target attributes for deduplication."""
    text = data.decode("utf-8")
    changed = False

    for old_name, new_name in filename_remap.items():
        old_ref = f"media/{old_name}"
        new_ref = f"media/{new_name}"
        if old_ref in text:
            text = text.replace(old_ref, new_ref)
            changed = True

    return text.encode("utf-8") if changed else data


# ---------------------------------------------------------------------------
# Image compression
# ---------------------------------------------------------------------------

# Maximum image dimension (width or height) for slide backgrounds
_MAX_IMAGE_DIM = 1920

# Target quality for JPEG compression (used for large photographic PNGs)
_JPEG_QUALITY = 85


def _compress_media(
    zf: zipfile.ZipFile,
    kept_names: set[str],
) -> tuple[dict[str, tuple[bytes, str]], dict[str, str], float]:
    """Compress media files: convert animated GIFs to static PNGs, downscale large images.

    Returns:
        media_replacements: {original_path: (new_data_bytes, new_extension)}
        filename_changes: {old_filename: new_filename} for rels remapping
        saved_mb: total MB saved
    """
    from io import BytesIO
    from PIL import Image

    media_replacements: dict[str, tuple[bytes, str]] = {}
    filename_changes: dict[str, str] = {}
    total_saved = 0

    media_files = sorted(
        n for n in kept_names if n.startswith("ppt/media/")
    )

    for media_path in media_files:
        info = zf.getinfo(media_path)
        original_size = info.file_size
        ext = media_path.rsplit(".", 1)[-1].lower()
        filename = media_path.split("/")[-1]

        blob = zf.read(media_path)

        # --- Animated GIF → static PNG (first frame) ---
        if ext == "gif" and original_size > 100_000:
            try:
                img = Image.open(BytesIO(blob))
                # Seek to first frame (already there by default)
                img = img.convert("RGBA")

                # Downscale if oversized
                w, h = img.size
                if max(w, h) > _MAX_IMAGE_DIM:
                    ratio = _MAX_IMAGE_DIM / max(w, h)
                    new_w = int(w * ratio)
                    new_h = int(h * ratio)
                    img = img.resize((new_w, new_h), Image.LANCZOS)

                buf = BytesIO()
                img.save(buf, format="PNG", optimize=True)
                new_data = buf.getvalue()

                if len(new_data) < original_size:
                    new_filename = filename.rsplit(".", 1)[0] + ".png"
                    media_replacements[media_path] = (new_data, "png")
                    filename_changes[filename] = new_filename
                    total_saved += original_size - len(new_data)
                    logger.debug(
                        f"  {filename}: GIF {original_size/(1024*1024):.1f} MB "
                        f"→ PNG {len(new_data)/(1024*1024):.1f} MB"
                    )
            except Exception as e:
                logger.debug(f"  Could not convert {filename}: {e}")
            continue

        # --- Oversized PNG → downscale ---
        if ext == "png" and original_size > 500_000:
            try:
                img = Image.open(BytesIO(blob))
                w, h = img.size

                if max(w, h) > _MAX_IMAGE_DIM:
                    ratio = _MAX_IMAGE_DIM / max(w, h)
                    new_w = int(w * ratio)
                    new_h = int(h * ratio)
                    img = img.resize((new_w, new_h), Image.LANCZOS)

                    buf = BytesIO()
                    img.save(buf, format="PNG", optimize=True)
                    new_data = buf.getvalue()

                    if len(new_data) < original_size * 0.8:  # Only if meaningful savings
                        media_replacements[media_path] = (new_data, "png")
                        total_saved += original_size - len(new_data)
                        logger.debug(
                            f"  {filename}: PNG {w}x{h} → {new_w}x{new_h}, "
                            f"{original_size/(1024*1024):.1f} MB → "
                            f"{len(new_data)/(1024*1024):.1f} MB"
                        )
            except Exception as e:
                logger.debug(f"  Could not resize {filename}: {e}")

    return media_replacements, filename_changes, total_saved / (1024 * 1024)


def _patch_content_types_extensions(
    data: bytes,
    filename_changes: dict[str, str],
) -> bytes:
    """Update [Content_Types].xml Override entries for files that changed extension.

    When a .gif is converted to .png, we need to update both the PartName and
    the ContentType in the Override entry.
    """
    text = data.decode("utf-8")

    for old_name, new_name in filename_changes.items():
        old_ext = old_name.rsplit(".", 1)[-1].lower()
        new_ext = new_name.rsplit(".", 1)[-1].lower()

        if old_ext == new_ext:
            continue

        # Replace PartName
        text = text.replace(
            f"/{old_name.replace(chr(47), chr(47))}",
            f"/{new_name}",
        )
        # More robust: replace in the context of the full media path
        text = text.replace(f"media/{old_name}", f"media/{new_name}")

        # Update content type for this specific entry
        content_type_map = {
            "gif": "image/gif",
            "png": "image/png",
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
        }
        old_ct = content_type_map.get(old_ext, f"image/{old_ext}")
        new_ct = content_type_map.get(new_ext, f"image/{new_ext}")

        if old_ct != new_ct:
            # Only replace the content type for the specific entry we just renamed
            # Find the Override that now references the new filename and fix its type
            pattern = (
                r'(<Override\s+PartName="[^"]*'
                + re.escape(new_name)
                + r'"[^>]*ContentType=")' + re.escape(old_ct) + r'"'
            )
            replacement = r"\g<1>" + new_ct + '"'
            text = re.sub(pattern, replacement, text)

            # Also try reversed attribute order
            pattern2 = (
                r'(<Override\s+ContentType=")' + re.escape(old_ct)
                + r'"(\s+PartName="[^"]*' + re.escape(new_name) + r'")'
            )
            replacement2 = r"\g<1>" + new_ct + r'"\g<2>'
            text = re.sub(pattern2, replacement2, text)

    return text.encode("utf-8")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Repair and compact a PPTX file for cross-platform compatibility"
    )
    parser.add_argument("input_file", type=Path, help="Input PPTX file")
    parser.add_argument(
        "-o", "--output", type=Path, default=None,
        help="Output PPTX path (default: overwrite input)"
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
    print(f"  Original size:        {stats['original_size_mb']:.1f} MB")
    print(f"  Final size:           {stats['final_size_mb']:.1f} MB  ({savings_pct:.0f}% smaller)")
    print(f"  Layouts removed:      {stats['layouts_removed']}")
    print(f"  Masters removed:      {stats['masters_removed']}")
    print(f"  Themes removed:       {stats['themes_removed']}")
    print(f"  Media deduplicated:   {stats['media_deduplicated']}")
    print(f"  Media compressed:     {stats['media_compressed']}  ({stats['media_saved_mb']:.1f} MB saved)")
    print(f"  Media removed:        {stats['media_removed']}")
    print(f"  ContentType entries:  {stats['content_types_cleaned']} cleaned")

    if stats["final_size_mb"] > 100:
        print(f"")
        print(f"  Note: File is still > 100 MB due to template background images.")
        print(f"  Consider a template with smaller background images for Google Slides.")

    print(f"")
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    main()
