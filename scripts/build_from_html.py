#!/usr/bin/env python3
"""Build a PowerPoint presentation from an HTML slide deck preview.

The HTML file (produced by render_html.py or the agent directly) is the
**visual contract** — every slide is a <section> with absolute-positioned
text elements, explicit fonts/colors, and data attributes that tell this
builder exactly how to construct the PPTX.

Two background modes:
  1. template_clone — clone a slide from a source PPTX (branded visuals)
     and replace text in the cloned shapes, then add any new elements.
  2. solid — create a slide from the base template layout with a solid
     background fill, then add all elements as new text boxes.

Usage:
    python scripts/build_from_html.py workspace/deck_preview.html \
        -o output.pptx \
        --base-template "templates/base/Shopify - Example Technical Workshop Slide Bank.pptx"
"""

import argparse
import json
import logging
import re
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from bs4 import BeautifulSoup, Tag

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt, Emu

from src.pptx_engine.slide_operations import (
    add_blank_slide,
    add_slide_from_layout,
    clear_clone_caches,
    clone_slide_as_is,
    open_base_template,
    create_presentation,
)
from src.pptx_engine.text_operations import (
    add_textbox,
    add_bullet_list,
)
from src.schemas.html_schema import DPI

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


# ---------------------------------------------------------------------------
# HTML parsing helpers
# ---------------------------------------------------------------------------

def _px_to_inches(px: float) -> float:
    """Convert pixels (at 96 DPI) to inches."""
    return px / DPI


def _parse_css_value(style: str, prop: str, default: float = 0.0) -> float:
    """Extract a numeric CSS property value from an inline style string.

    Handles px and pt units. Returns the number stripped of units.
    """
    pattern = rf"{prop}\s*:\s*([\d.]+)\s*(px|pt|em|rem|%)?"
    match = re.search(pattern, style)
    if not match:
        return default
    return float(match.group(1))


def _parse_css_color(style: str, prop: str = "color", default: str = "#000000") -> str:
    """Extract a hex color from a CSS style string."""
    pattern = rf"{prop}\s*:\s*(#[0-9A-Fa-f]{{6}})"
    match = re.search(pattern, style)
    if not match:
        return default
    return match.group(1)


def _parse_css_string(style: str, prop: str, default: str = "") -> str:
    """Extract a string CSS property value (e.g., font-family)."""
    pattern = rf"{prop}\s*:\s*'([^']+)'"
    match = re.search(pattern, style)
    if not match:
        # Try without quotes
        pattern = rf"{prop}\s*:\s*([^;]+)"
        match = re.search(pattern, style)
        if not match:
            return default
        return match.group(1).strip().strip("'\"")
    return match.group(1)


def _parse_css_bool(style: str, prop: str, value: str) -> bool:
    """Check if a CSS property has a specific value."""
    pattern = rf"{prop}\s*:\s*{value}"
    return bool(re.search(pattern, style, re.IGNORECASE))


def _hex_to_rgb(hex_color: str) -> RGBColor:
    """Convert hex color string to RGBColor."""
    hex_color = hex_color.lstrip("#")
    if len(hex_color) > 6:
        hex_color = hex_color[:6]
    return RGBColor(
        int(hex_color[0:2], 16),
        int(hex_color[2:4], 16),
        int(hex_color[4:6], 16),
    )


# ---------------------------------------------------------------------------
# Slide building from HTML sections
# ---------------------------------------------------------------------------

def _build_slide_from_section(
    prs: Presentation,
    section: Tag,
) -> None:
    """Build a single PPTX slide from an HTML <div class="slide"> section."""

    slide_type = section.get("data-slide-type", "content")

    # --- Determine background mode ---
    bg_div = section.find("div", class_="slide-bg")
    bg_type = bg_div.get("data-bg-type", "solid") if bg_div else "solid"

    slide = None

    if bg_type == "template_clone" and bg_div:
        # Clone from template
        template_file = bg_div.get("data-template-file", "")
        slide_index_str = bg_div.get("data-slide-index", "0")
        slide_index = int(slide_index_str) if slide_index_str else 0

        if template_file and Path(template_file).exists():
            try:
                slide = clone_slide_as_is(prs, template_file, slide_index)
                logger.info(
                    f"Slide {section.get('data-slide-number')}: "
                    f"cloned from {template_file} index {slide_index}"
                )
            except Exception as e:
                logger.warning(
                    f"Slide {section.get('data-slide-number')}: "
                    f"clone failed ({e}), falling back to layout"
                )

    if slide is None:
        # Solid background or clone fallback — create from layout
        slide = add_slide_from_layout(prs, slide_type)

        # Apply solid background color if specified
        if bg_div and bg_type == "solid":
            bg_color = bg_div.get("style", "")
            color_match = re.search(r"background\s*:\s*(#[0-9A-Fa-f]{6})", bg_color)
            if color_match:
                _set_slide_bg_color(slide, color_match.group(1))

    # --- Process text elements ---
    cloned_mode = (bg_type == "template_clone" and slide is not None)

    for elem_div in section.find_all("div", class_="element"):
        _add_element_to_slide(slide, elem_div, cloned_mode)

    # --- Speaker notes ---
    notes_div = section.find("div", class_="speaker-notes")
    if notes_div and notes_div.get_text(strip=True):
        _add_speaker_notes(slide, notes_div.get_text(strip=True))


def _add_element_to_slide(slide, elem_div: Tag, cloned_mode: bool) -> None:
    """Add a text element from HTML to a PPTX slide.

    In cloned mode, tries to find and replace text in existing shapes
    by shape_name first. If not found, falls back to creating a new textbox.
    In compose mode, always creates a new textbox.
    """
    style = elem_div.get("style", "")
    role = elem_div.get("data-role", "body")
    shape_name = elem_div.get("data-shape-name", "")

    # Extract position from CSS
    left_px = _parse_css_value(style, "left")
    top_px = _parse_css_value(style, "top")
    width_px = _parse_css_value(style, "width")
    height_px = _parse_css_value(style, "height")

    left = _px_to_inches(left_px)
    top = _px_to_inches(top_px)
    width = _px_to_inches(width_px)
    height = _px_to_inches(height_px)

    # Extract font properties from CSS
    font_family = _parse_css_string(style, "font-family", "Arial")
    font_size_pt = _parse_css_value(style, "font-size", 18)
    font_color = _parse_css_color(style, "color", "#000000")
    is_bold = _parse_css_bool(style, "font-weight", "bold")
    is_italic = _parse_css_bool(style, "font-style", "italic")
    alignment = _parse_css_string(style, "text-align", "left")
    line_spacing_raw = _parse_css_value(style, "line-height", 0)
    line_spacing = line_spacing_raw if line_spacing_raw > 0 else None

    # Extract content
    bullet_items = _extract_bullet_items(elem_div)
    text_content = _extract_text_content(elem_div)

    # --- Try shape replacement in cloned mode ---
    if cloned_mode and shape_name:
        replaced = _replace_shape_text_by_name(
            slide, shape_name, text_content, bullet_items,
            font_family, font_size_pt, font_color, is_bold, is_italic,
        )
        if replaced:
            return

    # --- Create new textbox ---
    if width <= 0 or height <= 0:
        return  # Skip elements with no dimensions

    if bullet_items:
        add_bullet_list(
            slide,
            items=bullet_items,
            left=left,
            top=top,
            width=width,
            height=height,
            font_name=font_family,
            font_size=int(font_size_pt),
            font_color=font_color,
            line_spacing=line_spacing or 1.2,
        )
    else:
        add_textbox(
            slide,
            text_content,
            left=left,
            top=top,
            width=width,
            height=height,
            font_name=font_family,
            font_size=int(font_size_pt),
            font_color=font_color,
            bold=is_bold,
            italic=is_italic,
            alignment=alignment,
            line_spacing=line_spacing,
        )


def _extract_bullet_items(elem_div: Tag) -> list[str]:
    """Extract bullet items from <li> elements within the div."""
    items = []
    for li in elem_div.find_all("li"):
        text = li.get_text(strip=True)
        if text:
            items.append(text)
    return items


def _extract_text_content(elem_div: Tag) -> str:
    """Extract plain text content from an element div.

    Handles <br> as newlines and strips HTML tags.
    """
    # Replace <br> tags with newlines before getting text
    for br in elem_div.find_all("br"):
        br.replace_with("\n")
    return elem_div.get_text(strip=False).strip()


def _replace_shape_text_by_name(
    slide,
    shape_name: str,
    text: str,
    bullet_items: list[str] | None,
    font_family: str,
    font_size_pt: float,
    font_color: str,
    bold: bool,
    italic: bool,
) -> bool:
    """Try to find a shape by name on the slide and replace its text.

    Returns True if the shape was found and text was replaced.
    """
    for shape in slide.shapes:
        if shape.name == shape_name and shape.has_text_frame:
            tf = shape.text_frame
            if not tf.paragraphs:
                return False

            # Preserve first-run formatting, apply our content
            first_para = tf.paragraphs[0]

            # Clear extra paragraphs
            ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
            p_elements = list(tf._element)
            for p_el in p_elements:
                if p_el.tag == f"{{{ns_a}}}p" and p_el is not first_para._element:
                    tf._element.remove(p_el)

            if bullet_items:
                # Multi-paragraph bullet content
                first_para.text = bullet_items[0]
                _apply_font_to_para(first_para, font_family, font_size_pt, font_color, bold, italic)
                for item in bullet_items[1:]:
                    p = tf.add_paragraph()
                    p.text = item
                    _apply_font_to_para(p, font_family, font_size_pt, font_color, bold, italic)
            else:
                first_para.text = text
                _apply_font_to_para(first_para, font_family, font_size_pt, font_color, bold, italic)

            return True
    return False


def _apply_font_to_para(
    para, font_family: str, font_size_pt: float,
    font_color: str, bold: bool, italic: bool
) -> None:
    """Apply font formatting to all runs in a paragraph."""
    for run in para.runs:
        run.font.name = font_family
        run.font.size = Pt(font_size_pt)
        try:
            run.font.color.rgb = _hex_to_rgb(font_color)
        except Exception:
            pass
        run.font.bold = bold
        run.font.italic = italic


def _set_slide_bg_color(slide, hex_color: str) -> None:
    """Set a solid background color on a slide."""
    try:
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = _hex_to_rgb(hex_color)
    except Exception as e:
        logger.debug(f"Could not set background color: {e}")


def _add_speaker_notes(slide, notes_text: str) -> None:
    """Add speaker notes to a slide."""
    try:
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = notes_text
    except Exception as e:
        logger.debug(f"Could not add speaker notes: {e}")


# ---------------------------------------------------------------------------
# Post-processing: clear unmapped shapes on cloned slides
# ---------------------------------------------------------------------------

def _clear_unmapped_shapes(slide, mapped_shape_names: set[str]) -> None:
    """Clear text from shapes that weren't mapped to any HTML element.

    Only applies to cloned slides to remove stale template text.
    """
    for shape in slide.shapes:
        if shape.has_text_frame and shape.name not in mapped_shape_names:
            tf = shape.text_frame
            if tf.paragraphs:
                # Clear all text
                ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
                p_elements = list(tf._element)
                first_p = tf.paragraphs[0]._element
                for p_el in p_elements:
                    if p_el.tag == f"{{{ns_a}}}p" and p_el is not first_p:
                        tf._element.remove(p_el)
                tf.paragraphs[0].text = ""

                # Move off-canvas
                try:
                    shape.left = Emu(914400 * 20)
                except Exception:
                    pass


# ---------------------------------------------------------------------------
# Main build pipeline
# ---------------------------------------------------------------------------

def build_from_html(
    html_path: Path,
    output_path: Path,
    base_template: Path | None = None,
) -> Path:
    """Build a PPTX presentation from an HTML slide deck file.

    Args:
        html_path: Path to the HTML deck preview file.
        output_path: Where to save the output PPTX.
        base_template: Optional base template for layouts/masters.

    Returns:
        Path to the saved PPTX.
    """
    html_content = html_path.read_text(encoding="utf-8")
    soup = BeautifulSoup(html_content, "html.parser")

    # Open base template
    default_base = Path("templates/base/Shopify - Example Technical Workshop Slide Bank.pptx")
    base_path = base_template or default_base

    if base_path.exists():
        prs = open_base_template(base_path)
    else:
        logger.warning(f"Base template not found at {base_path}, creating blank")
        prs = create_presentation()

    # Find all slide sections
    slide_divs = soup.find_all("div", class_="slide")
    if not slide_divs:
        logger.error("No <div class='slide'> sections found in HTML")
        sys.exit(1)

    # Sort by data-slide-number
    def slide_sort_key(div):
        num = div.get("data-slide-number", "0")
        try:
            return int(num)
        except ValueError:
            return 0

    slide_divs.sort(key=slide_sort_key)

    cloned = 0
    composed = 0

    for slide_div in slide_divs:
        bg_div = slide_div.find("div", class_="slide-bg")
        bg_type = bg_div.get("data-bg-type", "solid") if bg_div else "solid"

        _build_slide_from_section(prs, slide_div)

        if bg_type == "template_clone":
            cloned += 1
        else:
            composed += 1

    # Save
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))

    # Free caches
    clear_clone_caches()

    logger.info(
        f"Saved: {output_path} "
        f"({cloned} cloned, {composed} composed, {len(slide_divs)} total)"
    )
    return output_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Build PPTX from HTML slide deck")
    parser.add_argument("html_file", type=Path, help="Path to HTML deck file")
    parser.add_argument("-o", "--output", type=Path, default=Path("output/presentation.pptx"),
                        help="Output PPTX path (default: output/presentation.pptx)")
    parser.add_argument("--base-template", type=Path, default=None,
                        help="Base template PPTX (default: templates/base/...)")
    args = parser.parse_args()

    if not args.html_file.exists():
        print(f"Error: HTML file not found: {args.html_file}", file=sys.stderr)
        sys.exit(1)

    result_path = build_from_html(
        html_path=args.html_file,
        output_path=args.output,
        base_template=args.base_template,
    )

    # Copy HTML alongside the PPTX in the output folder
    import shutil
    html_output = result_path.parent / "presentation.html"
    shutil.copy2(str(args.html_file), str(html_output))

    print(f"PPTX: {result_path}")
    print(f"HTML: {html_output}")


if __name__ == "__main__":
    main()
