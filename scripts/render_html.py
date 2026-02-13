#!/usr/bin/env python3
"""Render a branded HTML slide deck from a deck schema, template matches, and design system.

The HTML is the **primary creative output** — a pixel-level branded presentation
that the PPTX builder (`build_from_html.py`) faithfully reproduces.  The design
system is injected as CSS custom properties so every color, font, and size from
the brand is available as `var(--token-name)`.

Usage:
    # From deck schema (auto-maps content to zones):
    python scripts/render_html.py workspace/deck_schema.json \
        -o workspace/deck_preview.html \
        --design-system design_systems/shopify_technical_workshop.yaml \
        --matches workspace/template_matches.json \
        --template-registry template_registry.json

    # From pre-built HtmlDeck JSON (agent already composed the layout):
    python scripts/render_html.py workspace/html_deck.json \
        -o workspace/deck_preview.html \
        --design-system design_systems/shopify_technical_workshop.yaml \
        --from-html-deck
"""

import argparse
import json
import sys
from pathlib import Path
from html import escape

# Add project root to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.schemas.design_system import DesignSystem
from src.schemas.slide_schema import DeckSchema, SlideSpec
from src.schemas.template_schema import TemplateRegistry
from src.schemas.html_schema import (
    DPI,
    SLIDE_HEIGHT_PX,
    SLIDE_WIDTH_PX,
    ElementPosition,
    FontSpec,
    HtmlDeck,
    HtmlSlide,
    SlideBackground,
    TextElement,
)


# ---------------------------------------------------------------------------
# Design-system → CSS custom properties
# ---------------------------------------------------------------------------

def generate_design_tokens_css(design: DesignSystem) -> str:
    """Generate CSS custom properties from a DesignSystem.

    These tokens let the agent (and the scaffold) use `var(--color-primary)`
    instead of hardcoding hex values.  Every color, font, size, spacing,
    and per-slide-type override is emitted.
    """
    lines = [":root {"]

    # --- Brand colors ---
    lines.append("    /* Brand colors */")
    lines.append(f"    --color-primary: {design.colors.primary};")
    lines.append(f"    --color-secondary: {design.colors.secondary};")
    lines.append(f"    --color-accent: {design.colors.accent};")
    lines.append(f"    --color-text-dark: {design.colors.text_dark};")
    lines.append(f"    --color-text-light: {design.colors.text_light};")
    lines.append(f"    --color-background: {design.colors.background};")
    lines.append(f"    --color-surface: {design.surface_resolved};")
    lines.append(f"    --color-surface-accent: {design.surface_accent_resolved};")
    lines.append(f"    --color-brand-green: {design.brand_green_resolved};")
    lines.append(f"    --color-text-secondary: {design.text_secondary_resolved};")
    lines.append(f"    --color-text-heading: {design.text_heading_resolved};")

    # --- Slide-type backgrounds ---
    lines.append("")
    lines.append("    /* Slide-type backgrounds */")
    lines.append(f"    --bg-title: {design.title_bg};")
    lines.append(f"    --bg-section-header: {design.section_header_bg};")
    lines.append(f"    --bg-content: {design.content_bg};")
    lines.append(f"    --bg-closing: {design.closing_bg};")
    lines.append(f"    --bg-quote: {design.quote_bg};")
    lines.append(f"    --bg-bullet-list: {design.bullet_list_bg};")
    lines.append(f"    --bg-data-point: {design.content_bg};")

    # --- Decoration colors ---
    lines.append("")
    lines.append("    /* Decoration */")
    lines.append(f"    --color-divider: {design.divider_line_color_resolved};")
    lines.append(f"    --color-accent-bar: {design.accent_bar_color_resolved};")
    lines.append(f"    --color-badge-fill: {design.badge_fill_resolved};")
    lines.append(f"    --color-badge-text: {design.badge_text_resolved};")
    lines.append(f"    --color-data-point-accent: {design.data_point_accent};")

    # --- Fonts ---
    lines.append("")
    lines.append("    /* Fonts */")
    lines.append(f"    --font-title: '{design.fonts.title_font}';")
    lines.append(f"    --font-body: '{design.fonts.body_font}';")
    lines.append(f"    --font-emphasis: '{design.emphasis_font_resolved}';")
    lines.append(f"    --font-light: '{design.light_font_resolved}';")
    lines.append(f"    --font-medium: '{design.medium_font_resolved}';")
    lines.append(f"    --font-extra-light: '{design.extra_light_font_resolved}';")
    lines.append(f"    --font-label: '{design.label_font_resolved}';")
    lines.append(f"    --font-quote: '{design.quote_font_resolved}';")

    # --- Sizes ---
    lines.append("")
    lines.append("    /* Font sizes */")
    lines.append(f"    --size-title: {design.fonts.title_size}pt;")
    lines.append(f"    --size-subtitle: {design.fonts.subtitle_size}pt;")
    lines.append(f"    --size-body: {design.fonts.body_size}pt;")
    lines.append(f"    --size-bullet: {design.fonts.bullet_size}pt;")
    lines.append(f"    --size-data-point: {design.data_point_size_resolved}pt;")
    lines.append(f"    --size-quote: {design.quote_size_resolved}pt;")
    lines.append(f"    --size-caption: {design.caption_size_resolved}pt;")
    lines.append(f"    --size-label: {design.label_size_resolved}pt;")

    # --- Spacing ---
    lines.append("")
    lines.append("    /* Spacing and margins (px at 96 DPI) */")
    lines.append(f"    --margin-left: {design.content_area.margin_left * DPI:.1f}px;")
    lines.append(f"    --margin-top: {design.content_area.margin_top * DPI:.1f}px;")
    lines.append(f"    --margin-right: {design.content_area.margin_right * DPI:.1f}px;")
    lines.append(f"    --margin-bottom: {design.content_area.margin_bottom * DPI:.1f}px;")

    # --- Slide dimensions ---
    lines.append("")
    lines.append("    /* Slide dimensions */")
    lines.append(f"    --slide-width: {int(design.dimensions.width * DPI)}px;")
    lines.append(f"    --slide-height: {int(design.dimensions.height * DPI)}px;")

    lines.append("}")
    return "\n".join(lines)


def generate_auto_styling_css() -> str:
    """Generate CSS rules that auto-apply branding per slide type and visual profile.

    These rules mean the agent can set `data-slide-type` and `data-visual-profile`
    and get correct backgrounds and text colors automatically.
    """
    return """\
/* --- Auto-apply branded backgrounds per slide type --- */
.slide[data-slide-type="title"] > .slide-bg-auto { background: var(--bg-title); }
.slide[data-slide-type="section_header"] > .slide-bg-auto { background: var(--bg-section-header); }
.slide[data-slide-type="content"] > .slide-bg-auto { background: var(--bg-content); }
.slide[data-slide-type="two_column"] > .slide-bg-auto { background: var(--bg-content); }
.slide[data-slide-type="bullet_list"] > .slide-bg-auto { background: var(--bg-bullet-list); }
.slide[data-slide-type="closing"] > .slide-bg-auto { background: var(--bg-closing); }
.slide[data-slide-type="quote"] > .slide-bg-auto { background: var(--bg-quote); }
.slide[data-slide-type="data_point"] > .slide-bg-auto { background: var(--bg-data-point); }
.slide[data-slide-type="comparison"] > .slide-bg-auto { background: var(--bg-content); }
.slide[data-slide-type="timeline"] > .slide-bg-auto { background: var(--bg-content); }
.slide[data-slide-type="image_full"] > .slide-bg-auto { background: var(--bg-content); }
.slide[data-slide-type="image_with_text"] > .slide-bg-auto { background: var(--bg-content); }
.slide[data-slide-type="chart"] > .slide-bg-auto { background: var(--bg-content); }
.slide[data-slide-type="team"] > .slide-bg-auto { background: var(--bg-content); }

/* --- Auto text colors per visual profile --- */
.slide[data-visual-profile="dark"] { color: var(--color-text-light); }
.slide[data-visual-profile="light"] { color: var(--color-text-dark); }
.slide[data-visual-profile="branded_image"] { color: var(--color-text-light); }
.slide[data-visual-profile="minimal"] { color: var(--color-text-dark); }
"""


# ---------------------------------------------------------------------------
# Base CSS stylesheet (layout chrome, not brand-specific)
# ---------------------------------------------------------------------------

BASE_CSS = """\
/* =========================================================
   Warhol — HTML Slide Deck Stylesheet
   ========================================================= */

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

/* --- Page chrome --- */
html {
    background: #1a1a1a;
    color-scheme: dark;
}
body {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 48px;
    padding: 48px 20px 80px;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    min-height: 100vh;
}

/* --- Deck title bar --- */
.deck-header {
    text-align: center;
    color: #ccc;
    padding: 16px 0 8px;
}
.deck-header h1 {
    font-size: 20px;
    font-weight: 600;
    color: #eee;
    margin-bottom: 4px;
}
.deck-header p {
    font-size: 13px;
    color: #888;
}

/* --- Slide card --- */
.slide-wrapper {
    position: relative;
    box-shadow: 0 8px 32px rgba(0,0,0,0.5), 0 2px 8px rgba(0,0,0,0.3);
    border-radius: 6px;
    overflow: hidden;
    transition: transform 0.15s ease;
}
.slide-wrapper:hover {
    transform: scale(1.005);
}

/* --- Slide frame --- */
.slide {
    position: relative;
    overflow: hidden;
    background: var(--color-background, #fff);
}

/* --- Background layer --- */
.slide-bg {
    position: absolute;
    inset: 0;
    z-index: 0;
}
.slide-bg-auto {
    /* Color set by data-slide-type auto-styling rules above */
}
.slide-bg-solid {
    /* Solid fills applied via inline style */
}
.slide-bg-template {
    /* Template-clone: actual background color applied via inline style.
       The data attributes tell the PPTX builder which template to clone. */
}

/* --- Text elements --- */
.element {
    position: absolute;
    overflow: hidden;
    word-wrap: break-word;
    overflow-wrap: break-word;
    z-index: 1;
}
.element:hover {
    outline: 1px dashed rgba(100, 180, 255, 0.35);
    outline-offset: 2px;
}

/* Bullet lists */
.element ul, .element ol {
    margin: 0;
    padding-left: 1.3em;
    list-style-type: disc;
}
.element ol {
    list-style-type: decimal;
}
.element li {
    margin-bottom: 0.3em;
    line-height: 1.35;
}
.element li:last-child {
    margin-bottom: 0;
}

/* --- Speaker notes (hidden in preview, present in DOM for builder) --- */
.speaker-notes {
    display: none;
}

/* --- Slide number badge (lower-right) --- */
.slide-number-badge {
    position: absolute;
    bottom: 8px;
    right: 14px;
    font-size: 10px;
    color: rgba(255,255,255,0.3);
    font-family: 'SF Mono', 'Fira Code', monospace;
    pointer-events: none;
    z-index: 2;
}

/* --- Slide label below card --- */
.slide-label {
    text-align: center;
    color: #666;
    font-size: 11px;
    margin-top: 6px;
    font-family: 'SF Mono', 'Fira Code', monospace;
    letter-spacing: 0.02em;
    max-width: 100%;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}

/* --- Visual profile indicators (border accent) --- */
.slide[data-visual-profile="dark"] {
    border-left: 3px solid var(--color-primary, #6366f1);
}
.slide[data-visual-profile="light"] {
    border-left: 3px solid var(--color-accent, #10b981);
}
.slide[data-visual-profile="branded_image"] {
    border-left: 3px solid var(--color-secondary, #f59e0b);
}
.slide[data-visual-profile="minimal"] {
    border-left: 3px solid #94a3b8;
}

/* --- Responsive --- */
@media (max-width: 1000px) {
    .slide-wrapper {
        transform: scale(0.85);
        transform-origin: top center;
    }
}
@media (max-width: 700px) {
    .slide-wrapper {
        transform: scale(0.6);
        transform-origin: top center;
    }
    body { gap: 24px; padding: 24px 10px; }
}

/* --- Print --- */
@media print {
    html { background: white; }
    body { gap: 0; padding: 0; }
    .slide-wrapper {
        box-shadow: none;
        page-break-after: always;
        border-radius: 0;
    }
    .slide-label { display: none; }
    .deck-header { display: none; }
    .slide-number-badge { display: none; }
    .element:hover { outline: none; }
}
"""


# ---------------------------------------------------------------------------
# Deck-schema → HtmlDeck mapping
# ---------------------------------------------------------------------------

def _zone_position_to_px(position: tuple[float, float, float, float]) -> ElementPosition:
    """Convert a content zone position (inches) to ElementPosition (px at 96 DPI)."""
    left, top, width, height = position
    return ElementPosition(
        left=round(left * DPI, 1),
        top=round(top * DPI, 1),
        width=round(width * DPI, 1),
        height=round(height * DPI, 1),
    )


def _font_for_role(
    role: str, design: DesignSystem, visual_profile: str = "light"
) -> FontSpec:
    """Build a FontSpec for a given content role using the design system."""
    is_dark = visual_profile in ("dark", "branded_image")
    text_color = design.dark_slide_text if is_dark else design.text_heading_resolved

    if role == "title":
        return FontSpec(
            family=design.emphasis_font_resolved,
            size_pt=design.fonts.title_size,
            color=text_color,
            bold=True,
            alignment=design.paragraph.title_alignment,
            line_spacing=design.paragraph.title_line_spacing,
        )
    elif role == "subtitle":
        return FontSpec(
            family=design.light_font_resolved,
            size_pt=design.fonts.subtitle_size,
            color=text_color,
            alignment=design.paragraph.subtitle_alignment,
            line_spacing=design.paragraph.body_line_spacing,
        )
    elif role in ("body", "bullets", "bullet_area"):
        return FontSpec(
            family=design.fonts.body_font,
            size_pt=design.fonts.body_size,
            color=design.text_secondary_resolved if not is_dark else text_color,
            alignment=design.paragraph.body_alignment,
            line_spacing=design.paragraph.body_line_spacing,
        )
    elif role == "data_point":
        return FontSpec(
            family=design.extra_light_font_resolved,
            size_pt=design.data_point_size_resolved,
            color=design.data_point_accent,
            bold=True,
            alignment="center",
        )
    elif role == "quote":
        return FontSpec(
            family=design.quote_font_resolved,
            size_pt=design.quote_size_resolved,
            color=text_color,
            italic=True,
            alignment="left",
        )
    elif role in ("caption", "label", "section_marker"):
        return FontSpec(
            family=design.label_font_resolved,
            size_pt=design.label_size_resolved,
            color=design.text_secondary_resolved if not is_dark else text_color,
            alignment="left",
        )
    else:
        return FontSpec(
            family=design.fonts.body_font,
            size_pt=design.fonts.body_size,
            color=text_color,
        )


def _map_role(zone_type: str) -> str:
    """Map a template content zone type to an HTML element role."""
    mapping = {
        "title": "title",
        "subtitle": "subtitle",
        "body": "body",
        "bullet_area": "bullets",
        "data_point": "data_point",
        "caption": "caption",
    }
    return mapping.get(zone_type, "body")


def _content_for_role(role: str, spec: SlideSpec) -> str:
    """Extract the best content string for a given role from the SlideSpec."""
    if role == "title":
        return spec.title or ""
    elif role == "subtitle":
        return spec.subtitle or ""
    elif role == "data_point":
        for block in spec.content_blocks:
            if block.type == "data_point":
                return block.content
        return ""
    elif role == "quote":
        for block in spec.content_blocks:
            if block.type == "quote":
                return block.content
        return ""
    else:
        parts = []
        for block in spec.content_blocks:
            if block.type in ("body", "bullets", "caption"):
                parts.append(block.content)
        return "\n\n".join(parts)


def _bullet_items_for_spec(spec: SlideSpec) -> list[str] | None:
    """Extract bullet items from a SlideSpec if any bullet blocks exist."""
    for block in spec.content_blocks:
        if block.type == "bullets":
            items = [line.strip() for line in block.content.split("\n") if line.strip()]
            if items:
                return items
    return None


def deck_schema_to_html_deck(
    deck: DeckSchema,
    design: DesignSystem,
    matches: list[dict] | None = None,
    registry: TemplateRegistry | None = None,
) -> HtmlDeck:
    """Convert a DeckSchema to an HtmlDeck using template matches and design system."""
    match_lookup: dict[int, dict] = {}
    if matches:
        for m in matches:
            match_lookup[m["slide_number"]] = m

    html_slides: list[HtmlSlide] = []

    for spec in deck.slides:
        match_info = match_lookup.get(spec.slide_number)
        visual_profile = spec.visual_profile or "light"

        elements: list[TextElement] = []
        template_index: int | None = None
        build_mode = "compose"  # default

        # --- Clone mode: template-matched slide with content zones ---
        if (
            match_info
            and match_info.get("match_type") == "use_as_is"
            and registry
        ):
            template_index = match_info["template_index"]
            template = registry.templates[template_index]
            visual_profile = template.visual_profile or visual_profile

            # Use template's background_color for accurate HTML preview
            bg_color = template.background_color or _bg_color_for_type(
                spec.slide_type.value, design
            )

            bg = SlideBackground(
                bg_type="template_clone",
                template_file=template.template_file,
                slide_index=template.slide_index,
                color=bg_color,
            )
            build_mode = "clone"

            # In clone mode, ONLY emit elements that map to content zones.
            # Each element MUST have a shape_name so the builder can target
            # the exact shape in the cloned slide. Content that doesn't map
            # to a zone is pushed to speaker notes.
            unmapped_content = []
            for zone in template.content_zones:
                role = _map_role(zone.zone_type)
                content = _content_for_role(role, spec)
                if not content:
                    continue

                pos = _zone_position_to_px(zone.position)

                zone_min, zone_max = zone.font_size_range
                font = _font_for_role(role, design, visual_profile)
                font.size_pt = min(font.size_pt, zone_max)
                font.size_pt = max(font.size_pt, zone_min)

                bullet_items = None
                if role == "bullets":
                    bullet_items = _bullet_items_for_spec(spec)

                elements.append(TextElement(
                    role=role,
                    content=content[:zone.max_chars] if zone.max_chars else content,
                    position=pos,
                    font=font,
                    shape_name=zone.shape_name,  # REQUIRED for clone mode
                    bullet_items=bullet_items,
                ))

        # --- Compose mode: build from scratch with branded layout ---
        else:
            bg_color = _bg_color_for_type(spec.slide_type.value, design)
            bg = SlideBackground(bg_type="layout", color=bg_color)
            build_mode = "compose"
            elements = _compose_elements(spec, design, visual_profile)

        html_slides.append(HtmlSlide(
            slide_number=spec.slide_number,
            slide_type=spec.slide_type.value,
            build_mode=build_mode,
            background=bg,
            elements=elements,
            visual_profile=visual_profile,
            speaker_notes=spec.speaker_notes,
            template_index=template_index,
            intent=spec.intent,
        ))

    return HtmlDeck(
        title=deck.title,
        subtitle=deck.subtitle,
        slides=html_slides,
    )


def _bg_color_for_type(slide_type: str, design: DesignSystem) -> str:
    """Get the background color for a slide type from the design system."""
    type_map = {
        "title": design.title_bg,
        "section_header": design.section_header_bg,
        "closing": design.closing_bg,
        "quote": design.quote_bg,
        "data_point": design.content_bg,
        "content": design.content_bg,
        "two_column": design.content_bg,
        "bullet_list": design.bullet_list_bg,
        "comparison": design.content_bg,
        "timeline": design.content_bg,
        "image_full": design.content_bg,
        "image_with_text": design.content_bg,
        "chart": design.content_bg,
        "team": design.content_bg,
    }
    return type_map.get(slide_type, design.content_bg)


def _compose_elements(
    spec: SlideSpec,
    design: DesignSystem,
    visual_profile: str,
) -> list[TextElement]:
    """Create elements for a compose-mode slide using design system defaults."""
    elements: list[TextElement] = []
    ml = design.content_area.margin_left * DPI
    mt = design.content_area.margin_top * DPI
    content_w = (design.dimensions.width - design.content_area.margin_left - design.content_area.margin_right) * DPI
    slide_h = design.dimensions.height * DPI

    if spec.title:
        title_h = 62
        elements.append(TextElement(
            role="title",
            content=spec.title,
            position=ElementPosition(left=ml, top=mt, width=content_w, height=title_h),
            font=_font_for_role("title", design, visual_profile),
        ))
        body_top = mt + title_h + 24
    else:
        body_top = mt

    if spec.subtitle:
        sub_h = 40
        elements.append(TextElement(
            role="subtitle",
            content=spec.subtitle,
            position=ElementPosition(left=ml, top=body_top, width=content_w, height=sub_h),
            font=_font_for_role("subtitle", design, visual_profile),
        ))
        body_top += sub_h + 16

    body_content = _content_for_role("body", spec)
    bullet_items = _bullet_items_for_spec(spec)
    body_h = slide_h - body_top - (design.content_area.margin_bottom * DPI)

    if body_content or bullet_items:
        role = "bullets" if bullet_items else "body"
        elements.append(TextElement(
            role=role,
            content=body_content,
            position=ElementPosition(left=ml, top=body_top, width=content_w, height=max(body_h, 40)),
            font=_font_for_role(role, design, visual_profile),
            bullet_items=bullet_items,
        ))

    data_text = _content_for_role("data_point", spec)
    if data_text:
        elements.append(TextElement(
            role="data_point",
            content=data_text,
            position=ElementPosition(left=ml, top=mt + 80, width=content_w, height=120),
            font=_font_for_role("data_point", design, visual_profile),
        ))

    return elements


# ---------------------------------------------------------------------------
# HTML rendering
# ---------------------------------------------------------------------------

def render_html(deck: HtmlDeck, design: DesignSystem | None = None) -> str:
    """Render an HtmlDeck to a complete HTML string.

    If a DesignSystem is provided, its tokens are injected as CSS custom
    properties so the HTML is fully branded.
    """
    slide_w = deck.slide_width_px
    slide_h = deck.slide_height_px

    # Build the CSS: base layout + design tokens + auto-styling
    css_parts = [BASE_CSS]
    if design:
        css_parts.append(generate_design_tokens_css(design))
        css_parts.append(generate_auto_styling_css())
    css_parts.append(f".slide-wrapper {{ width: {slide_w}px; }}")
    css_parts.append(f".slide {{ width: {slide_w}px; height: {slide_h}px; }}")
    full_css = "\n\n".join(css_parts)

    parts: list[str] = []
    subtitle_html = f"<p>{escape(deck.subtitle)}</p>" if deck.subtitle else ""
    slide_count = len(deck.slides)
    parts.append(f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{escape(deck.title)}</title>
<style>
{full_css}
</style>
</head>
<body>
<div class="deck-header">
<h1>{escape(deck.title)}</h1>
{subtitle_html}
<p>{slide_count} slides &middot; {slide_w}&times;{slide_h}px &middot; Warhol</p>
</div>
""")

    for slide in deck.slides:
        parts.append(_render_slide(slide, slide_w, slide_h))

    parts.append("</body>\n</html>\n")
    return "".join(parts)


def _render_slide(slide: HtmlSlide, w: int, h: int) -> str:
    """Render a single HtmlSlide to HTML."""
    attrs = [
        f'data-slide-number="{slide.slide_number}"',
        f'data-slide-type="{escape(slide.slide_type)}"',
        f'data-visual-profile="{escape(slide.visual_profile)}"',
        f'data-build-mode="{slide.build_mode}"',
    ]
    if slide.template_index is not None:
        attrs.append(f'data-template-index="{slide.template_index}"')

    attrs_str = " ".join(attrs)

    parts: list[str] = []
    parts.append(f'<div class="slide-wrapper">')
    parts.append(f'<div class="slide" {attrs_str}>')

    # Background
    bg = slide.background
    if bg.bg_type == "template_clone" and bg.template_file is not None:
        # Clone mode: data attributes tell the PPTX builder which slide to clone.
        bg_style = ""
        if bg.color:
            bg_style = f' style="background:{bg.color};"'
        parts.append(
            f'<div class="slide-bg slide-bg-template"'
            f' data-bg-type="template_clone"'
            f' data-template-file="{escape(bg.template_file)}"'
            f' data-slide-index="{bg.slide_index}"'
            f'{bg_style}'
            f'></div>'
        )
    elif bg.bg_type == "layout" and bg.color:
        # Compose mode: use a branded layout from the base template.
        # The builder creates a slide from a layout (master bg preserved,
        # no content shapes cloned).
        parts.append(
            f'<div class="slide-bg slide-bg-solid"'
            f' data-bg-type="layout"'
            f' style="background:{bg.color};"'
            f'></div>'
        )
    elif bg.bg_type == "solid" and bg.color:
        parts.append(
            f'<div class="slide-bg slide-bg-solid"'
            f' data-bg-type="solid"'
            f' style="background:{bg.color};"'
            f'></div>'
        )
    else:
        # Auto background from slide type
        parts.append(
            f'<div class="slide-bg slide-bg-auto"'
            f' data-bg-type="layout"'
            f'></div>'
        )

    # Elements
    for elem in slide.elements:
        parts.append(_render_element(elem))

    # Slide number badge
    parts.append(f'<div class="slide-number-badge">{slide.slide_number}</div>')

    # Speaker notes (hidden)
    if slide.speaker_notes:
        parts.append(f'<div class="speaker-notes">{escape(slide.speaker_notes)}</div>')

    parts.append('</div>')  # .slide

    # Label below slide
    label = f"Slide {slide.slide_number} — {slide.slide_type}"
    if slide.intent:
        label += f" — {slide.intent[:60]}"
    parts.append(f'<div class="slide-label">{escape(label)}</div>')

    parts.append('</div>\n')  # .slide-wrapper

    return "\n".join(parts)


def _render_element(elem: TextElement) -> str:
    """Render a single TextElement to an HTML div."""
    pos = elem.position
    font = elem.font

    style_parts = [
        f"left:{pos.left:.1f}px",
        f"top:{pos.top:.1f}px",
        f"width:{pos.width:.1f}px",
        f"height:{pos.height:.1f}px",
        f"font-family:'{escape(font.family)}'",
        f"font-size:{font.size_pt}pt",
        f"color:{font.color}",
    ]
    if font.bold:
        style_parts.append("font-weight:bold")
    if font.italic:
        style_parts.append("font-style:italic")
    if font.alignment:
        style_parts.append(f"text-align:{font.alignment}")
    if font.line_spacing:
        style_parts.append(f"line-height:{font.line_spacing}")

    style = "; ".join(style_parts)

    data_attrs = [f'data-role="{elem.role}"']
    if elem.shape_name:
        data_attrs.append(f'data-shape-name="{escape(elem.shape_name)}"')
    data_str = " ".join(data_attrs)

    if elem.bullet_items:
        items_html = "\n".join(f"  <li>{escape(item)}</li>" for item in elem.bullet_items)
        inner = f"<ul>\n{items_html}\n</ul>"
    else:
        inner = escape(elem.content).replace("\n\n", "<br><br>").replace("\n", "<br>")

    return (
        f'<div class="element" {data_str} style="{style}">'
        f'{inner}'
        f'</div>'
    )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Render branded HTML slide deck")
    parser.add_argument("input_json", type=Path, help="Deck schema JSON or HtmlDeck JSON")
    parser.add_argument("-o", "--output", type=Path, default=Path("workspace/deck_preview.html"),
                        help="Output HTML path (default: workspace/deck_preview.html)")
    parser.add_argument("--design-system", type=Path, default=Path("design_systems/default.yaml"),
                        help="Design system YAML path")
    parser.add_argument("--matches", type=Path, default=None,
                        help="Optional template_matches.json path")
    parser.add_argument("--template-registry", type=Path, default=None,
                        help="Optional template_registry.json path")
    parser.add_argument("--from-html-deck", action="store_true",
                        help="Input is already an HtmlDeck JSON (skip conversion)")
    args = parser.parse_args()

    if not args.input_json.exists():
        print(f"Error: Input not found: {args.input_json}", file=sys.stderr)
        sys.exit(1)

    raw = args.input_json.read_text(encoding="utf-8")
    data = json.loads(raw)

    # Always load the design system for CSS token injection
    design = None
    if args.design_system.exists():
        design = DesignSystem.from_yaml(args.design_system)

    if args.from_html_deck:
        html_deck = HtmlDeck.model_validate(data)
    else:
        if design is None:
            print(f"Error: Design system not found: {args.design_system}", file=sys.stderr)
            sys.exit(1)

        deck = DeckSchema.model_validate(data)

        matches = None
        if args.matches and args.matches.exists():
            matches_data = json.loads(args.matches.read_text(encoding="utf-8"))
            matches = matches_data.get("matches", [])

        registry = None
        if args.template_registry and args.template_registry.exists():
            registry = TemplateRegistry.load(args.template_registry)

        html_deck = deck_schema_to_html_deck(deck, design, matches, registry)

    # Render with design tokens
    html_str = render_html(html_deck, design=design)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(html_str, encoding="utf-8")

    # Also save the HtmlDeck JSON for validation / re-rendering
    json_path = args.output.with_suffix(".json")
    json_path.write_text(html_deck.model_dump_json(indent=2), encoding="utf-8")

    print(f"HTML deck: {args.output}")
    print(f"HtmlDeck JSON: {json_path}")
    print(f"Slides: {len(html_deck.slides)}")
    total_elements = sum(len(s.elements) for s in html_deck.slides)
    print(f"Total elements: {total_elements}")
    if design:
        print(f"Design system: {design.name}")


if __name__ == "__main__":
    main()
