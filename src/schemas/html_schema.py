"""Pydantic models for the HTML intermediate slide format.

The HTML layer sits between the DeckSchema (content intent) and the final
PPTX build.  The agent generates an HTML file where each slide is a
<section> with absolute-positioned elements.  The HTML serves two purposes:

1. **Visual contract** -- pixel-level specification the PPTX builder
   faithfully reproduces (positions, fonts, colors are explicit).
2. **Preview artifact** -- can be opened in a browser for visual QA
   before the PPTX build runs.

All spatial values use a 96-DPI coordinate system:
    px → inches:  px / 96
    inches → EMU: inches * 914400

Font sizes are in typographic points (pt), matching PPTX Pt directly.
Colors are 6-digit hex RGB strings (e.g., "#FFFFFF").
"""

from typing import Literal, Optional

from pydantic import BaseModel, Field


# ---------------------------------------------------------------------------
# Slide dimensions (constant, matches base template)
# ---------------------------------------------------------------------------

SLIDE_WIDTH_IN = 10.0
SLIDE_HEIGHT_IN = 5.625
DPI = 96
SLIDE_WIDTH_PX = int(SLIDE_WIDTH_IN * DPI)   # 960
SLIDE_HEIGHT_PX = int(SLIDE_HEIGHT_IN * DPI)  # 540


# ---------------------------------------------------------------------------
# Element position / size
# ---------------------------------------------------------------------------

class ElementPosition(BaseModel):
    """Absolute position and size within the slide frame (px at 96 DPI)."""

    left: float = Field(ge=0, description="Left edge in px")
    top: float = Field(ge=0, description="Top edge in px")
    width: float = Field(gt=0, description="Width in px")
    height: float = Field(gt=0, description="Height in px")

    @property
    def left_inches(self) -> float:
        return self.left / DPI

    @property
    def top_inches(self) -> float:
        return self.top / DPI

    @property
    def width_inches(self) -> float:
        return self.width / DPI

    @property
    def height_inches(self) -> float:
        return self.height / DPI


# ---------------------------------------------------------------------------
# Font specification
# ---------------------------------------------------------------------------

class FontSpec(BaseModel):
    """Explicit font specification for a text element."""

    family: str = Field(description="Font family name, e.g. 'Shopify Sans Bold'")
    size_pt: float = Field(gt=0, description="Font size in points")
    color: str = Field(
        default="#000000",
        pattern=r"^#[0-9A-Fa-f]{6}$",
        description="Hex RGB color (e.g., '#FFFFFF')",
    )
    bold: bool = False
    italic: bool = False
    alignment: Literal["left", "center", "right"] = "left"
    line_spacing: Optional[float] = Field(
        default=None,
        description="Line spacing multiplier (e.g., 1.0, 1.15, 1.2). None = inherit.",
    )


# ---------------------------------------------------------------------------
# Slide background
# ---------------------------------------------------------------------------

class SlideBackground(BaseModel):
    """Background specification for a slide.

    Three modes:
    - template_clone: clone a slide from a PPTX template file (preserves all
      branded visuals — backgrounds, images, decorations).
    - layout: use a branded layout from the base template (master/background
      inherited, but no content shapes cloned).
    - solid: simple solid-color fill.
    """

    bg_type: Literal["template_clone", "layout", "solid"] = Field(
        description="'template_clone' to clone from a PPTX, 'layout' for a branded base layout, 'solid' for a solid color fill"
    )

    # --- template_clone fields ---
    template_file: Optional[str] = Field(
        default=None,
        description="Path to the source PPTX template (relative to project root)",
    )
    slide_index: Optional[int] = Field(
        default=None,
        description="0-based slide index in the source template to clone",
    )

    # --- thumbnail for HTML preview ---
    thumbnail_path: Optional[str] = Field(
        default=None,
        description="Path to slide thumbnail PNG for richer HTML preview",
    )

    # --- solid fields ---
    color: Optional[str] = Field(
        default=None,
        pattern=r"^#[0-9A-Fa-f]{6}$",
        description="Solid fill color as hex RGB",
    )


# ---------------------------------------------------------------------------
# Text element
# ---------------------------------------------------------------------------

class TextElement(BaseModel):
    """A single text element on a slide.

    Each element maps to one PPTX text shape.  The builder creates (or
    targets) a shape at the specified position and populates it with the
    given content, font, and alignment.
    """

    role: Literal[
        "title", "subtitle", "body", "bullets", "quote",
        "caption", "data_point", "section_marker", "label",
    ] = Field(description="Semantic role — determines builder behavior")

    content: str = Field(description="Text content (plain text or bullet items separated by \\n)")

    position: ElementPosition = Field(description="Absolute position within the slide")
    font: FontSpec = Field(description="Font specification")

    shape_name: Optional[str] = Field(
        default=None,
        description="Target shape name in the cloned template (for content zone mapping)",
    )

    # Bullet-specific
    bullet_items: Optional[list[str]] = Field(
        default=None,
        description="Individual bullet strings (when role='bullets'). "
                    "If set, overrides content for bullet rendering.",
    )


# ---------------------------------------------------------------------------
# Single slide
# ---------------------------------------------------------------------------

class HtmlSlide(BaseModel):
    """Complete specification for one slide in the HTML deck."""

    slide_number: int = Field(ge=1)
    slide_type: str = Field(description="Slide type from SlideType enum value")

    build_mode: Literal["clone", "compose"] = Field(
        default="compose",
        description="'clone' = clone template slide and replace text in named shapes only; "
                    "'compose' = build from scratch using a branded layout",
    )

    background: SlideBackground = Field(description="How the slide background is produced")
    elements: list[TextElement] = Field(
        default_factory=list,
        description="All text elements on this slide",
    )

    visual_profile: Literal["dark", "light", "branded_image", "minimal"] = "minimal"

    speaker_notes: Optional[str] = Field(
        default=None,
        description="Speaker notes (hidden in HTML, added to PPTX notes)",
    )

    # Metadata carried through for traceability
    template_index: Optional[int] = Field(
        default=None,
        description="Index into template_registry.templates (if template-matched)",
    )
    intent: Optional[str] = Field(
        default=None,
        description="What this slide should accomplish (from deck schema)",
    )


# ---------------------------------------------------------------------------
# Full HTML deck
# ---------------------------------------------------------------------------

class HtmlDeck(BaseModel):
    """Complete HTML deck specification.

    This is the top-level model for the HTML intermediate file.  It can be
    serialised to JSON for validation, and the render_html.py script
    converts it to the actual HTML file.
    """

    title: str
    subtitle: Optional[str] = None
    slide_width_px: int = Field(default=SLIDE_WIDTH_PX)
    slide_height_px: int = Field(default=SLIDE_HEIGHT_PX)
    slides: list[HtmlSlide] = Field(default_factory=list)
