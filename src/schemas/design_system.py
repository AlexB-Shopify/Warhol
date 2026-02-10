"""Pydantic models for brand/design system configuration.

The DesignSystem captures everything needed to visually style generated slides:
fonts (with weight variants), colors, paragraph formatting, content area
positioning, decoration parameters, and per-slide-type overrides.
"""

from pathlib import Path
from typing import Optional

import yaml
from pydantic import BaseModel, Field


# ---------------------------------------------------------------------------
# Font configuration
# ---------------------------------------------------------------------------

class FontConfig(BaseModel):
    """Font configuration for the presentation.

    Supports weight-specific font names (e.g., 'Inter Tight SemiBold' for
    emphasis, 'Inter Tight ExtraLight' for subtitles). When weight variants
    are not set, the base title_font / body_font are used everywhere.
    """

    title_font: str = "Arial"
    body_font: str = "Arial"
    title_size: int = Field(default=44, description="Title font size in points")
    subtitle_size: int = Field(default=28, description="Subtitle font size in points")
    body_size: int = Field(default=18, description="Body text font size in points")
    bullet_size: int = Field(default=16, description="Bullet text font size in points")

    # Weight variants — override the base font for specific contexts
    emphasis_font: Optional[str] = Field(
        default=None,
        description="Font for bold/emphasis text (e.g., 'Inter Tight SemiBold'). Falls back to title_font.",
    )
    light_font: Optional[str] = Field(
        default=None,
        description="Light-weight font for subtitles/captions (e.g., 'Inter Tight Light'). Falls back to body_font.",
    )
    medium_font: Optional[str] = Field(
        default=None,
        description="Medium-weight font for labels/column heads (e.g., 'Inter Tight Medium'). Falls back to body_font.",
    )
    extra_light_font: Optional[str] = Field(
        default=None,
        description="Extra-light font for large display text / hero numbers (e.g., 'Inter Tight ExtraLight'). Falls back to light_font.",
    )
    label_font: Optional[str] = Field(
        default=None,
        description="Font for small labels and section markers (e.g., 'Poppins Medium'). Falls back to medium_font.",
    )
    quote_font: Optional[str] = Field(
        default=None,
        description="Font for quote text (e.g., 'Inter Tight Light'). Falls back to body_font italic.",
    )

    # Size variants
    quote_size: Optional[int] = Field(
        default=None,
        description="Font size for quotes in points. Falls back to subtitle_size.",
    )
    data_point_size: Optional[int] = Field(
        default=None,
        description="Font size for hero data point numbers. Falls back to title_size * 1.5.",
    )
    caption_size: Optional[int] = Field(
        default=None,
        description="Font size for captions and attribution text. Falls back to bullet_size - 2.",
    )
    label_size: Optional[int] = Field(
        default=None,
        description="Font size for small labels, section markers, dates (8-10pt). Falls back to 9.",
    )
    section_marker_size: Optional[int] = Field(
        default=None,
        description="Font size for '01 | Topic' markers. Falls back to label_size.",
    )
    hero_number_size: Optional[int] = Field(
        default=None,
        description="Font size for oversized hero numbers (80-100pt). Falls back to data_point_size.",
    )


# ---------------------------------------------------------------------------
# Color configuration
# ---------------------------------------------------------------------------

class ColorConfig(BaseModel):
    """Color palette for the presentation.

    Extended with secondary/surface colors used in the base template for
    shape fills, card backgrounds, and contextual text colors.
    """

    primary: str = "#1a73e8"
    secondary: str = "#34a853"
    accent: str = "#ea4335"
    text_dark: str = "#202124"
    text_light: str = "#5f6368"
    background: str = "#ffffff"

    # Extended palette — matches base template usage
    text_secondary: Optional[str] = Field(
        default=None,
        description="Secondary text color for body copy (e.g., #434343). Falls back to text_dark.",
    )
    text_heading: Optional[str] = Field(
        default=None,
        description="Heading-specific text color (e.g., #191E17). Falls back to text_dark.",
    )
    surface: Optional[str] = Field(
        default=None,
        description="Light surface/card fill color (e.g., #F4F4F4). Falls back to #F4F4F4.",
    )
    surface_accent: Optional[str] = Field(
        default=None,
        description="Tinted surface fill (e.g., #F1FACF). Falls back to surface.",
    )
    brand_green: Optional[str] = Field(
        default=None,
        description="Brand green for badges/highlights (e.g., #0E8155). Falls back to secondary.",
    )


# ---------------------------------------------------------------------------
# Decoration configuration
# ---------------------------------------------------------------------------

class DecorationConfig(BaseModel):
    """Visual decoration parameters for shapes, lines, badges.

    Controls the decorative elements that give slides visual richness:
    divider lines, accent bars, badge shapes, image placeholder areas.
    """

    divider_line_color: Optional[str] = Field(
        default=None,
        description="Color for horizontal/vertical divider lines. Falls back to colors.text_light.",
    )
    divider_line_width: float = Field(
        default=0.75,
        description="Width of divider lines in points.",
    )
    accent_bar_color: Optional[str] = Field(
        default=None,
        description="Color for accent bars / highlight strips. Falls back to colors.primary.",
    )
    accent_bar_height: float = Field(
        default=0.06,
        description="Height of accent bars in inches.",
    )
    badge_fill_color: Optional[str] = Field(
        default=None,
        description="Fill color for numbered badges (e.g., #CDF986). Falls back to colors.primary.",
    )
    badge_text_color: Optional[str] = Field(
        default=None,
        description="Text color inside badges. Falls back to colors.secondary.",
    )
    badge_size: float = Field(
        default=0.4,
        description="Badge diameter/size in inches.",
    )
    image_placeholder_fill: Optional[str] = Field(
        default=None,
        description="Fill color for image placeholder areas. Falls back to colors.surface.",
    )


# ---------------------------------------------------------------------------
# Paragraph / typography formatting
# ---------------------------------------------------------------------------

class ParagraphConfig(BaseModel):
    """Paragraph-level formatting defaults extracted from templates.

    Controls line spacing, paragraph spacing, and text alignment so
    generated slides match the template's typographic rhythm.
    """

    title_alignment: str = Field(
        default="left",
        description="Default alignment for titles: left, center, right",
    )
    body_alignment: str = Field(
        default="left",
        description="Default alignment for body text: left, center, right",
    )
    subtitle_alignment: str = Field(
        default="left",
        description="Default alignment for subtitles: left, center, right",
    )

    body_line_spacing: Optional[float] = Field(
        default=None,
        description="Line spacing multiplier for body text (e.g., 1.0, 1.15, 1.2). None = inherit from layout.",
    )
    title_line_spacing: Optional[float] = Field(
        default=None,
        description="Line spacing multiplier for title text. None = inherit from layout.",
    )
    bullet_line_spacing: Optional[float] = Field(
        default=None,
        description="Line spacing multiplier for bullet items. None = inherit from body_line_spacing.",
    )

    space_after_title: Optional[int] = Field(
        default=None,
        description="Space after title paragraphs in points. None = inherit from layout.",
    )
    space_after_body: Optional[int] = Field(
        default=None,
        description="Space after body paragraphs in points. None = inherit from layout.",
    )
    space_before_body: Optional[int] = Field(
        default=None,
        description="Space before body paragraphs in points. None = inherit from layout.",
    )
    space_after_bullet: Optional[int] = Field(
        default=None,
        description="Space after each bullet item in points. None = auto.",
    )

    bullet_character: str = Field(
        default="\u2022",
        description="Character used for bullet points.",
    )
    bullet_indent: Optional[float] = Field(
        default=None,
        description="Bullet indent in inches from the left margin.",
    )


# ---------------------------------------------------------------------------
# Content area / margins
# ---------------------------------------------------------------------------

class ContentAreaConfig(BaseModel):
    """Content area positioning — where text boxes should be placed.

    Extracted from typical placeholder positions in the template.
    Used when the builder needs to add manual text boxes (quotes,
    data points) outside of layout placeholders.
    """

    margin_left: float = Field(default=0.5, description="Left margin in inches")
    margin_top: float = Field(default=0.5, description="Top margin in inches")
    margin_right: float = Field(default=0.5, description="Right margin in inches")
    margin_bottom: float = Field(default=0.5, description="Bottom margin in inches")

    title_left: Optional[float] = Field(default=None, description="Title area left edge (inches)")
    title_top: Optional[float] = Field(default=None, description="Title area top edge (inches)")
    title_width: Optional[float] = Field(default=None, description="Title area width (inches)")
    title_height: Optional[float] = Field(default=None, description="Title area height (inches)")

    body_left: Optional[float] = Field(default=None, description="Body area left edge (inches)")
    body_top: Optional[float] = Field(default=None, description="Body area top edge (inches)")
    body_width: Optional[float] = Field(default=None, description="Body area width (inches)")
    body_height: Optional[float] = Field(default=None, description="Body area height (inches)")


# ---------------------------------------------------------------------------
# Slide dimensions
# ---------------------------------------------------------------------------

class SlideDimensions(BaseModel):
    """Slide dimensions extracted from the template."""

    width: float = Field(default=10.0, description="Slide width in inches")
    height: float = Field(default=5.625, description="Slide height in inches")


# ---------------------------------------------------------------------------
# Per-slide-type overrides
# ---------------------------------------------------------------------------

class SlideTypeOverrides(BaseModel):
    """Optional per-slide-type color overrides.

    These allow fine-tuning of how different slide types are styled.
    When not set, sensible defaults are derived from the base ColorConfig.
    """

    section_header_bg: Optional[str] = Field(
        default=None,
        description="Background color for section header slides. Defaults to colors.primary.",
    )
    data_point_accent: Optional[str] = Field(
        default=None,
        description="Accent color for hero numbers on data_point slides. Defaults to colors.accent.",
    )
    dark_slide_text: Optional[str] = Field(
        default=None,
        description="Text color used on dark-background slides. Defaults to colors.text_light.",
    )
    quote_bg: Optional[str] = Field(
        default=None,
        description="Background color for quote slides. Defaults to colors.primary.",
    )
    title_bg: Optional[str] = Field(
        default=None,
        description="Background color for title slides. Defaults to colors.secondary.",
    )
    content_bg: Optional[str] = Field(
        default=None,
        description="Background color for content slides. Defaults to colors.background.",
    )
    closing_bg: Optional[str] = Field(
        default=None,
        description="Background color for closing slides. Defaults to title_bg.",
    )
    bullet_list_bg: Optional[str] = Field(
        default=None,
        description="Background color for bullet list slides. Defaults to content_bg.",
    )


# ---------------------------------------------------------------------------
# Top-level design system
# ---------------------------------------------------------------------------

class DesignSystem(BaseModel):
    """Complete design system configuration.

    Captures all visual parameters needed to generate on-brand slides:
    fonts (with weight variants), colors, paragraph formatting, content
    area positioning, decoration parameters, slide dimensions, and
    per-slide-type overrides.
    """

    name: str = "Default"
    fonts: FontConfig = Field(default_factory=FontConfig)
    colors: ColorConfig = Field(default_factory=ColorConfig)
    paragraph: ParagraphConfig = Field(default_factory=ParagraphConfig)
    content_area: ContentAreaConfig = Field(default_factory=ContentAreaConfig)
    dimensions: SlideDimensions = Field(default_factory=SlideDimensions)
    overrides: SlideTypeOverrides = Field(default_factory=SlideTypeOverrides)
    decoration: DecorationConfig = Field(default_factory=DecorationConfig)
    logo_path: Optional[str] = None

    # --- Convenience accessors: slide-type backgrounds ---

    @property
    def section_header_bg(self) -> str:
        return self.overrides.section_header_bg or self.colors.secondary

    @property
    def data_point_accent(self) -> str:
        return self.overrides.data_point_accent or self.colors.accent

    @property
    def dark_slide_text(self) -> str:
        return self.overrides.dark_slide_text or self.colors.text_light

    @property
    def quote_bg(self) -> str:
        return self.overrides.quote_bg or self.colors.secondary

    @property
    def title_bg(self) -> str:
        return self.overrides.title_bg or self.colors.secondary

    @property
    def content_bg(self) -> str:
        return self.overrides.content_bg or self.colors.background

    @property
    def closing_bg(self) -> str:
        return self.overrides.closing_bg or self.title_bg

    @property
    def bullet_list_bg(self) -> str:
        return self.overrides.bullet_list_bg or self.content_bg

    # --- Convenience accessors: resolved colors ---

    @property
    def text_secondary_resolved(self) -> str:
        return self.colors.text_secondary or self.colors.text_dark

    @property
    def text_heading_resolved(self) -> str:
        return self.colors.text_heading or self.colors.text_dark

    @property
    def surface_resolved(self) -> str:
        return self.colors.surface or "#F4F4F4"

    @property
    def surface_accent_resolved(self) -> str:
        return self.colors.surface_accent or self.surface_resolved

    @property
    def brand_green_resolved(self) -> str:
        return self.colors.brand_green or self.colors.secondary

    # --- Convenience accessors: resolved decoration ---

    @property
    def divider_line_color_resolved(self) -> str:
        return self.decoration.divider_line_color or self.colors.text_light

    @property
    def accent_bar_color_resolved(self) -> str:
        return self.decoration.accent_bar_color or self.colors.primary

    @property
    def badge_fill_resolved(self) -> str:
        return self.decoration.badge_fill_color or self.colors.primary

    @property
    def badge_text_resolved(self) -> str:
        return self.decoration.badge_text_color or self.colors.secondary

    @property
    def image_placeholder_fill_resolved(self) -> str:
        return self.decoration.image_placeholder_fill or self.surface_resolved

    # --- Resolved font accessors ---

    @property
    def emphasis_font_resolved(self) -> str:
        """Font for emphasis/bold text — uses weight variant or falls back to title_font."""
        return self.fonts.emphasis_font or self.fonts.title_font

    @property
    def light_font_resolved(self) -> str:
        """Light font for subtitles — uses weight variant or falls back to body_font."""
        return self.fonts.light_font or self.fonts.body_font

    @property
    def medium_font_resolved(self) -> str:
        """Medium-weight font for labels/column heads."""
        return self.fonts.medium_font or self.fonts.body_font

    @property
    def extra_light_font_resolved(self) -> str:
        """Extra-light font for large display text / hero numbers."""
        return self.fonts.extra_light_font or self.light_font_resolved

    @property
    def label_font_resolved(self) -> str:
        """Font for small labels and section markers."""
        return self.fonts.label_font or self.medium_font_resolved

    @property
    def quote_font_resolved(self) -> str:
        """Font for quotes — uses weight variant or falls back to body_font."""
        return self.fonts.quote_font or self.fonts.body_font

    @property
    def quote_size_resolved(self) -> int:
        """Font size for quotes."""
        return self.fonts.quote_size or self.fonts.subtitle_size

    @property
    def data_point_size_resolved(self) -> int:
        """Font size for hero data point numbers."""
        return self.fonts.data_point_size or int(self.fonts.title_size * 1.5)

    @property
    def caption_size_resolved(self) -> int:
        """Font size for captions."""
        return self.fonts.caption_size or max(10, self.fonts.bullet_size - 2)

    @property
    def label_size_resolved(self) -> int:
        """Font size for small labels and section markers."""
        return self.fonts.label_size or 9

    @property
    def section_marker_size_resolved(self) -> int:
        """Font size for '01 | Topic' markers."""
        return self.fonts.section_marker_size or self.label_size_resolved

    @property
    def hero_number_size_resolved(self) -> int:
        """Font size for oversized hero numbers."""
        return self.fonts.hero_number_size or self.data_point_size_resolved

    @classmethod
    def from_yaml(cls, path: str | Path) -> "DesignSystem":
        """Load design system from a YAML configuration file."""
        path = Path(path)
        if not path.exists():
            raise FileNotFoundError(f"Design system file not found: {path}")
        with open(path) as f:
            data = yaml.safe_load(f)
        return cls.model_validate(data)

    def to_yaml(self, path: str | Path) -> None:
        """Save design system to a YAML configuration file."""
        path = Path(path)
        data = self.model_dump(exclude_none=True)
        with open(path, "w") as f:
            yaml.dump(data, f, default_flow_style=False, sort_keys=False)
