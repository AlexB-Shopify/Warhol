"""Pydantic models for template metadata and registry."""

from pathlib import Path
from typing import Literal, Optional

from pydantic import BaseModel, Field

from .slide_schema import SlideType


class PlaceholderInfo(BaseModel):
    """Metadata about a single placeholder in a template slide."""

    name: str
    type: str = Field(description="Placeholder type: title, body, picture, chart, subtitle, other")
    position: tuple[float, float, float, float] = Field(
        description="(left, top, width, height) in inches"
    )


class ContentZone(BaseModel):
    """A replaceable text area within a template slide.

    Content zones identify which shapes hold text that should be replaced
    when the template is cloned for a new slide. Design elements (images,
    decorations, backgrounds) are NOT content zones and are preserved as-is.
    """

    zone_type: Literal["title", "body", "subtitle", "bullet_area", "data_point", "caption"] = "body"
    shape_name: str = Field(description="Name of the shape in the slide (for targeting)")
    position: tuple[float, float, float, float] = Field(
        description="(left, top, width, height) in inches"
    )
    max_chars: int = Field(default=200, description="Approximate max character capacity")
    font_size_range: tuple[int, int] = Field(
        default=(10, 44), description="(min, max) font size in points"
    )


class TextContent(BaseModel):
    """Extracted text content from a template slide."""

    title: str = Field(default="", description="Text from title placeholders")
    body: str = Field(default="", description="Text from body/subtitle placeholders and shapes")
    all_text: str = Field(default="", description="All visible text concatenated")


class TemplateSlide(BaseModel):
    """Metadata for a single slide within a template file."""

    template_file: str
    slide_index: int
    slide_type: SlideType
    layout_name: str = ""
    placeholders: list[PlaceholderInfo] = Field(default_factory=list)
    color_scheme: list[str] = Field(default_factory=list, description="Hex colors used")
    font_families: list[str] = Field(default_factory=list, description="Font families used")
    tags: list[str] = Field(
        default_factory=list,
        description="Semantic tags like 'corporate', 'bold', 'minimal'",
    )
    complexity: int = Field(default=1, ge=1, le=5, description="Visual complexity 1-5")
    shape_count: int = 0
    has_images: bool = False
    has_background: bool = False
    description: str = Field(default="", description="LLM-generated layout description")
    embedding: Optional[list[float]] = Field(
        default=None, description="Semantic embedding for matching"
    )

    # --- Semantic fields ---
    text_content: Optional[TextContent] = Field(
        default=None,
        description="Extracted text content from the slide (title, body, all_text)",
    )
    content_keywords: list[str] = Field(
        default_factory=list,
        description="Topic keywords extracted from slide text (e.g., 'revenue', 'growth', 'checkout')",
    )
    visual_elements: list[str] = Field(
        default_factory=list,
        description="Visual element descriptors (e.g., 'bar chart', 'process diagram', 'photo grid')",
    )
    suitable_for: list[str] = Field(
        default_factory=list,
        description="Content intents this slide works well for (e.g., 'data presentation', 'case study')",
    )

    # --- Content zone mapping (for clone-and-replace) ---
    content_zones: list[ContentZone] = Field(
        default_factory=list,
        description="Replaceable text areas — shapes where content can be swapped during clone-and-replace",
    )
    background_type: Literal["solid", "gradient", "image", "master_inherited", "none"] = Field(
        default="none",
        description="Type of background on this slide",
    )
    visual_profile: Literal["dark", "light", "branded_image", "minimal"] = Field(
        default="minimal",
        description="Overall visual character of the slide",
    )
    content_capacity: Literal["low", "medium", "high"] = Field(
        default="medium",
        description="How much text content this slide can hold",
    )
    image_type: Literal["none", "decorative", "content"] = Field(
        default="none",
        description=(
            "Image classification: 'none' = no images, "
            "'decorative' = branded backgrounds/logos/accents (safe to keep on clone), "
            "'content' = product shots/screenshots/photos (irrelevant when cloned for different content)"
        ),
    )
    image_count: int = Field(
        default=0,
        description="Number of picture shapes on the slide",
    )


class TemplateRegistry(BaseModel):
    """Complete registry of all analyzed templates."""

    templates: list[TemplateSlide] = Field(default_factory=list)
    source_files: list[str] = Field(
        default_factory=list, description="All .pptx files that were analyzed"
    )

    def find_by_type(self, slide_type: SlideType) -> list[TemplateSlide]:
        """Find templates matching a given slide type."""
        return [t for t in self.templates if t.slide_type == slide_type]

    def find_by_tags(self, tags: list[str]) -> list[TemplateSlide]:
        """Find templates that have at least one of the given tags."""
        tag_set = set(tags)
        return [t for t in self.templates if tag_set & set(t.tags)]

    def save(self, path: str | Path) -> None:
        """Serialize registry to JSON file."""
        Path(path).write_text(self.model_dump_json(indent=2))

    @classmethod
    def load(cls, path: str | Path) -> "TemplateRegistry":
        """Load registry from JSON file."""
        return cls.model_validate_json(Path(path).read_text())
