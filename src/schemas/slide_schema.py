"""Pydantic models for slide structure and deck specification."""

from enum import Enum
from typing import Literal, Optional

from pydantic import BaseModel, Field


class SlideType(str, Enum):
    """Supported slide layout types."""

    TITLE = "title"
    SECTION_HEADER = "section_header"
    CONTENT = "content"
    TWO_COLUMN = "two_column"
    BULLET_LIST = "bullet_list"
    IMAGE_FULL = "image_full"
    IMAGE_WITH_TEXT = "image_with_text"
    CHART = "chart"
    QUOTE = "quote"
    COMPARISON = "comparison"
    TIMELINE = "timeline"
    TEAM = "team"
    DATA_POINT = "data_point"
    CLOSING = "closing"


class ContentBlock(BaseModel):
    """A single block of content within a slide."""

    type: Literal["title", "subtitle", "body", "bullets", "quote", "caption", "data_point"]
    content: str
    emphasis: Literal["normal", "bold", "highlight"] = "normal"


class SlideSpec(BaseModel):
    """Specification for a single slide in the deck."""

    slide_number: int
    slide_type: SlideType
    intent: str = Field(description="What this slide should accomplish for the audience")
    title: Optional[str] = None
    subtitle: Optional[str] = None
    content_blocks: list[ContentBlock] = Field(default_factory=list)
    speaker_notes: Optional[str] = None
    image_suggestions: list[str] = Field(default_factory=list)
    layout_hints: list[str] = Field(
        default_factory=list,
        description="Hints like 'emphasize_title', 'minimal_text', 'large_visual'",
    )
    visual_profile: Optional[str] = Field(
        default=None,
        description="Desired visual feel: 'dark', 'light', 'branded_image', 'minimal'",
    )


class DeckSchema(BaseModel):
    """Complete specification for a presentation deck."""

    title: str
    subtitle: Optional[str] = None
    target_audience: str
    key_message: str
    slides: list[SlideSpec] = Field(default_factory=list)


class ContentInventory(BaseModel):
    """Structured content extracted from an input document."""

    main_topic: str
    themes: list[str] = Field(default_factory=list)
    sections: list["ContentSection"] = Field(default_factory=list)
    key_data_points: list[str] = Field(default_factory=list)
    quotes: list[str] = Field(default_factory=list)
    summary: str = ""


class ContentSection(BaseModel):
    """A logical section of content from the source document."""

    heading: str
    content: str
    bullet_points: list[str] = Field(default_factory=list)
    importance: Literal["high", "medium", "low"] = "medium"


class ContentMaturity(BaseModel):
    """Assessment of input content maturity level and required pipeline stages."""

    maturity_level: int = Field(ge=1, le=4, description="1=raw ideas, 2=outline, 3=draft, 4=ready")
    maturity_label: Literal["raw_ideas", "outline", "draft", "presentation_ready"]
    reasoning: str = Field(description="2-3 sentence explanation of the assessment")
    pipeline_stages: list[Literal["research", "content_development", "editor", "design"]] = Field(
        description="Which stages to run, in order"
    )
    content_gaps: list[str] = Field(default_factory=list, description="What's missing or weak")
    strengths: list[str] = Field(default_factory=list, description="What's already strong")
    word_count: int = Field(ge=0, description="Approximate word count of parsed content")
    section_count: int = Field(ge=0, description="Number of identifiable sections or topics")


# Rebuild models that have forward references
ContentInventory.model_rebuild()
