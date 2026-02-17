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
    sub_sections: list["ContentSection"] = Field(
        default_factory=list,
        description="Nested sub-sections for hierarchical content structure",
    )
    estimated_slides: Optional[int] = Field(
        default=None,
        description="Estimated number of slides this section warrants (set by content planner)",
    )


class ContentMaturity(BaseModel):
    """Assessment of input content maturity level and required pipeline stages."""

    maturity_level: int = Field(ge=1, le=4, description="1=raw ideas, 2=outline, 3=draft, 4=ready")
    maturity_label: Literal["raw_ideas", "outline", "draft", "presentation_ready"]
    reasoning: str = Field(description="2-3 sentence explanation of the assessment")
    pipeline_stages: list[
        Literal["research", "content_planning", "content_development", "editor", "design"]
    ] = Field(description="Which stages to run, in order")
    content_gaps: list[str] = Field(default_factory=list, description="What's missing or weak")
    strengths: list[str] = Field(default_factory=list, description="What's already strong")
    word_count: int = Field(ge=0, description="Approximate word count of parsed content")
    section_count: int = Field(ge=0, description="Number of identifiable sections or topics")


# ---------------------------------------------------------------------------
# Deck Planning models (produced by the Content Planner agent)
# ---------------------------------------------------------------------------


class SlidePlan(BaseModel):
    """Plan for a single slide within a section."""

    working_title: str = Field(description="Draft title capturing the slide's core point")
    slide_type: SlideType
    content_source: Literal["existing", "expand", "create"] = Field(
        description=(
            "'existing' = content is ready in the inventory, "
            "'expand' = bullet/idea exists but needs development, "
            "'create' = new content must be authored from scratch"
        )
    )
    content_brief: str = Field(
        description="What this slide should contain — specific enough to write from"
    )
    visual_intent: str = Field(
        default="",
        description="Description of the ideal visual treatment for this slide",
    )
    evidence_needed: list[str] = Field(
        default_factory=list,
        description="Data points, examples, or quotes to find or create",
    )


class SectionPlan(BaseModel):
    """Plan for how a content section maps to one or more slides."""

    section_ref: str = Field(description="Heading from ContentInventory this plan maps to")
    importance: Literal["high", "medium", "low"]
    depth: Literal["deep", "standard", "light", "skip"] = Field(
        description=(
            "'deep' = 3+ slides (central topic), "
            "'standard' = 1-2 slides (supporting point), "
            "'light' = shared slide (minor point combined with another section), "
            "'skip' = not included in the deck"
        )
    )
    slide_count: int = Field(ge=0, description="Number of slides allocated to this section")
    slide_plan: list[SlidePlan] = Field(
        default_factory=list,
        description="Per-slide breakdown for this section",
    )
    expansion_needed: list[str] = Field(
        default_factory=list,
        description="Content that must be researched or created for this section",
    )
    merge_with: Optional[str] = Field(
        default=None,
        description="Heading of another section to combine with (for 'light' depth)",
    )


class DeckPlan(BaseModel):
    """Structural blueprint for the entire deck.

    Produced by the Content Planner agent between content extraction and
    content development. Drives all downstream content and design decisions.
    """

    thesis: str = Field(description="The single core argument or message of this presentation")
    narrative_arc: str = Field(
        description="Structural pattern: 'problem_solution', 'journey', 'educational', 'strategic'"
    )
    target_slide_count: int = Field(
        description="Total slides planned for the deck (sum of all section slide_counts + bookends)"
    )
    sections: list[SectionPlan] = Field(
        default_factory=list,
        description="Plans for sections that exist in the content inventory",
    )
    new_sections_needed: list[SectionPlan] = Field(
        default_factory=list,
        description="Plans for sections not present in the input but needed for the narrative",
    )
    visual_strategy: str = Field(
        default="",
        description="Overall visual approach — guides template selection and design decisions",
    )


# Rebuild models that have forward references
ContentInventory.model_rebuild()
ContentSection.model_rebuild()
DeckPlan.model_rebuild()
