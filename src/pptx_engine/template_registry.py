"""Template registry loader and lookup for the PPTX engine."""

from pathlib import Path

from src.schemas.template_schema import TemplateRegistry, TemplateSlide
from src.schemas.slide_schema import SlideType


def load_registry(path: str | Path) -> TemplateRegistry:
    """Load a template registry from a JSON file."""
    return TemplateRegistry.load(path)


def find_best_match(
    registry: TemplateRegistry,
    slide_type: SlideType,
    tags: list[str] | None = None,
    exclude_indices: set[int] | None = None,
) -> TemplateSlide | None:
    """Find the best template match for a slide type and optional tags.

    Args:
        registry: The template registry to search.
        slide_type: The type of slide to match.
        tags: Optional semantic tags to prefer.
        exclude_indices: Template indices to skip (for variety).

    Returns:
        Best matching TemplateSlide, or None if no match found.
    """
    exclude_indices = exclude_indices or set()

    # First pass: match on type
    candidates = [
        t
        for i, t in enumerate(registry.templates)
        if t.slide_type == slide_type and i not in exclude_indices
    ]

    if not candidates:
        # Fallback: any template not excluded
        candidates = [
            t
            for i, t in enumerate(registry.templates)
            if i not in exclude_indices
        ]

    if not candidates:
        # Last resort: any template at all
        candidates = registry.templates

    if not candidates:
        return None

    # Score candidates by tag overlap
    if tags:
        tag_set = set(tags)
        scored = [(t, len(tag_set & set(t.tags))) for t in candidates]
        scored.sort(key=lambda x: x[1], reverse=True)
        return scored[0][0]

    return candidates[0]


def get_template_variety_tracker() -> dict[str, int]:
    """Create a tracker to ensure layout variety across a deck.

    Returns a dict mapping template_file:slide_index to usage count.
    """
    return {}


def record_usage(tracker: dict[str, int], template: TemplateSlide) -> None:
    """Record that a template was used, for variety tracking."""
    key = f"{template.template_file}:{template.slide_index}"
    tracker[key] = tracker.get(key, 0) + 1
