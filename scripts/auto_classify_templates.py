#!/usr/bin/env python3
"""Auto-classify template slides from descriptions JSON.

Produces richer classifications including content_keywords, visual_elements,
suitable_for, and multi-sentence descriptions. Uses heuristics on the extracted
structural metadata and text content.

Usage:
    python scripts/auto_classify_templates.py workspace/template_descriptions.json \
        -o workspace/classifications.json
"""

import argparse
import json
import re
import sys
from pathlib import Path


def _tokenize(text: str) -> list[str]:
    """Extract lowercase word tokens from text."""
    if not text:
        return []
    return re.findall(r"[a-z][a-z0-9]+", text.lower())


_STOP_WORDS = {
    "the", "a", "an", "is", "are", "was", "were", "be", "been", "being",
    "have", "has", "had", "do", "does", "did", "will", "would", "shall",
    "should", "may", "might", "must", "can", "could", "of", "in", "to",
    "for", "with", "on", "at", "by", "from", "as", "into", "through",
    "and", "or", "but", "not", "this", "that", "these", "those", "it",
    "its", "they", "them", "their", "we", "our", "you", "your", "about",
    "which", "what", "when", "where", "who", "how", "than", "then",
    "each", "every", "all", "both", "few", "more", "most", "other",
    "some", "such", "no", "nor", "only", "own", "same", "so", "too",
    "very", "just", "because", "if", "while", "after", "before", "during",
    "between", "up", "out", "off", "over", "under", "again", "further",
    "also", "still", "here", "there", "where", "google", "shape",
    "slide", "layout", "new", "like", "per", "use", "using", "used",
}


def _extract_keywords(text: str, max_kw: int = 10) -> list[str]:
    """Extract meaningful keywords from text, sorted by frequency."""
    tokens = _tokenize(text)
    freq: dict[str, int] = {}
    for t in tokens:
        if t not in _STOP_WORDS and len(t) > 2:
            freq[t] = freq.get(t, 0) + 1
    # Sort by frequency descending, take top N
    sorted_kw = sorted(freq, key=lambda k: freq[k], reverse=True)
    return sorted_kw[:max_kw]


def _classify_slide_type(slide: dict) -> str:
    """Classify slide type based on structural metadata."""
    phs = slide.get("placeholders", [])
    ph_types = {ph["type"].upper() for ph in phs}
    shape_count = slide.get("shape_count", 0)
    has_images = slide.get("has_images", False)
    layout_name = slide.get("layout_name", "").lower()
    text_content = slide.get("text_content", {})
    all_text = (text_content.get("all_text", "") if text_content else "").lower()

    # Title slides: have TITLE+SUBTITLE or CENTER_TITLE+SUBTITLE, few shapes
    if ("TITLE" in ph_types or "CENTER_TITLE" in ph_types) and "SUBTITLE" in ph_types:
        if "BODY" not in ph_types and "OBJECT" not in ph_types:
            return "title"

    # Section header: title-only or title+subtitle with minimal body, often dark bg
    if "section" in layout_name or "divider" in layout_name:
        return "section_header"
    if ("TITLE" in ph_types or "CENTER_TITLE" in ph_types) and "BODY" not in ph_types and "OBJECT" not in ph_types:
        if shape_count <= 8 and not has_images:
            return "section_header"

    # Quote: look for quote markers in text
    if any(c in all_text for c in ["\u201c", "\u201d", "\u2018", "\u2019"]) or "quote" in layout_name:
        if shape_count <= 10:
            return "quote"

    # Two-column: layout name or multiple body placeholders
    if "two_column" in layout_name or "two column" in layout_name:
        return "two_column"
    body_count = sum(1 for ph in phs if ph["type"].upper() in ("BODY", "OBJECT"))
    if body_count >= 2:
        return "two_column"

    # Chart: look for chart keywords in text or layout
    if "chart" in layout_name or "graph" in all_text:
        return "chart"

    # Image-heavy slides
    if has_images and shape_count >= 5 and "BODY" not in ph_types:
        return "image_with_text"
    if has_images and not phs:
        return "image_full"

    # Team slide: look for names/people keywords
    if any(kw in all_text for kw in ["team", "leadership", "about us", "our people"]):
        return "team"

    # Closing: look for closing keywords
    if any(kw in all_text for kw in ["thank you", "questions", "contact", "next steps", "get in touch"]):
        return "closing"

    # Comparison
    if any(kw in all_text for kw in ["vs", "versus", "compare", "comparison", "before", "after"]):
        return "comparison"

    # Timeline
    if any(kw in all_text for kw in ["timeline", "roadmap", "milestones", "phases"]):
        return "timeline"

    # Bullet list: body placeholder with many lines
    if "BODY" in ph_types or "OBJECT" in ph_types:
        body_text = text_content.get("body", "") if text_content else ""
        if body_text.count("\n") >= 3 or body_text.count("•") >= 2:
            return "bullet_list"
        return "content"

    # Default: content
    return "content"


def _classify_tags(slide: dict, slide_type: str) -> list[str]:
    """Derive semantic tags from metadata."""
    tags = []
    has_images = slide.get("has_images", False)
    has_background = slide.get("has_background", False)
    shape_count = slide.get("shape_count", 0)
    complexity = min(5, max(1, shape_count // 5 + 1))

    if has_background:
        tags.append("branded")
    if has_images:
        tags.append("visual")
    if shape_count > 15:
        tags.append("complex")
    elif shape_count <= 5:
        tags.append("minimal")
    if complexity >= 4:
        tags.append("data-heavy")

    # Style tags from fonts
    fonts = slide.get("font_families", [])
    font_str = " ".join(fonts).lower()
    if "inter" in font_str or "helvetica" in font_str:
        tags.append("corporate")
    if "poppins" in font_str or "montserrat" in font_str:
        tags.append("modern")

    if slide_type in ("section_header", "title", "closing"):
        tags.append("bold")
    if not has_images and shape_count <= 8:
        tags.append("text-heavy")

    # Dark background detection
    colors = slide.get("color_scheme", [])
    dark_colors = [c for c in colors if _is_dark_hex(c)]
    if has_background and len(dark_colors) >= 1:
        tags.append("dark-bg")

    return list(set(tags))[:8]


def _is_dark_hex(hex_color: str) -> bool:
    """Check if a hex color is dark."""
    h = hex_color.lstrip("#")
    try:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return (0.299 * r + 0.587 * g + 0.114 * b) < 80
    except (ValueError, IndexError):
        return False


def _derive_visual_elements(slide: dict) -> list[str]:
    """Derive visual element descriptors."""
    elements = []
    has_images = slide.get("has_images", False)
    has_background = slide.get("has_background", False)
    shape_count = slide.get("shape_count", 0)
    text_content = slide.get("text_content", {})
    all_text = (text_content.get("all_text", "") if text_content else "").lower()

    if has_background:
        elements.append("branded background")
    if has_images:
        elements.append("imagery")
    if shape_count > 15:
        elements.append("complex layout")

    # Detect specific visual types from text/context
    if any(kw in all_text for kw in ["chart", "graph", "bar", "pie", "line chart"]):
        elements.append("chart")
    if any(kw in all_text for kw in ["diagram", "flow", "process", "arrow"]):
        elements.append("diagram")
    if any(kw in all_text for kw in ["icon", "icons"]):
        elements.append("icons")
    if any(kw in all_text for kw in ["screenshot", "demo", "ui "]):
        elements.append("screenshot")
    if any(kw in all_text for kw in ["stat", "metric", "%", "number"]):
        elements.append("stat callout")

    # Look for big numbers (standalone numbers as visual elements)
    import re
    if re.search(r'\b\d{2,}[%+]?\b', all_text):
        elements.append("big number")

    return elements[:6]


def _derive_suitable_for(slide: dict, slide_type: str) -> list[str]:
    """Derive what content intents this slide is suitable for."""
    suitable = []
    text_content = slide.get("text_content", {})
    all_text = (text_content.get("all_text", "") if text_content else "").lower()
    has_images = slide.get("has_images", False)
    shape_count = slide.get("shape_count", 0)

    # Type-based defaults
    type_map = {
        "title": ["opening", "deck introduction"],
        "section_header": ["section transition", "chapter break"],
        "closing": ["closing", "call to action", "next steps"],
        "quote": ["key insight", "testimonial", "thought leadership"],
        "two_column": ["comparison", "side-by-side analysis"],
        "comparison": ["comparison", "before-after", "pros-cons"],
        "chart": ["data presentation", "metrics overview"],
        "image_full": ["visual impact", "product showcase"],
        "image_with_text": ["case study", "product feature"],
        "team": ["team introduction", "about us"],
        "timeline": ["roadmap", "project phases", "milestones"],
        "bullet_list": ["key points", "agenda", "summary"],
        "content": ["general content", "explanation"],
    }
    suitable.extend(type_map.get(slide_type, ["general content"]))

    # Content-based additions
    if any(kw in all_text for kw in ["case study", "customer", "brand", "story"]):
        suitable.append("case study")
    if any(kw in all_text for kw in ["data", "metric", "stat", "revenue", "growth", "%"]):
        suitable.append("data presentation")
    if any(kw in all_text for kw in ["framework", "model", "approach", "methodology"]):
        suitable.append("framework overview")
    if any(kw in all_text for kw in ["agenda", "outline", "topics", "overview"]):
        suitable.append("agenda")
    if any(kw in all_text for kw in ["demo", "product", "feature", "platform"]):
        suitable.append("product feature")

    return list(set(suitable))[:5]


def _build_description(slide: dict, slide_type: str, keywords: list[str]) -> str:
    """Build a 2-3 sentence description covering structure and content."""
    parts = []

    # Structural description
    phs = slide.get("placeholders", [])
    shape_count = slide.get("shape_count", 0)
    has_images = slide.get("has_images", False)
    has_background = slide.get("has_background", False)
    layout_name = slide.get("layout_name", "")

    type_labels = {
        "title": "Title slide",
        "section_header": "Section header/divider slide",
        "content": "Content slide",
        "bullet_list": "Bullet list slide",
        "two_column": "Two-column layout slide",
        "comparison": "Comparison slide",
        "quote": "Quote/callout slide",
        "chart": "Chart/data visualization slide",
        "image_full": "Full-bleed image slide",
        "image_with_text": "Image with text slide",
        "timeline": "Timeline/roadmap slide",
        "team": "Team/people slide",
        "closing": "Closing/CTA slide",
    }
    label = type_labels.get(slide_type, "Content slide")

    struct_parts = []
    if has_background:
        struct_parts.append("branded background")
    if has_images:
        struct_parts.append("imagery")
    if len(phs) > 0:
        struct_parts.append(f"{len(phs)} placeholders")
    struct_parts.append(f"{shape_count} shapes")

    parts.append(f"{label} with {', '.join(struct_parts)}.")

    # Content description from text
    text_content = slide.get("text_content", {})
    if text_content:
        title = text_content.get("title", "")
        body = text_content.get("body", "")
        if title:
            # Truncate long titles
            if len(title) > 80:
                title = title[:80] + "..."
            parts.append(f"Title: \"{title}\".")
        if keywords:
            parts.append(f"Topics: {', '.join(keywords[:5])}.")

    return " ".join(parts)


def classify_all(slides: list[dict]) -> list[dict]:
    """Classify all slides and return rich classification objects."""
    classifications = []

    for slide in slides:
        slide_type = _classify_slide_type(slide)

        # Compute complexity from shape count
        shape_count = slide.get("shape_count", 0)
        complexity = min(5, max(1, (shape_count - 1) // 4 + 1))

        # Extract keywords from text content
        text_content = slide.get("text_content", {})
        all_text = text_content.get("all_text", "") if text_content else ""
        keywords = _extract_keywords(all_text)

        tags = _classify_tags(slide, slide_type)
        visual_elements = _derive_visual_elements(slide)
        suitable_for = _derive_suitable_for(slide, slide_type)
        description = _build_description(slide, slide_type, keywords)

        classifications.append({
            "slide_type": slide_type,
            "tags": tags,
            "complexity": complexity,
            "description": description,
            "content_keywords": keywords,
            "visual_elements": visual_elements,
            "suitable_for": suitable_for,
        })

    return classifications


def main():
    parser = argparse.ArgumentParser(description="Auto-classify template slides")
    parser.add_argument("descriptions", type=Path, help="Path to template_descriptions.json")
    parser.add_argument(
        "-o", "--output", type=Path,
        default=Path("workspace/classifications.json"),
        help="Output classifications JSON path",
    )
    args = parser.parse_args()

    if not args.descriptions.exists():
        print(f"Error: File not found: {args.descriptions}", file=sys.stderr)
        sys.exit(1)

    data = json.loads(args.descriptions.read_text(encoding="utf-8"))
    slides = data.get("slides", [])
    print(f"Classifying {len(slides)} slides...")

    classifications = classify_all(slides)

    result = {"classifications": classifications}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2), encoding="utf-8")

    # Stats
    type_counts: dict[str, int] = {}
    for cls in classifications:
        t = cls["slide_type"]
        type_counts[t] = type_counts.get(t, 0) + 1

    print(f"Written to: {args.output}")
    print(f"\nSlide type distribution:")
    for t, count in sorted(type_counts.items(), key=lambda x: -x[1]):
        print(f"  {t:20s}: {count}")


if __name__ == "__main__":
    main()
