#!/usr/bin/env python3
"""Algorithmic template matching — scores template slides against deck schema slides.

Produces a ranked match for each slide in the deck schema using a weighted scoring
function that combines:
  1. Slide type compatibility    (12%)
  2. Content/keyword similarity  (8%)
  3. Structural compatibility    (30%)
  4. Visual richness alignment   (25%)
  5. Tag overlap                 (10%)
  6. Content capacity fit        (15%)
  + Image relevance penalty      (additive)
  + Consistency groups for section headers and bookend slides

Usage:
    python scripts/match_templates.py workspace/deck_schema.json template_registry.json \
        -o workspace/template_matches.json [--threshold 0.45] [--max-dropin-pct 0.35]
"""

import argparse
import json
import re
import sys
import unicodedata
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.schemas.slide_schema import DeckSchema, SlideSpec, SlideType
from src.schemas.template_schema import TemplateRegistry, TemplateSlide


# -----------------------------------------------------------------------
# Slide-type compatibility matrix
# -----------------------------------------------------------------------

# Groups of compatible slide types (same structural family)
_TYPE_COMPAT = {
    "title": {"title", "closing"},
    "closing": {"title", "closing"},
    "section_header": {"section_header"},
    "content": {"content", "bullet_list", "image_with_text", "timeline"},
    "bullet_list": {"content", "bullet_list"},
    "two_column": {"two_column", "comparison"},
    "comparison": {"two_column", "comparison"},
    "quote": {"quote"},
    "chart": {"chart", "content"},
    "image_full": {"image_full", "image_with_text"},
    "image_with_text": {"image_with_text", "content", "image_full"},
    "timeline": {"timeline", "content"},
    "team": {"team"},
}

# Known incompatible dimensions (source files that are NOT 10x5.625)
_INCOMPATIBLE_FILES = {
    "2025 Enterprise Second Pitch Deck",
    "L'Oréal_Shopify",
    "basic.pptx",
}

# Target: slides from the Workshop Slide Bank get a small preference boost
_PREFERRED_SOURCE = "Shopify - Example Technical Workshop Slide Bank"

# -----------------------------------------------------------------------
# Consistency groups — slide types that should use the same template
# throughout a deck for visual coherence
# -----------------------------------------------------------------------

_CONSISTENCY_GROUPS: dict[str, str] = {
    # key = slide type, value = group name
    "section_header": "section_header",
    "title": "bookend",
    "closing": "bookend",
}

# Capacity ordinal for comparison
_CAPACITY_ORDINAL = {"low": 0, "medium": 1, "high": 2}


# -----------------------------------------------------------------------
# Tokenizer for keyword similarity
# -----------------------------------------------------------------------

_STOP_WORDS = {
    "the", "a", "an", "is", "are", "was", "were", "be", "been", "being",
    "have", "has", "had", "do", "does", "did", "will", "would", "shall",
    "should", "may", "might", "must", "can", "could", "of", "in", "to",
    "for", "with", "on", "at", "by", "from", "as", "into", "through",
    "and", "or", "but", "not", "this", "that", "these", "those", "it",
    "its", "they", "them", "their", "we", "our", "you", "your", "i",
    "my", "me", "he", "she", "his", "her", "slide", "slides", "content",
    "about", "using", "how", "what", "when", "why", "where", "which",
}


def _tokenize(text: str) -> set[str]:
    """Extract meaningful lowercase tokens from text, removing stop words."""
    if not text:
        return set()
    tokens = set(re.findall(r"[a-z0-9]+", text.lower()))
    return tokens - _STOP_WORDS


# -----------------------------------------------------------------------
# Scoring components
# -----------------------------------------------------------------------

def _score_type_match(deck_type: str, template_type: str) -> float:
    """Score slide type compatibility (0.0 - 1.0)."""
    if deck_type == template_type:
        return 1.0
    compat = _TYPE_COMPAT.get(deck_type, {deck_type})
    if template_type in compat:
        return 0.6
    return 0.0


def _template_text_weight(template: TemplateSlide) -> float:
    """How much visible text does this template have? (0.0 = no text, 1.0 = heavy text).

    Templates with heavy text are BAD drop-in candidates because the audience
    will read the wrong content. Templates that are primarily visual (backgrounds,
    images, branded graphics with minimal text) are GOOD drop-in candidates.
    """
    all_text = ""
    if template.text_content:
        all_text = template.text_content.all_text or ""

    word_count = len(all_text.split())
    if word_count <= 5:
        return 0.0   # Barely any text — great for drop-in
    elif word_count <= 15:
        return 0.3
    elif word_count <= 40:
        return 0.6
    else:
        return 1.0   # Heavy text — bad for drop-in


def _score_content_similarity(
    deck_slide: SlideSpec,
    template: TemplateSlide,
) -> float:
    """Score content/keyword similarity between deck slide and template (0.0 - 1.0).

    KEY INSIGHT: A drop-in slide's existing text IS the content the audience sees.
    If the template text is about a different topic, the match is HARMFUL, not neutral.

    Strategy:
    - Low-text templates (visual/branded): score based on structural suitability
    - High-text templates: score on actual topic overlap, with PENALTY for mismatch
    """
    text_weight = _template_text_weight(template)

    # If template has minimal text, it's a visual/branded slide — suitable
    # for any topic. Score based on type/intent suitability instead.
    if text_weight <= 0.3:
        # Use suitable_for matching against intent
        intent_parts = []
        if deck_slide.intent:
            intent_parts.append(deck_slide.intent)
        if deck_slide.title:
            intent_parts.append(deck_slide.title)
        intent_tokens = _tokenize(" ".join(intent_parts))

        suitable_tokens = _tokenize(" ".join(template.suitable_for))
        if intent_tokens and suitable_tokens:
            overlap = intent_tokens & suitable_tokens
            return min(1.0, len(overlap) / max(3, len(suitable_tokens)))
        return 0.3  # Neutral for visual-only slides

    # For text-heavy templates, require actual topic overlap
    deck_parts = []
    if deck_slide.title:
        deck_parts.append(deck_slide.title)
    if deck_slide.intent:
        deck_parts.append(deck_slide.intent)
    for block in deck_slide.content_blocks:
        deck_parts.append(block.content)
    deck_tokens = _tokenize(" ".join(deck_parts))

    if not deck_tokens:
        return 0.0

    tmpl_parts = []
    tmpl_parts.extend(template.content_keywords)
    if template.text_content:
        if template.text_content.title:
            tmpl_parts.append(template.text_content.title)
        if template.text_content.body:
            tmpl_parts.append(template.text_content.body)
    tmpl_tokens = _tokenize(" ".join(tmpl_parts))

    if not tmpl_tokens:
        return 0.0

    # Overlap relative to the SMALLER set (precision-oriented)
    intersection = deck_tokens & tmpl_tokens
    smaller = min(len(deck_tokens), len(tmpl_tokens))
    if smaller == 0:
        return 0.0

    precision = len(intersection) / smaller

    # Apply penalty proportional to text weight:
    # heavy-text templates that DON'T match get penalized hard
    if precision < 0.15:
        return -0.2 * text_weight  # Negative score = actively harmful match

    return precision * (1.0 - 0.3 * text_weight)  # Dampen score for text-heavy slides


def _estimate_content_volume(deck_slide: SlideSpec) -> int:
    """Estimate the total character count of a deck slide's content."""
    total = len(deck_slide.title or "")
    total += len(deck_slide.subtitle or "")
    for block in deck_slide.content_blocks:
        total += len(block.content)
    return total


def _score_content_capacity_fit(
    deck_slide: SlideSpec,
    template: TemplateSlide,
) -> float:
    """Score how well the deck slide's content volume fits the template capacity (0.0 - 1.0).

    A text-heavy slide matched to a low-capacity (image-heavy) template produces
    bare-looking output. Conversely, a single-line slide on a high-capacity template
    wastes space. This score penalises both mismatches.
    """
    content_blocks = len(deck_slide.content_blocks)
    char_count = _estimate_content_volume(deck_slide)
    tmpl_cap = _CAPACITY_ORDINAL.get(template.content_capacity, 1)

    # Estimate deck slide's "needed capacity": low / medium / high
    if content_blocks <= 1 and char_count < 100:
        needed_cap = 0  # low
    elif content_blocks <= 3 and char_count < 400:
        needed_cap = 1  # medium
    else:
        needed_cap = 2  # high

    diff = abs(needed_cap - tmpl_cap)
    if diff == 0:
        return 1.0
    elif diff == 1:
        return 0.6
    else:
        # e.g. high-content slide → low-capacity template: very poor fit
        return 0.15


def _score_image_relevance(
    deck_slide: SlideSpec,
    template: TemplateSlide,
) -> float:
    """Penalise matching content-image templates to slides that don't need images.

    Returns a penalty value (0.0 = no penalty, negative = penalty).
    """
    image_type = getattr(template, "image_type", "none")
    if image_type != "content":
        return 0.0  # No penalty for decorative or no images

    # Content images should only match slides that request images
    has_image_hints = bool(deck_slide.image_suggestions)
    is_image_slide = deck_slide.slide_type.value in ("image_full", "image_with_text")

    if has_image_hints or is_image_slide:
        return 0.0  # Good match — slide wants images

    # Template has content images but slide doesn't want them — penalise
    image_count = getattr(template, "image_count", 0)
    if image_count >= 4:
        return -0.15  # Heavy image template, strong penalty
    return -0.08  # Moderate penalty


def _score_structural_fit(
    deck_slide: SlideSpec,
    template: TemplateSlide,
) -> float:
    """Score structural compatibility (0.0 - 1.0).

    Compares content block count to content zones / placeholders, and checks
    whether the template has the right structure for the content.
    Uses content_zones (Phase 2) when available, falls back to placeholders.
    """
    content_blocks = len(deck_slide.content_blocks)

    # Use content zones if available (more precise), else placeholders
    if template.content_zones:
        zones = len(template.content_zones)
        # Check zone types match content block types
        deck_has_title = bool(deck_slide.title)
        deck_has_bullets = any(b.type == "bullets" for b in deck_slide.content_blocks)
        deck_has_data = any(b.type == "data_point" for b in deck_slide.content_blocks)

        tmpl_has_title = any(z.zone_type == "title" for z in template.content_zones)
        tmpl_has_body = any(z.zone_type in ("body", "bullet_area") for z in template.content_zones)
        tmpl_has_data = any(z.zone_type == "data_point" for z in template.content_zones)

        type_bonus = 0.0
        if deck_has_title and tmpl_has_title:
            type_bonus += 0.2
        if (deck_has_bullets or content_blocks > 0) and tmpl_has_body:
            type_bonus += 0.2
        if deck_has_data and tmpl_has_data:
            type_bonus += 0.3

        needed = content_blocks + (1 if deck_has_title else 0)
        diff = abs(zones - needed)
        if diff == 0:
            base = 1.0
        elif diff == 1:
            base = 0.7
        elif diff == 2:
            base = 0.4
        else:
            base = 0.1

        return min(1.0, base * 0.6 + type_bonus)
    else:
        placeholders = len(template.placeholders)
        needed = content_blocks + (1 if deck_slide.title else 0)
        if placeholders == 0:
            return 0.3

        diff = abs(placeholders - needed)
        if diff == 0:
            return 1.0
        elif diff == 1:
            return 0.7
        elif diff == 2:
            return 0.4
        else:
            return 0.1


def _score_visual_alignment(
    deck_slide: SlideSpec,
    template: TemplateSlide,
) -> float:
    """Score visual richness alignment (0.0 - 1.0).

    Uses visual_profile (Phase 2) when available for precise matching.
    Falls back to heuristic visual scoring.
    """
    # Determine desired visual profile from deck slide
    desired_profile = getattr(deck_slide, "visual_profile", None) or None

    # Infer desired profile from slide type and hints if not specified
    if not desired_profile:
        if deck_slide.slide_type in (SlideType.SECTION_HEADER, SlideType.TITLE, SlideType.CLOSING):
            desired_profile = "dark"
        elif deck_slide.slide_type == SlideType.QUOTE:
            desired_profile = "dark"
        elif deck_slide.slide_type in (SlideType.IMAGE_FULL, SlideType.IMAGE_WITH_TEXT):
            desired_profile = "branded_image"
        elif any(h in ("dark_background", "dark") for h in deck_slide.layout_hints):
            desired_profile = "dark"
        else:
            desired_profile = "light"

    # Score against template's visual profile (Phase 2 field)
    tmpl_profile = getattr(template, "visual_profile", "minimal")

    # Direct profile match scoring
    profile_score = 0.5  # neutral default
    if desired_profile == tmpl_profile:
        profile_score = 1.0
    elif desired_profile == "dark" and tmpl_profile in ("dark", "branded_image"):
        profile_score = 0.8
    elif desired_profile == "light" and tmpl_profile in ("light", "minimal"):
        profile_score = 0.8
    elif desired_profile == "branded_image" and tmpl_profile in ("branded_image", "dark"):
        profile_score = 0.7
    elif desired_profile == "dark" and tmpl_profile == "light":
        profile_score = 0.2  # Bad mismatch
    elif desired_profile == "light" and tmpl_profile == "dark":
        profile_score = 0.2  # Bad mismatch

    # Background quality bonus — templates with rich backgrounds are preferred
    bg_type = getattr(template, "background_type", "none")
    bg_bonus = 0.0
    if bg_type == "image":
        bg_bonus = 0.15
    elif bg_type == "gradient":
        bg_bonus = 0.1
    elif bg_type == "master_inherited" and template.has_background:
        bg_bonus = 0.08

    return min(1.0, profile_score + bg_bonus)


def _score_tag_overlap(
    deck_slide: SlideSpec,
    template: TemplateSlide,
) -> float:
    """Score tag overlap (0.0 - 1.0)."""
    # Build "desired tags" from deck slide hints and type
    desired = set()
    for hint in deck_slide.layout_hints:
        desired.add(hint.lower())

    # Add type-derived desires
    if deck_slide.slide_type in (SlideType.SECTION_HEADER, SlideType.TITLE):
        desired.update(["bold", "branded", "visual"])
    if deck_slide.slide_type == SlideType.QUOTE:
        desired.update(["minimal", "dark-bg"])
    if deck_slide.slide_type in (SlideType.CHART,):
        desired.update(["data-heavy"])

    if not desired or not template.tags:
        return 0.5  # Neutral when we can't compare

    tag_set = set(t.lower() for t in template.tags)
    overlap = desired & tag_set
    union = desired | tag_set
    if not union:
        return 0.5

    return len(overlap) / len(union)


# -----------------------------------------------------------------------
# Master scorer
# -----------------------------------------------------------------------

def score_template_for_slide(
    deck_slide: SlideSpec,
    template: TemplateSlide,
    template_index: int,
) -> dict:
    """Score a single template against a single deck slide.

    Returns a dict with component scores, total score, and reasoning.
    """
    # Hard filter: dimension-incompatible files
    # Normalize both sides to NFC to handle macOS NFD filenames
    template_file_norm = unicodedata.normalize("NFC", template.template_file.lower())
    for bad in _INCOMPATIBLE_FILES:
        if unicodedata.normalize("NFC", bad.lower()) in template_file_norm:
            return {
                "template_index": template_index,
                "total_score": 0.0,
                "rejected": True,
                "reason": f"Dimension-incompatible source: {bad}",
            }

    # Component scores
    type_score = _score_type_match(deck_slide.slide_type.value, template.slide_type.value)
    content_score = _score_content_similarity(deck_slide, template)
    struct_score = _score_structural_fit(deck_slide, template)
    visual_score = _score_visual_alignment(deck_slide, template)
    tag_score = _score_tag_overlap(deck_slide, template)
    capacity_score = _score_content_capacity_fit(deck_slide, template)
    image_penalty = _score_image_relevance(deck_slide, template)

    # Hard gate: if content score is negative, the template's text actively
    # conflicts with the deck slide's topic — reject this match
    if content_score < 0:
        return {
            "template_index": template_index,
            "total_score": 0.0,
            "rejected": True,
            "reason": f"Topic mismatch (content_score={content_score:.2f})",
        }

    # Weighted total — structural fit and visual alignment are key for clone-and-replace
    # Content capacity added to ensure templates can hold the content
    total = (
        0.12 * type_score
        + 0.08 * content_score
        + 0.30 * struct_score
        + 0.25 * visual_score
        + 0.10 * tag_score
        + 0.15 * capacity_score
    )

    # Image relevance penalty (applied as additive adjustment)
    total += image_penalty

    # Preferred source bonus
    if _PREFERRED_SOURCE.lower() in template.template_file.lower():
        total += 0.03

    # Visual/branded templates make better drop-ins than text-heavy ones
    text_weight = _template_text_weight(template)
    if text_weight <= 0.3 and template.has_images:
        total += 0.05  # Visual-first templates get a boost
    elif text_weight >= 0.6:
        total -= 0.05  # Text-heavy templates get penalized as drop-ins

    # Background bonus
    if template.has_background:
        total += 0.02

    return {
        "template_index": template_index,
        "total_score": round(total, 4),
        "type_score": round(type_score, 3),
        "content_score": round(content_score, 3),
        "struct_score": round(struct_score, 3),
        "visual_score": round(visual_score, 3),
        "tag_score": round(tag_score, 3),
        "capacity_score": round(capacity_score, 3),
        "image_penalty": round(image_penalty, 3),
        "rejected": False,
    }


# -----------------------------------------------------------------------
# Main matching logic
# -----------------------------------------------------------------------

def match_all(
    deck_schema: DeckSchema,
    registry: TemplateRegistry,
    threshold: float = 0.30,
    max_dropin_pct: float = 1.0,
) -> list[dict]:
    """Match all deck slides against the template registry.

    Clone-and-replace is the PRIMARY build mode. Every slide gets a template
    match unless no template scores above the minimum threshold. The threshold
    is deliberately low because even a loosely-matching template with branded
    visuals is better than a blank-canvas compose.

    Consistency groups ensure that slide types like section_header and
    title/closing all use the same template throughout the deck for visual
    coherence.

    Returns a list of match dicts compatible with the existing MatchResult schema.
    """
    total_slides = len(deck_schema.slides)
    max_dropins = int(total_slides * max_dropin_pct)

    # Score all templates for all slides
    slide_candidates: list[list[dict]] = []
    for slide_spec in deck_schema.slides:
        candidates = []
        for idx, tmpl in enumerate(registry.templates):
            result = score_template_for_slide(slide_spec, tmpl, idx)
            if not result["rejected"]:
                candidates.append(result)
        # Sort by score descending
        candidates.sort(key=lambda c: c["total_score"], reverse=True)
        slide_candidates.append(candidates)

    # ------------------------------------------------------------------
    # Pass 1: Determine the best template for each consistency group.
    # For each group, we pick the template that has the highest average
    # score across all slides in that group.
    # ------------------------------------------------------------------
    group_slide_indices: dict[str, list[int]] = {}  # group_name → [slide index]
    for i, slide_spec in enumerate(deck_schema.slides):
        group = _CONSISTENCY_GROUPS.get(slide_spec.slide_type.value)
        if group:
            group_slide_indices.setdefault(group, []).append(i)

    group_template: dict[str, int] = {}  # group_name → template_index
    for group, indices in group_slide_indices.items():
        if len(indices) < 2:
            continue  # No consistency needed for a single slide
        # Aggregate scores: for each template, compute average score across group slides
        tmpl_scores: dict[int, float] = {}
        for si in indices:
            for cand in slide_candidates[si]:
                tidx = cand["template_index"]
                tmpl_scores[tidx] = tmpl_scores.get(tidx, 0.0) + cand["total_score"]
        if tmpl_scores:
            # Average and pick the best
            best_tidx = max(tmpl_scores, key=lambda t: tmpl_scores[t] / len(indices))
            avg = tmpl_scores[best_tidx] / len(indices)
            if avg >= threshold:
                group_template[group] = best_tidx

    # ------------------------------------------------------------------
    # Pass 2: Greedy assignment with consistency group enforcement.
    # ------------------------------------------------------------------
    matches = []
    used_templates: set[int] = set()
    dropin_count = 0

    for i, slide_spec in enumerate(deck_schema.slides):
        candidates = slide_candidates[i]
        best = None

        # Check if this slide belongs to a consistency group with a locked template
        group = _CONSISTENCY_GROUPS.get(slide_spec.slide_type.value)
        locked_tidx = group_template.get(group) if group else None

        if locked_tidx is not None:
            # Use the group template — find its score for this slide
            for cand in candidates:
                if cand["template_index"] == locked_tidx:
                    # Accept unless the score is catastrophically low (>60% drop from best)
                    best_possible = candidates[0]["total_score"] if candidates else 0
                    if cand["total_score"] >= threshold or (
                        best_possible > 0 and cand["total_score"] >= 0.4 * best_possible
                    ):
                        best = cand
                    break
            # If the locked template wasn't even in the candidate list (rejected),
            # fall through to normal matching

        if not best:
            # Normal matching: prefer unused templates for variety
            for cand in candidates:
                if cand["template_index"] in used_templates:
                    continue
                if cand["total_score"] >= threshold and dropin_count < max_dropins:
                    best = cand
                    break

            # If no unused match found above threshold, allow reuse of templates
            if not best:
                for cand in candidates:
                    if cand["total_score"] >= threshold:
                        best = cand
                        break

        if best:
            tmpl = registry.templates[best["template_index"]]
            source_name = Path(tmpl.template_file).stem
            is_group_match = locked_tidx is not None and best["template_index"] == locked_tidx
            reasoning = (
                f"Template index {best['template_index']} "
                f"(slide_index {tmpl.slide_index}) — "
                f"{tmpl.description or tmpl.slide_type.value} "
                f"[type={best['type_score']}, content={best['content_score']}, "
                f"struct={best['struct_score']}, visual={best['visual_score']}, "
                f"capacity={best.get('capacity_score', 'n/a')}]. "
                f"Source: {source_name}"
            )
            if is_group_match:
                reasoning += f" (consistency group: {group})"

            matches.append({
                "slide_number": slide_spec.slide_number,
                "match_type": "use_as_is",
                "template_index": best["template_index"],
                "confidence": round(best["total_score"], 2),
                "reasoning": reasoning,
                "modifications": [],
            })
            used_templates.add(best["template_index"])
            dropin_count += 1
        else:
            matches.append({
                "slide_number": slide_spec.slide_number,
                "match_type": "build_from_base",
                "template_index": -1,
                "confidence": 0.9,
                "reasoning": (
                    f"No template scored above threshold ({threshold}) for "
                    f"'{slide_spec.title or slide_spec.intent[:60]}'. "
                    f"Build from base {slide_spec.slide_type.value} layout."
                ),
                "modifications": [],
            })

    return matches


def main():
    parser = argparse.ArgumentParser(
        description="Algorithmic template matching for slide deck generation"
    )
    parser.add_argument("deck_schema", type=Path, help="Path to deck_schema.json")
    parser.add_argument("template_registry", type=Path, help="Path to template_registry.json")
    parser.add_argument(
        "-o", "--output", type=Path,
        default=Path("workspace/template_matches.json"),
        help="Output path for matches JSON",
    )
    parser.add_argument(
        "--threshold", type=float, default=0.30,
        help="Minimum score to qualify for clone-and-replace (default: 0.30)",
    )
    parser.add_argument(
        "--max-dropin-pct", type=float, default=1.0,
        help="Maximum percentage of slides that can be cloned (default: 1.0 = all)",
    )
    args = parser.parse_args()

    # Validate inputs
    if not args.deck_schema.exists():
        print(f"Error: Deck schema not found: {args.deck_schema}", file=sys.stderr)
        sys.exit(1)
    if not args.template_registry.exists():
        print(f"Error: Template registry not found: {args.template_registry}", file=sys.stderr)
        sys.exit(1)

    # Load
    deck_data = json.loads(args.deck_schema.read_text(encoding="utf-8"))
    deck_schema = DeckSchema.model_validate(deck_data)
    registry = TemplateRegistry.load(args.template_registry)

    print(f"Deck: {deck_schema.title} ({len(deck_schema.slides)} slides)")
    print(f"Registry: {len(registry.templates)} templates from {len(registry.source_files)} files")
    print(f"Threshold: {args.threshold}, Max drop-in: {args.max_dropin_pct:.0%}")

    # Match
    matches = match_all(
        deck_schema, registry,
        threshold=args.threshold,
        max_dropin_pct=args.max_dropin_pct,
    )

    # Stats
    dropins = sum(1 for m in matches if m["match_type"] == "use_as_is")
    generated = sum(1 for m in matches if m["match_type"] == "build_from_base")
    print(f"\nResults: {dropins} drop-in, {generated} build-from-base")

    # Write
    result = {"matches": matches}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(f"Written to: {args.output}")

    # Show top matches
    for m in matches:
        status = "DROP-IN" if m["match_type"] == "use_as_is" else "BUILD  "
        print(f"  Slide {m['slide_number']:2d}: [{status}] conf={m['confidence']:.2f} — {m['reasoning'][:80]}")


if __name__ == "__main__":
    main()
