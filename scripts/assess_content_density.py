#!/usr/bin/env python3
"""Assess content density — flags slides where content exceeds template capacity.

Compares each slide's content volume against the matched template's capacity
and content zones. Produces a density report that the design optimization step
can use to split over-dense slides.

Usage:
    python scripts/assess_content_density.py \
        workspace/deck_schema.json \
        workspace/template_matches.json \
        template_registry.json \
        -o workspace/density_report.json
"""

import argparse
import json
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.schemas.slide_schema import DeckSchema, SlideSpec
from src.schemas.template_schema import TemplateRegistry, TemplateSlide

# ---------------------------------------------------------------------------
# Content volume estimation
# ---------------------------------------------------------------------------

# Maximum recommended content per slide type
_MAX_CONTENT_BLOCKS: dict[str, int] = {
    "title": 2,
    "closing": 2,
    "section_header": 2,
    "content": 4,
    "bullet_list": 6,  # 6x6 rule: max 6 bullets
    "two_column": 6,   # 3 per column
    "comparison": 6,
    "quote": 2,
    "chart": 2,
    "image_full": 1,
    "image_with_text": 3,
    "timeline": 5,
    "team": 4,
    "data_point": 2,
}

_MAX_CHARS_PER_SLIDE: dict[str, int] = {
    "title": 150,
    "closing": 150,
    "section_header": 100,
    "content": 500,
    "bullet_list": 400,
    "two_column": 600,
    "comparison": 600,
    "quote": 200,
    "chart": 200,
    "image_full": 80,
    "image_with_text": 300,
    "timeline": 400,
    "team": 300,
    "data_point": 150,
}


def _count_chars(spec: SlideSpec) -> int:
    """Count total characters in a slide's content."""
    total = len(spec.title or "")
    total += len(spec.subtitle or "")
    for block in spec.content_blocks:
        total += len(block.content)
    return total


def _count_bullets(spec: SlideSpec) -> int:
    """Count total bullet items in a slide's content blocks."""
    count = 0
    for block in spec.content_blocks:
        if block.type == "bullets":
            count += len([l for l in block.content.split("\n") if l.strip()])
    return count


def _template_total_max_chars(template: TemplateSlide) -> int:
    """Sum max_chars across all content zones in a template."""
    if not template.content_zones:
        # Estimate from content_capacity
        cap_map = {"low": 100, "medium": 300, "high": 600}
        return cap_map.get(template.content_capacity, 300)

    return sum(z.max_chars for z in template.content_zones)


# ---------------------------------------------------------------------------
# Density assessment
# ---------------------------------------------------------------------------

def assess_density(
    deck_schema: DeckSchema,
    matches: list[dict],
    registry: TemplateRegistry,
) -> list[dict]:
    """Assess content density for each slide.

    Returns a list of density reports, one per slide, with:
    - char_count: total characters
    - block_count: number of content blocks
    - bullet_count: number of bullet items
    - template_max_chars: template's total zone capacity (None if no template)
    - density_ratio: char_count / template_max_chars (None if N/A)
    - status: "ok", "dense", "spread_candidate"
    - recommendation: human-readable suggestion
    """
    match_lookup = {}
    for m in matches:
        match_lookup[m["slide_number"]] = m

    reports = []
    for spec in deck_schema.slides:
        slide_type = spec.slide_type.value
        char_count = _count_chars(spec)
        block_count = len(spec.content_blocks)
        bullet_count = _count_bullets(spec)

        max_blocks = _MAX_CONTENT_BLOCKS.get(slide_type, 4)
        max_chars = _MAX_CHARS_PER_SLIDE.get(slide_type, 400)

        # Template-specific capacity
        match = match_lookup.get(spec.slide_number)
        template_max_chars = None
        template_zone_count = None
        if match and match.get("match_type") == "use_as_is" and match.get("template_index", -1) >= 0:
            tidx = match["template_index"]
            if tidx < len(registry.templates):
                tmpl = registry.templates[tidx]
                template_max_chars = _template_total_max_chars(tmpl)
                template_zone_count = len(tmpl.content_zones) if tmpl.content_zones else None

        # Use the more constrained of slide-type max and template max
        effective_max_chars = max_chars
        if template_max_chars is not None:
            effective_max_chars = min(max_chars, template_max_chars)

        density_ratio = char_count / max(1, effective_max_chars)

        # Determine status
        status = "ok"
        recommendation = None

        if density_ratio > 1.5 or block_count > max_blocks + 2:
            status = "spread_candidate"
            # Calculate how many slides this should be split into
            suggested_slides = max(2, int(density_ratio + 0.5))
            recommendation = (
                f"Content too dense for a single {slide_type} slide "
                f"({char_count} chars vs {effective_max_chars} capacity, "
                f"{block_count} blocks vs {max_blocks} recommended). "
                f"Consider splitting into {suggested_slides} slides."
            )
        elif density_ratio > 1.0 or block_count > max_blocks:
            status = "dense"
            recommendation = (
                f"Content is dense for a {slide_type} slide "
                f"({char_count} chars vs {effective_max_chars} capacity). "
                f"Consider trimming bullets or moving detail to speaker notes."
            )

        reports.append({
            "slide_number": spec.slide_number,
            "slide_type": slide_type,
            "char_count": char_count,
            "block_count": block_count,
            "bullet_count": bullet_count,
            "max_blocks_recommended": max_blocks,
            "max_chars_recommended": max_chars,
            "template_max_chars": template_max_chars,
            "template_zone_count": template_zone_count,
            "effective_max_chars": effective_max_chars,
            "density_ratio": round(density_ratio, 2),
            "status": status,
            "recommendation": recommendation,
        })

    return reports


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Assess content density against template capacity"
    )
    parser.add_argument("deck_schema", type=Path, help="Path to deck_schema.json")
    parser.add_argument("template_matches", type=Path, help="Path to template_matches.json")
    parser.add_argument("template_registry", type=Path, help="Path to template_registry.json")
    parser.add_argument(
        "-o", "--output", type=Path,
        default=Path("workspace/density_report.json"),
        help="Output path for density report JSON",
    )
    args = parser.parse_args()

    # Validate inputs
    for f in [args.deck_schema, args.template_matches, args.template_registry]:
        if not f.exists():
            print(f"Error: File not found: {f}", file=sys.stderr)
            sys.exit(1)

    # Load
    deck_data = json.loads(args.deck_schema.read_text(encoding="utf-8"))
    deck_schema = DeckSchema.model_validate(deck_data)

    matches_data = json.loads(args.template_matches.read_text(encoding="utf-8"))
    matches = matches_data.get("matches", [])

    registry = TemplateRegistry.load(args.template_registry)

    print(f"Assessing density for {len(deck_schema.slides)} slides...")

    # Assess
    reports = assess_density(deck_schema, matches, registry)

    # Stats
    ok_count = sum(1 for r in reports if r["status"] == "ok")
    dense_count = sum(1 for r in reports if r["status"] == "dense")
    spread_count = sum(1 for r in reports if r["status"] == "spread_candidate")

    result = {
        "total_slides": len(reports),
        "ok": ok_count,
        "dense": dense_count,
        "spread_candidates": spread_count,
        "reports": reports,
    }

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(f"Written to: {args.output}")
    print(f"Results: {ok_count} ok, {dense_count} dense, {spread_count} spread candidates")

    # Show details for non-ok slides
    for r in reports:
        if r["status"] != "ok":
            print(
                f"  Slide {r['slide_number']:2d} [{r['status'].upper():16s}] "
                f"density={r['density_ratio']:.1f}x — {r['recommendation'][:80]}"
            )


if __name__ == "__main__":
    main()
