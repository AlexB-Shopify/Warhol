#!/usr/bin/env python3
"""Spread over-dense slides across multiple output slides.

Takes a density report and deck schema, splits slides flagged as
"spread_candidate" into multiple slides that each fit within recommended
content limits. Re-numbers all slides afterward.

Usage:
    python scripts/spread_content.py \
        workspace/deck_schema.json \
        workspace/density_report.json \
        -o workspace/deck_schema.json
"""

import argparse
import copy
import json
import math
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.schemas.slide_schema import DeckSchema

# ---------------------------------------------------------------------------
# Splitting strategies per content pattern
# ---------------------------------------------------------------------------


def _split_by_bullets(slide_dict: dict, max_per_slide: int = 4) -> list[dict]:
    """Split a slide with too many bullets into multiple slides.

    Distributes bullet items across multiple slides, each getting up to
    max_per_slide bullets. The title is shared (first slide gets original,
    subsequent slides get "Title (cont.)").
    """
    content_blocks = slide_dict.get("content_blocks", [])
    title = slide_dict.get("title", "")

    # Separate bullet blocks from non-bullet blocks
    bullet_items: list[str] = []
    non_bullet_blocks: list[dict] = []
    for block in content_blocks:
        if block["type"] == "bullets":
            for line in block["content"].split("\n"):
                line = line.strip()
                if line:
                    bullet_items.append(line)
        else:
            non_bullet_blocks.append(block)

    if len(bullet_items) <= max_per_slide:
        return [slide_dict]  # No split needed

    # Split bullets into chunks
    chunks = [
        bullet_items[i:i + max_per_slide]
        for i in range(0, len(bullet_items), max_per_slide)
    ]

    slides = []
    for i, chunk in enumerate(chunks):
        new_slide = copy.deepcopy(slide_dict)
        suffix = f" (cont.)" if i > 0 else ""
        new_slide["title"] = title + suffix

        # Build content blocks: non-bullet blocks only on first slide, bullets on all
        new_blocks = []
        if i == 0:
            new_blocks.extend(copy.deepcopy(non_bullet_blocks))
        new_blocks.append({
            "type": "bullets",
            "content": "\n".join(chunk),
            "emphasis": "normal",
        })
        new_slide["content_blocks"] = new_blocks

        # Speaker notes: full notes only on first slide, abbreviated on rest
        if i > 0:
            new_slide["speaker_notes"] = f"Continuation of: {title}"

        slides.append(new_slide)

    return slides


def _split_by_content_blocks(slide_dict: dict, max_blocks: int = 3) -> list[dict]:
    """Split a slide with too many content blocks across multiple slides.

    Groups content blocks into chunks of max_blocks size.
    """
    content_blocks = slide_dict.get("content_blocks", [])
    title = slide_dict.get("title", "")

    if len(content_blocks) <= max_blocks:
        return [slide_dict]

    chunks = [
        content_blocks[i:i + max_blocks]
        for i in range(0, len(content_blocks), max_blocks)
    ]

    slides = []
    for i, chunk in enumerate(chunks):
        new_slide = copy.deepcopy(slide_dict)
        suffix = f" (cont.)" if i > 0 else ""
        new_slide["title"] = title + suffix
        new_slide["content_blocks"] = copy.deepcopy(chunk)

        if i > 0:
            new_slide["speaker_notes"] = f"Continuation of: {title}"

        slides.append(new_slide)

    return slides


def _split_by_char_count(slide_dict: dict, max_chars: int = 400) -> list[dict]:
    """Split a slide based on total character count.

    Distributes content blocks across slides such that each stays under
    max_chars.
    """
    content_blocks = slide_dict.get("content_blocks", [])
    title = slide_dict.get("title", "")
    title_chars = len(title or "")

    if not content_blocks:
        return [slide_dict]

    slides = []
    current_blocks = []
    current_chars = title_chars

    for block in content_blocks:
        block_chars = len(block.get("content", ""))
        if current_chars + block_chars > max_chars and current_blocks:
            # Start a new slide
            new_slide = copy.deepcopy(slide_dict)
            suffix = f" (cont.)" if slides else ""
            new_slide["title"] = title + suffix
            new_slide["content_blocks"] = current_blocks
            if slides:
                new_slide["speaker_notes"] = f"Continuation of: {title}"
            slides.append(new_slide)
            current_blocks = [copy.deepcopy(block)]
            current_chars = title_chars + block_chars
        else:
            current_blocks.append(copy.deepcopy(block))
            current_chars += block_chars

    # Flush remaining blocks
    if current_blocks:
        new_slide = copy.deepcopy(slide_dict)
        suffix = f" (cont.)" if slides else ""
        new_slide["title"] = title + suffix
        new_slide["content_blocks"] = current_blocks
        if slides:
            new_slide["speaker_notes"] = f"Continuation of: {title}"
        slides.append(new_slide)

    return slides if slides else [slide_dict]


# ---------------------------------------------------------------------------
# Main spreading logic
# ---------------------------------------------------------------------------

def spread_slides(
    deck_data: dict,
    density_reports: list[dict],
) -> dict:
    """Process the deck schema, spreading over-dense slides.

    For each slide flagged as "spread_candidate" in the density report,
    splits it into multiple slides using the appropriate strategy.

    Returns the modified deck data dict (not yet validated).
    """
    slides = deck_data.get("slides", [])
    spread_lookup = {}
    for report in density_reports:
        if report["status"] == "spread_candidate":
            spread_lookup[report["slide_number"]] = report

    new_slides = []
    splits_made = 0

    for slide_dict in slides:
        slide_num = slide_dict.get("slide_number", 0)
        report = spread_lookup.get(slide_num)

        if report is None:
            new_slides.append(slide_dict)
            continue

        # Choose splitting strategy based on content pattern
        slide_type = slide_dict.get("slide_type", "content")
        bullet_count = report.get("bullet_count", 0)
        block_count = report.get("block_count", 0)
        max_chars = report.get("effective_max_chars", 400)

        if bullet_count > 4:
            # Bullet-heavy: split by bullets
            split_slides = _split_by_bullets(slide_dict, max_per_slide=4)
        elif block_count > 3:
            # Many content blocks: split by blocks
            split_slides = _split_by_content_blocks(slide_dict, max_blocks=3)
        else:
            # Character-heavy: split by char count
            split_slides = _split_by_char_count(slide_dict, max_chars=max_chars)

        if len(split_slides) > 1:
            splits_made += 1
            print(
                f"  Slide {slide_num}: split into {len(split_slides)} slides "
                f"({report['recommendation'][:60]})"
            )

        new_slides.extend(split_slides)

    # Re-number all slides sequentially
    for i, slide_dict in enumerate(new_slides, 1):
        slide_dict["slide_number"] = i

    deck_data["slides"] = new_slides

    return deck_data, splits_made


def main():
    parser = argparse.ArgumentParser(
        description="Spread over-dense slides across multiple output slides"
    )
    parser.add_argument("deck_schema", type=Path, help="Path to deck_schema.json")
    parser.add_argument("density_report", type=Path, help="Path to density_report.json")
    parser.add_argument(
        "-o", "--output", type=Path,
        default=Path("workspace/deck_schema.json"),
        help="Output path for updated deck schema JSON",
    )
    args = parser.parse_args()

    # Validate inputs
    for f in [args.deck_schema, args.density_report]:
        if not f.exists():
            print(f"Error: File not found: {f}", file=sys.stderr)
            sys.exit(1)

    # Load
    deck_data = json.loads(args.deck_schema.read_text(encoding="utf-8"))
    density_data = json.loads(args.density_report.read_text(encoding="utf-8"))
    reports = density_data.get("reports", [])

    original_count = len(deck_data.get("slides", []))
    spread_count = sum(1 for r in reports if r["status"] == "spread_candidate")

    print(f"Deck: {deck_data.get('title', 'Untitled')} ({original_count} slides)")
    print(f"Spread candidates: {spread_count}")

    if spread_count == 0:
        print("No slides need spreading. Deck schema unchanged.")
        return

    # Spread
    updated_data, splits = spread_slides(deck_data, reports)
    new_count = len(updated_data.get("slides", []))

    # Validate the result
    try:
        DeckSchema.model_validate(updated_data)
        print("Validation: OK")
    except Exception as e:
        print(f"Validation warning: {e}")
        print("Writing anyway — manual review recommended.")

    # Write
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(updated_data, indent=2), encoding="utf-8")
    print(f"\nWritten to: {args.output}")
    print(f"Slides: {original_count} → {new_count} ({splits} slides split)")
    print(f"\nIMPORTANT: Re-run template matching after spreading:")
    print(f"  python scripts/match_templates.py {args.output} template_registry.json")


if __name__ == "__main__":
    main()
