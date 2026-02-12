#!/usr/bin/env python3
"""Validate a JSON file against a named Pydantic schema.

Usage:
    python scripts/validate_schema.py <json_file> <schema_name>

Schema names:
    ContentMaturity   -- Content maturity assessment and pipeline routing
    ContentInventory  -- Extracted content from input documents
    DeckSchema        -- Complete slide-by-slide deck specification
    TemplateRegistry  -- Analyzed template metadata
    DesignSystem      -- Font and color configuration
    MatchResult       -- Template-to-slide matches
    HtmlDeck          -- HTML intermediate deck specification
"""

import argparse
import json
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pydantic import BaseModel, Field

from src.schemas.slide_schema import ContentInventory, ContentMaturity, DeckSchema
from src.schemas.template_schema import TemplateRegistry
from src.schemas.design_system import DesignSystem
from src.schemas.html_schema import HtmlDeck


# MatchResult is defined here (was previously in template_matcher agent)
class SlideMatch(BaseModel):
    """A single slide-to-template match."""
    slide_number: int
    match_type: str = Field(
        default="build_from_base",
        description="'use_as_is' to clone template slide exactly, 'build_from_base' to generate from layout",
    )
    template_index: int = Field(default=-1, description="Index into the template registry list (-1 = generate)")
    confidence: float = Field(ge=0.0, le=1.0, description="Match confidence 0-1")
    reasoning: str = Field(description="Why this template/approach was chosen")
    modifications: list[str] = Field(
        default_factory=list,
        description="Any modifications needed (only for build_from_base)",
    )


class MatchResult(BaseModel):
    """Complete matching result for all slides."""
    matches: list[SlideMatch]


SCHEMA_MAP: dict[str, type[BaseModel]] = {
    "ContentInventory": ContentInventory,
    "ContentMaturity": ContentMaturity,
    "DeckSchema": DeckSchema,
    "TemplateRegistry": TemplateRegistry,
    "DesignSystem": DesignSystem,
    "MatchResult": MatchResult,
    "HtmlDeck": HtmlDeck,
}


def main():
    parser = argparse.ArgumentParser(description="Validate JSON against a Pydantic schema")
    parser.add_argument("json_file", type=Path, help="Path to JSON file to validate")
    parser.add_argument("schema_name", choices=list(SCHEMA_MAP.keys()),
                        help="Name of the Pydantic schema to validate against")
    args = parser.parse_args()

    if not args.json_file.exists():
        print(f"Error: JSON file not found: {args.json_file}", file=sys.stderr)
        sys.exit(1)

    schema_cls = SCHEMA_MAP[args.schema_name]

    try:
        raw = args.json_file.read_text(encoding="utf-8")
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON in {args.json_file}: {e}", file=sys.stderr)
        sys.exit(1)

    try:
        instance = schema_cls.model_validate(data)
        print(f"Validation PASSED: {args.json_file} conforms to {args.schema_name}")

        # Print summary info based on schema type
        if args.schema_name == "ContentMaturity":
            print(f"  Level: {instance.maturity_level} ({instance.maturity_label})")
            print(f"  Stages: {' → '.join(instance.pipeline_stages)}")
            print(f"  Word count: {instance.word_count}")
            print(f"  Gaps: {len(instance.content_gaps)}")
        elif args.schema_name == "ContentInventory":
            print(f"  Main topic: {instance.main_topic}")
            print(f"  Sections: {len(instance.sections)}")
            print(f"  Themes: {len(instance.themes)}")
            print(f"  Data points: {len(instance.key_data_points)}")
        elif args.schema_name == "DeckSchema":
            print(f"  Title: {instance.title}")
            print(f"  Slides: {len(instance.slides)}")
            print(f"  Target audience: {instance.target_audience}")
        elif args.schema_name == "TemplateRegistry":
            print(f"  Templates: {len(instance.templates)}")
            print(f"  Source files: {len(instance.source_files)}")
        elif args.schema_name == "MatchResult":
            print(f"  Matches: {len(instance.matches)}")
        elif args.schema_name == "DesignSystem":
            print(f"  Name: {instance.name}")
            print(f"  Title font: {instance.fonts.title_font}")
            print(f"  Primary color: {instance.colors.primary}")
        elif args.schema_name == "HtmlDeck":
            print(f"  Title: {instance.title}")
            print(f"  Slides: {len(instance.slides)}")
            total_elements = sum(len(s.elements) for s in instance.slides)
            print(f"  Total elements: {total_elements}")

    except Exception as e:
        print(f"Validation FAILED: {args.json_file} does not conform to {args.schema_name}", file=sys.stderr)
        print(f"  Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
