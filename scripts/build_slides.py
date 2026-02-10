#!/usr/bin/env python3
"""Build a PowerPoint presentation from a deck schema and design system.

Usage:
    # Generated slides only (no template bank):
    python scripts/build_slides.py workspace/deck_schema.json -o output.pptx --design-system design_systems/corporate.yaml

    # With template matching (generated + drop-in):
    python scripts/build_slides.py workspace/deck_schema.json -o output.pptx --design-system design_systems/corporate.yaml \
        --matches workspace/template_matches.json --template-registry template_registry.json

    # Custom base template:
    python scripts/build_slides.py workspace/deck_schema.json -o output.pptx --design-system design_systems/corporate.yaml \
        --base-template templates/base/shopify_base.pptx
"""

import argparse
import json
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.schemas.design_system import DesignSystem
from src.schemas.slide_schema import DeckSchema
from src.schemas.template_schema import TemplateRegistry
from src.agents.slide_builder import SlideBuilderAgent
from src.pptx_engine.slide_operations import clear_clone_caches


def main():
    parser = argparse.ArgumentParser(description="Build PPTX from deck schema")
    parser.add_argument("deck_schema", type=Path, help="Path to deck_schema.json")
    parser.add_argument("-o", "--output", type=Path, default=Path("output.pptx"),
                        help="Output PPTX path (default: output.pptx)")
    parser.add_argument("--design-system", type=Path, default=Path("design_systems/default.yaml"),
                        help="Design system YAML path")
    parser.add_argument("--matches", type=Path, default=None,
                        help="Optional template_matches.json path")
    parser.add_argument("--template-registry", type=Path, default=None,
                        help="Optional template_registry.json path (needed for drop-in slides)")
    parser.add_argument("--base-template", type=Path, default=None,
                        help="Base template PPTX (default: templates/base/shopify_base.pptx)")
    args = parser.parse_args()

    # Validate inputs
    if not args.deck_schema.exists():
        print(f"Error: Deck schema not found: {args.deck_schema}", file=sys.stderr)
        sys.exit(1)

    if not args.design_system.exists():
        print(f"Error: Design system not found: {args.design_system}", file=sys.stderr)
        sys.exit(1)

    # Load deck schema
    deck_data = json.loads(args.deck_schema.read_text(encoding="utf-8"))
    deck_schema = DeckSchema.model_validate(deck_data)

    # Load design system
    design_system = DesignSystem.from_yaml(args.design_system)

    # Load matches if provided
    matches = None
    if args.matches and args.matches.exists():
        matches_data = json.loads(args.matches.read_text(encoding="utf-8"))
        matches = matches_data.get("matches", [])

    # Load template registry if provided
    template_registry = None
    if args.template_registry and args.template_registry.exists():
        template_registry = TemplateRegistry.load(args.template_registry)

    # Ensure output directory exists
    args.output.parent.mkdir(parents=True, exist_ok=True)

    # Build
    builder = SlideBuilderAgent()
    result_path = builder.build(
        deck_schema=deck_schema,
        design_system=design_system,
        output_path=args.output,
        matches=matches,
        template_registry=template_registry,
        base_template=args.base_template,
    )

    # Count match types
    generated = len(deck_schema.slides)
    drop_ins = 0
    if matches:
        drop_ins = sum(1 for m in matches if m.get("match_type") == "use_as_is")
        generated -= drop_ins

    # Free cached source presentations and imported layouts
    clear_clone_caches()

    print(f"Presentation generated: {result_path}")
    print(f"Slides: {len(deck_schema.slides)} ({generated} generated, {drop_ins} drop-in)")
    print(f"Design system: {design_system.name}")
    print(f"Base template: {args.base_template or 'templates/base/Shopify - Example Technical Workshop Slide Bank.pptx'}")


if __name__ == "__main__":
    main()
