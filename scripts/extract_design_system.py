#!/usr/bin/env python3
"""Extract a design system (fonts, colors, sizes) from PPTX template files.

Usage:
    # From a single file:
    python scripts/extract_design_system.py templates/corporate/brand.pptx -o design_systems/extracted.yaml

    # From a directory:
    python scripts/extract_design_system.py templates/ -o design_systems/extracted.yaml
"""

import argparse
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.pptx_engine.design_system_extractor import (
    extract_design_system,
    extract_design_system_from_file,
)


def main():
    parser = argparse.ArgumentParser(description="Extract design system from PPTX files")
    parser.add_argument("source", type=Path,
                        help="Source .pptx file or directory of .pptx files")
    parser.add_argument("-o", "--output", type=Path,
                        default=Path("design_systems/extracted.yaml"),
                        help="Output YAML path")
    parser.add_argument("--name", type=str, default="Extracted",
                        help="Name for the design system")
    args = parser.parse_args()

    if not args.source.exists():
        print(f"Error: Not found: {args.source}", file=sys.stderr)
        sys.exit(1)

    is_single_file = args.source.is_file()

    if is_single_file:
        if args.source.suffix.lower() != ".pptx":
            print(f"Error: Expected a .pptx file, got: {args.source.suffix}", file=sys.stderr)
            sys.exit(1)
        print(f"Extracting design system from: {args.source.name}")
        ds = extract_design_system_from_file(args.source, name=args.name)
    else:
        from src.utils.file_utils import find_pptx_files
        pptx_files = find_pptx_files(args.source)
        if not pptx_files:
            print(f"No .pptx files found in {args.source}", file=sys.stderr)
            sys.exit(0)
        print(f"Analyzing {len(pptx_files)} .pptx files for design patterns...")
        ds = extract_design_system(args.source, name=args.name)

    # Ensure output directory exists
    args.output.parent.mkdir(parents=True, exist_ok=True)
    ds.to_yaml(args.output)

    print(f"\nDesign system saved to: {args.output}")
    print(f"  Name: {ds.name}")
    print(f"  Title font: {ds.fonts.title_font} ({ds.fonts.title_size}pt)")
    print(f"  Body font: {ds.fonts.body_font} ({ds.fonts.body_size}pt)")
    print(f"  Primary color: {ds.colors.primary}")
    print(f"  Secondary color: {ds.colors.secondary}")
    print(f"  Accent color: {ds.colors.accent}")


if __name__ == "__main__":
    main()
