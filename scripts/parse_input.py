#!/usr/bin/env python3
"""Parse an input document and write extracted text to a file.

Supports: PDF, DOCX, PPTX, TXT, MD

Usage:
    python scripts/parse_input.py <input_file> [-o workspace/parsed_content.txt]
"""

import argparse
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.parsers import parse


def main():
    parser = argparse.ArgumentParser(description="Parse input document to text")
    parser.add_argument("input_file", type=Path, help="Input document path (.txt, .md, .pdf, .docx, .pptx)")
    parser.add_argument("-o", "--output", type=Path, default=Path("workspace/parsed_content.txt"),
                        help="Output text file path (default: workspace/parsed_content.txt)")
    args = parser.parse_args()

    if not args.input_file.exists():
        print(f"Error: Input file not found: {args.input_file}", file=sys.stderr)
        sys.exit(1)

    # Ensure output directory exists
    args.output.parent.mkdir(parents=True, exist_ok=True)

    print(f"Parsing: {args.input_file}")
    text = parse(args.input_file)

    args.output.write_text(text, encoding="utf-8")
    print(f"Parsed content written to: {args.output}")
    print(f"Content length: {len(text)} characters, {len(text.splitlines())} lines")


if __name__ == "__main__":
    main()
