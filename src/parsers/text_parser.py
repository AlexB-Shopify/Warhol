"""Parser for plain text and markdown files."""

from pathlib import Path


def parse_text(path: Path) -> str:
    """Read a .txt or .md file and return its content.

    Markdown files are returned as-is since they already have structural markers.
    Plain text is returned with minimal normalization.
    """
    text = path.read_text(encoding="utf-8")

    # Normalize excessive blank lines
    lines = text.splitlines()
    normalized: list[str] = []
    blank_count = 0
    for line in lines:
        if line.strip() == "":
            blank_count += 1
            if blank_count <= 2:
                normalized.append("")
        else:
            blank_count = 0
            normalized.append(line)

    return "\n".join(normalized).strip()
