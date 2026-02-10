"""Parser for PDF files using pdfplumber."""

from pathlib import Path


def parse_pdf(path: Path) -> str:
    """Extract text from a PDF file with layout awareness.

    Uses pdfplumber which handles tables, columns, and complex layouts
    better than basic text extraction.
    """
    import pdfplumber

    sections: list[str] = []

    with pdfplumber.open(str(path)) as pdf:
        for i, page in enumerate(pdf.pages):
            page_text = page.extract_text() or ""

            # Extract tables separately for better structure
            tables = page.extract_tables()

            if page_text.strip():
                # Add page marker for multi-page documents
                if len(pdf.pages) > 1:
                    sections.append(f"\n--- Page {i + 1} ---\n")

                # Process text: detect likely headings (short lines in ALL CAPS or
                # lines that are significantly larger -- pdfplumber doesn't give
                # font size directly, so we heuristic on line length and caps)
                lines = page_text.split("\n")
                processed: list[str] = []
                for line in lines:
                    stripped = line.strip()
                    if not stripped:
                        processed.append("")
                        continue

                    # Heuristic: short ALL-CAPS lines are likely headings
                    if (
                        len(stripped) < 80
                        and stripped.isupper()
                        and len(stripped.split()) <= 10
                    ):
                        processed.append(f"# {stripped.title()}")
                    # Heuristic: lines starting with bullet characters
                    elif stripped[0] in ("•", "●", "○", "■", "▪", "►", "‣"):
                        processed.append(f"- {stripped[1:].strip()}")
                    elif stripped[:2] in ("- ", "* ", "· "):
                        processed.append(f"- {stripped[2:].strip()}")
                    else:
                        processed.append(stripped)

                sections.append("\n".join(processed))

            # Append tables as markdown-style tables
            for table in tables:
                if not table or not table[0]:
                    continue
                table_lines: list[str] = []
                for row_idx, row in enumerate(table):
                    cells = [str(cell or "").strip() for cell in row]
                    table_lines.append("| " + " | ".join(cells) + " |")
                    if row_idx == 0:
                        table_lines.append("|" + "|".join(["---"] * len(cells)) + "|")
                sections.append("\n" + "\n".join(table_lines) + "\n")

    result = "\n".join(sections).strip()
    if not result:
        raise ValueError(f"No text content could be extracted from {path}")
    return result
