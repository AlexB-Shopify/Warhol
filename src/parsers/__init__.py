from pathlib import Path

from .text_parser import parse_text
from .pdf_parser import parse_pdf
from .docx_parser import parse_docx
from .pptx_parser import parse_pptx

_PARSER_MAP = {
    ".txt": parse_text,
    ".md": parse_text,
    ".pdf": parse_pdf,
    ".docx": parse_docx,
    ".pptx": parse_pptx,
}


def parse(path: str | Path) -> str:
    """Parse a document file and return structured text content.

    Dispatches to the appropriate parser based on file extension.
    Supported formats: .txt, .md, .pdf, .docx, .pptx
    """
    path = Path(path)
    ext = path.suffix.lower()
    parser = _PARSER_MAP.get(ext)
    if parser is None:
        supported = ", ".join(sorted(_PARSER_MAP.keys()))
        raise ValueError(f"Unsupported file format '{ext}'. Supported: {supported}")
    return parser(path)


__all__ = ["parse", "parse_text", "parse_pdf", "parse_docx", "parse_pptx"]
