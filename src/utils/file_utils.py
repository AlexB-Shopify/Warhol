"""File I/O and path utilities."""

import json
import logging
import tempfile
from pathlib import Path
from typing import Any

import yaml

logger = logging.getLogger(__name__)


def ensure_directory(path: str | Path) -> Path:
    """Ensure a directory exists, creating it if necessary."""
    path = Path(path)
    path.mkdir(parents=True, exist_ok=True)
    return path


def load_yaml(path: str | Path) -> dict[str, Any]:
    """Load a YAML file and return its contents as a dict."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"YAML file not found: {path}")
    with open(path) as f:
        return yaml.safe_load(f) or {}


def save_yaml(data: dict[str, Any], path: str | Path) -> None:
    """Save a dict to a YAML file."""
    path = Path(path)
    ensure_directory(path.parent)
    with open(path, "w") as f:
        yaml.dump(data, f, default_flow_style=False, sort_keys=False)


def load_json(path: str | Path) -> dict[str, Any]:
    """Load a JSON file and return its contents."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"JSON file not found: {path}")
    with open(path) as f:
        return json.load(f)


def save_json(data: Any, path: str | Path, indent: int = 2) -> None:
    """Save data to a JSON file."""
    path = Path(path)
    ensure_directory(path.parent)
    with open(path, "w") as f:
        json.dump(data, f, indent=indent, default=str)


def get_temp_directory(prefix: str = "slide_builder_") -> Path:
    """Create and return a temporary directory."""
    return Path(tempfile.mkdtemp(prefix=prefix))


def find_pptx_files(directory: str | Path) -> list[Path]:
    """Recursively find all .pptx files in a directory."""
    directory = Path(directory)
    if not directory.is_dir():
        raise NotADirectoryError(f"Not a directory: {directory}")
    files = sorted(directory.rglob("*.pptx"))
    # Exclude temp/hidden files
    files = [f for f in files if not f.name.startswith(("~", "."))]
    return files


def read_input_file(path: str | Path) -> str:
    """Read an input file using the appropriate parser.

    Convenience wrapper around src.parsers.parse().
    """
    from src.parsers import parse

    return parse(path)
