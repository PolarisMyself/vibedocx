"""Configuration system with 3-tier priority: CLI > project > skill default.

Shared constants (PAPER_SIZES, ALIGN_MAP, etc.) are also defined here to
eliminate duplication across modules.
"""

import json
import os
from pathlib import Path

from docx.enum.text import WD_ALIGN_PARAGRAPH

# ── Shared constants ──

PAPER_SIZES = {
    "A4": (21.0, 29.7),
    "A3": (29.7, 42.0),
    "Letter": (21.59, 27.94),
}

ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

ALIGN_REVERSE = {
    WD_ALIGN_PARAGRAPH.LEFT: "left",
    WD_ALIGN_PARAGRAPH.CENTER: "center",
    WD_ALIGN_PARAGRAPH.RIGHT: "right",
    WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
}


# ── Config loading ──

def _deep_merge(base, override):
    """Recursively merge override dict into base dict."""
    for key, value in override.items():
        if key in base and isinstance(base[key], dict) and isinstance(value, dict):
            _deep_merge(base[key], value)
        else:
            base[key] = value
    return base


def _load_json(path):
    """Load a JSON file, returning {} on any failure."""
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError, PermissionError):
        return {}


def load_config(project_dir=None):
    """Load merged style config.

    Priority (low to high):
      1. Skill default: <package>/config/style.json
      2. Project: <project_dir>/.vibedocx/style.json
      3. (CLI overrides handled separately by argparse)

    Args:
        project_dir: Path to project root. If None, uses cwd.

    Returns:
        Merged config dict.
    """
    import helper
    skill_dir = Path(helper.__file__).parent
    default_path = skill_dir / "config" / "style.json"

    config = _load_json(str(default_path))

    proj_dir = project_dir or os.getcwd()
    project_path = Path(proj_dir) / ".vibedocx" / "style.json"
    project_config = _load_json(str(project_path))
    if project_config:
        _deep_merge(config, project_config)

    return config


def get_font(config, key):
    """Get a font value from config. Returns None if not found."""
    return config.get("fonts", {}).get(key)


def get_size(config, key):
    """Get a size value from config. Returns None if not found."""
    return config.get("sizes", {}).get(key)
