"""Unit parsing — converts user-facing values to python-docx Length objects.

Supports Chinese字号 (三号, 小四, etc.), point values (12, 14pt, 12pt),
and character-based indent.
"""

from docx.shared import Pt


# ── Chinese font size → points ──

SIZES_CN = {
    "初号": 42, "小初": 36,
    "一号": 26, "小一": 24,
    "二号": 22, "小二": 18,
    "三号": 16, "小三": 15,
    "四号": 14, "小四": 12,
    "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5,
    "七号": 5.5, "八号": 5,
}


def parse_size(value):
    """Parse a font size specification and return a Pt Length.

    Accepted formats:
      - int/float:       12 → Pt(12)
      - Chinese字号:     "三号" → Pt(16), "小四" → Pt(12)
      - string with pt:  "14pt" → Pt(14), "12pt" → Pt(12)
      - plain string:    "12" → Pt(12)
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return Pt(value)
    if isinstance(value, str):
        v = value.strip()
        if v in SIZES_CN:
            return Pt(SIZES_CN[v])
        if v.lower().endswith("pt"):
            return Pt(float(v[:-2].strip()))
        try:
            return Pt(float(v))
        except ValueError:
            raise ValueError(f"Invalid size: '{value}'")
    raise TypeError(f"Expected int, float, or str for size, got {type(value)}")


def parse_indent(chars, font_size=None):
    """Convert character-count indent to Pt.

    chars: number of characters (e.g. 2 for first-line indent).
    font_size: font size in Pt (for calculation). Defaults to 12pt.

    Word convention: 1 char = current font size in points.
    """
    if chars is None:
        return None
    if isinstance(chars, (int, float)):
        fs = font_size or 12
        return Pt(chars * fs)
    if isinstance(chars, str):
        v = chars.strip()
        if v.lower().endswith("pt"):
            return Pt(float(v[:-2].strip()))
        if v.lower().endswith("cm"):
            from docx.shared import Cm
            return Cm(float(v[:-2].strip()))
        try:
            return parse_indent(float(v), font_size)
        except ValueError:
            raise ValueError(f"Invalid indent: '{chars}'")
    raise TypeError(f"Expected int, float, or str for indent, got {type(chars)}")
