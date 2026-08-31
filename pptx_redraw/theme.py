"""Theme and style constants for drawing the slide."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Color:
    """RGB color container.

    Args:
        r: Red channel.
        g: Green channel.
        b: Blue channel.

    Returns:
        None.

    Raises:
        ValueError: Not raised by this dataclass directly.
    """

    r: int
    g: int
    b: int


BG = Color(248, 250, 252)
PANEL = Color(255, 255, 255)
TITLE = Color(15, 23, 42)
TEXT = Color(30, 41, 59)
MUTED = Color(71, 85, 105)
BORDER = Color(203, 213, 225)
ACCENT = Color(30, 136, 229)
ACCENT_SOFT = Color(227, 242, 253)
GOOD = Color(46, 125, 50)
PURPLE = Color(124, 58, 237)
ORANGE = Color(234, 88, 12)
TEAL = Color(13, 148, 136)

FONT_CJK = "Microsoft JhengHei"
FONT_UI = "Calibri"
