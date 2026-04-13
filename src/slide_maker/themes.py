"""Theme definitions for slide generation."""
from __future__ import annotations

from dataclasses import dataclass

from pptx.dml.color import RGBColor


@dataclass(frozen=True)
class Theme:
    # Background colors
    slide_bg: RGBColor          # main slide background
    header_bg: RGBColor         # header bar background
    accent: RGBColor            # accent / highlight color

    # Text colors
    text_on_dark: RGBColor      # text on dark backgrounds
    text_on_light: RGBColor     # text on light backgrounds
    text_muted: RGBColor        # secondary text (subtitles, etc.)

    # Font sizes (pt)
    title_size: int = 44
    heading_size: int = 28
    body_size: int = 18
    bullet_size: int = 16
    caption_size: int = 13

    # Bullet marker character
    bullet_char: str = "●"

    # Font family (None = PowerPoint default)
    font_name: str | None = None


THEMES: dict[str, Theme] = {
    # AWS brand palette
    "aws": Theme(
        slide_bg=RGBColor(0xFF, 0xFF, 0xFF),
        header_bg=RGBColor(0x23, 0x2F, 0x3E),   # AWS Squid Ink
        accent=RGBColor(0xFF, 0x99, 0x00),        # AWS Orange
        text_on_dark=RGBColor(0xFF, 0xFF, 0xFF),
        text_on_light=RGBColor(0x16, 0x1E, 0x2D),
        text_muted=RGBColor(0xCC, 0xCC, 0xCC),
        bullet_char="▶",
    ),
    # Clean minimal (dark navy + teal)
    "minimal": Theme(
        slide_bg=RGBColor(0xFF, 0xFF, 0xFF),
        header_bg=RGBColor(0x1A, 0x1A, 0x2E),
        accent=RGBColor(0x0F, 0x3A, 0x5C),
        text_on_dark=RGBColor(0xFF, 0xFF, 0xFF),
        text_on_light=RGBColor(0x1A, 0x1A, 0x2E),
        text_muted=RGBColor(0xAA, 0xAA, 0xAA),
        bullet_char="◆",
    ),
    # Dark theme
    "dark": Theme(
        slide_bg=RGBColor(0x0D, 0x1B, 0x2A),
        header_bg=RGBColor(0x1B, 0x2C, 0x3C),
        accent=RGBColor(0x00, 0xB4, 0xD8),
        text_on_dark=RGBColor(0xFF, 0xFF, 0xFF),
        text_on_light=RGBColor(0xFF, 0xFF, 0xFF),
        text_muted=RGBColor(0x99, 0xBB, 0xCC),
        bullet_char="◆",
    ),
}

# Alias
THEMES["default"] = THEMES["aws"]


def get_theme(name: str) -> Theme:
    return THEMES.get(name, THEMES["aws"])
