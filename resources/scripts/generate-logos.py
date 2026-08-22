#!/usr/bin/env python3
"""Generates the XlsxSharp brand assets in resources/logo.

The brand mark is a rounded tile holding a white grid that reads both as a
spreadsheet and as a sharp sign (#) - "xlsx" plus "sharp". The wordmark is set
in Open Sans (Apache-2.0) and converted to outlines, so the SVGs do not depend
on an installed font.

Usage:
    python3 resources/scripts/generate-logos.py

Requirements: fonttools (glyph outlines) and rsvg-convert (PNG rendering).
"""

import os
import subprocess
import urllib.request

from fontTools.pens.svgPathPen import SVGPathPen
from fontTools.pens.transformPen import TransformPen
from fontTools.ttLib import TTFont

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, "..", "logo"))
FONT_CACHE_DIR = os.path.join(SCRIPT_DIR, ".fonts")
FONT_FILE = "OpenSans-Regular.ttf"
FONT_URL = (
    "https://github.com/googlefonts/opensans/raw/main/fonts/ttf/OpenSans-Regular.ttf"
)

WORD_FIRST = "Xlsx"
WORD_SECOND = "Sharp"

# Brand colours
GREEN = "#21A366"
PURPLE = "#512BD4"
INK = "#243746"
WHITE = "#FFFFFF"
BLACK = "#000000"

# Icon geometry, expressed on a 1000 x 1000 canvas
ICON_SIZE = 1000.0
TILE_INSET = 100.0
TILE_RADIUS = 136.0
GRID_INSET = 218.0
BAR_WIDTH = 84.0
BAR_RADIUS = 10.0
OUTLINE_WIDTH = 40.0


class Variant:
    """One colour flavour of the logo."""

    def __init__(self, name, mark, first, second, outlined=False):
        self.name = name
        self.mark = mark
        self.first = first
        self.second = second
        self.outlined = outlined


VARIANTS = [
    Variant("green", GREEN, GREEN, INK),
    Variant("purple", PURPLE, PURPLE, INK),
    Variant("black", BLACK, BLACK, BLACK, outlined=True),
    Variant("white", WHITE, WHITE, WHITE, outlined=True),
]


def font_path():
    path = os.path.join(FONT_CACHE_DIR, FONT_FILE)
    if not os.path.exists(path):
        os.makedirs(FONT_CACHE_DIR, exist_ok=True)
        urllib.request.urlretrieve(FONT_URL, path)
    return path


class Wordmark:
    """Turns text into SVG path data using the outlines of a font."""

    def __init__(self, path):
        self.font = TTFont(path)
        self.glyphs = self.font.getGlyphSet()
        self.cmap = self.font.getBestCmap()
        self.upem = self.font["head"].unitsPerEm
        self.cap_height = self.font["OS/2"].sCapHeight

    def font_size(self, cap_height):
        return cap_height * self.upem / self.cap_height

    def width(self, text, size):
        scale = size / self.upem
        return sum(self.glyphs[self.cmap[ord(c)]].width for c in text) * scale

    def paths(self, text, size, x, baseline):
        """Returns the path data of every glyph, already placed on the canvas."""
        scale = size / self.upem
        data = []
        pen_x = x
        for char in text:
            glyph = self.glyphs[self.cmap[ord(char)]]
            pen = SVGPathPen(self.glyphs, ntos=lambda v: f"{v:.2f}")
            glyph.draw(TransformPen(pen, (scale, 0, 0, -scale, pen_x, baseline)))
            commands = pen.getCommands()
            if commands:
                data.append(commands)
            pen_x += glyph.width * scale
        return data


def icon_elements(variant, x, y, size, indent):
    """The brand mark: a rounded tile with a grid that doubles as a sharp sign."""
    scale = size / ICON_SIZE

    def px(value):
        return f"{x + value * scale:.3f}"

    def py(value):
        return f"{y + value * scale:.3f}"

    def length(value):
        return f"{value * scale:.3f}"

    pad = " " * indent
    tile_size = ICON_SIZE - 2 * TILE_INSET
    cell = (ICON_SIZE - 2 * GRID_INSET - 2 * BAR_WIDTH) / 3
    bar_offsets = [GRID_INSET + cell, GRID_INSET + 2 * cell + BAR_WIDTH]
    grid_length = ICON_SIZE - 2 * GRID_INSET

    elements = []
    if variant.outlined:
        inset = TILE_INSET + OUTLINE_WIDTH / 2
        elements.append(
            f'<rect x="{px(inset)}" y="{py(inset)}" '
            f'width="{length(tile_size - OUTLINE_WIDTH)}" '
            f'height="{length(tile_size - OUTLINE_WIDTH)}" '
            f'rx="{length(TILE_RADIUS - OUTLINE_WIDTH / 2)}" fill="none" '
            f'stroke="{variant.mark}" stroke-width="{length(OUTLINE_WIDTH)}"/>'
        )
        grid_color = variant.mark
        highlight_opacity = "0.2"
    else:
        elements.append(
            f'<rect x="{px(TILE_INSET)}" y="{py(TILE_INSET)}" '
            f'width="{length(tile_size)}" height="{length(tile_size)}" '
            f'rx="{length(TILE_RADIUS)}" fill="{variant.mark}"/>'
        )
        grid_color = WHITE
        highlight_opacity = "0.3"

    # The centre cell is tinted to give the grid depth and to mark the "active cell".
    elements.append(
        f'<rect x="{px(bar_offsets[0] + BAR_WIDTH)}" '
        f'y="{py(bar_offsets[0] + BAR_WIDTH)}" width="{length(cell)}" '
        f'height="{length(cell)}" fill="{grid_color}" opacity="{highlight_opacity}"/>'
    )

    for offset in bar_offsets:
        elements.append(
            f'<rect x="{px(offset)}" y="{py(GRID_INSET)}" '
            f'width="{length(BAR_WIDTH)}" height="{length(grid_length)}" '
            f'rx="{length(BAR_RADIUS)}" fill="{grid_color}"/>'
        )
        elements.append(
            f'<rect x="{px(GRID_INSET)}" y="{py(offset)}" '
            f'width="{length(grid_length)}" height="{length(BAR_WIDTH)}" '
            f'rx="{length(BAR_RADIUS)}" fill="{grid_color}"/>'
        )

    return "\n".join(pad + element for element in elements)


def svg_document(width, height, body):
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        "<!-- XlsxSharp logo - generated by resources/scripts/generate-logos.py -->\n"
        '<svg xmlns="http://www.w3.org/2000/svg" '
        f'width="{width:.3f}" height="{height:.3f}" '
        f'viewBox="0 0 {width:.3f} {height:.3f}">\n'
        f"{body}\n"
        "</svg>\n"
    )


def build_icon(variant):
    body = icon_elements(variant, 0, 0, ICON_SIZE, 2)
    return svg_document(ICON_SIZE, ICON_SIZE, body)


def text_group(wordmark, variant, size, x, baseline, indent):
    pad = " " * indent
    groups = []
    for text, color in ((WORD_FIRST, variant.first), (WORD_SECOND, variant.second)):
        paths = wordmark.paths(text, size, x, baseline)
        x += wordmark.width(text, size)
        inner = "\n".join(f'{pad}  <path d="{data}"/>' for data in paths)
        groups.append(f'{pad}<g fill="{color}">\n{inner}\n{pad}</g>')
    return "\n".join(groups)


def build_horizontal(wordmark, variant):
    icon = 260.0
    padding = 71.576
    gap = 62.0
    cap = 130.0
    size = wordmark.font_size(cap)
    text_width = wordmark.width(WORD_FIRST + WORD_SECOND, size)

    height = icon + 2 * padding
    width = padding + icon + gap + text_width + padding
    baseline = padding + icon / 2 + cap / 2

    body = "\n".join(
        [
            icon_elements(variant, padding, padding, icon, 2),
            text_group(wordmark, variant, size, padding + icon + gap, baseline, 2),
        ]
    )
    return svg_document(width, height, body)


def build_stacked(wordmark, variant):
    icon = 420.0
    padding = 70.0
    gap = 58.0
    cap = 120.0
    size = wordmark.font_size(cap)
    text_width = wordmark.width(WORD_FIRST + WORD_SECOND, size)

    width = max(text_width, icon) + 2 * padding
    baseline = padding + icon + gap + cap
    height = baseline + padding
    icon_x = (width - icon) / 2
    text_x = (width - text_width) / 2

    body = "\n".join(
        [
            icon_elements(variant, icon_x, padding, icon, 2),
            text_group(wordmark, variant, size, text_x, baseline, 2),
        ]
    )
    return svg_document(width, height, body)


def write(name, content):
    path = os.path.join(OUTPUT_DIR, name)
    with open(path, "w", encoding="utf-8", newline="\n") as file:
        file.write(content)
    return path


def render(svg_name, png_name, height):
    subprocess.run(
        [
            "rsvg-convert",
            "-h",
            str(height),
            "-o",
            os.path.join(OUTPUT_DIR, png_name),
            os.path.join(OUTPUT_DIR, svg_name),
        ],
        check=True,
    )


def main():
    wordmark = Wordmark(font_path())
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    for index, variant in enumerate(VARIANTS, start=1):
        name = f"favicon-{index:02d}"
        write(name + ".svg", build_icon(variant))
        render(name + ".svg", name + ".png", 1389)

    for index, variant in enumerate(VARIANTS, start=5):
        name = f"logotype-a-{index:02d}"
        write(name + ".svg", build_horizontal(wordmark, variant))
        render(name + ".svg", name + ".png", 840)

    for index, variant in enumerate(VARIANTS, start=9):
        name = f"logotype-b-{index:02d}"
        write(name + ".svg", build_stacked(wordmark, variant))
        render(name + ".svg", name + ".png", 1235)

    # Package icon (NuGet) and the banner used in README.md
    render("favicon-01.svg", "nuget-logo.png", 1209)
    render("logotype-a-05.svg", "readme.png", 110)


if __name__ == "__main__":
    main()
