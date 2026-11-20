"""Generate the multi-resolution Windows 11 Fluent-style QuickConvert app icon.

Renders at 1024x1024 (4x supersampling) and downsamples via LANCZOS into
16, 24, 32, 48, 64, 128, and 256 px frames embedded in assets/quickconvert.ico.
"""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw


ASSETS_DIR = Path(__file__).resolve().parent
OUTPUT_ICO = ASSETS_DIR / "quickconvert.ico"
ICON_SIZES: tuple[int, ...] = (16, 24, 32, 48, 64, 128, 256)

# Windows 11 Fluent Deep Cobalt palette
COBALT_FILL = (0, 103, 192, 255)          # #0067C0
INNER_BORDER = (255, 255, 255, 52)        # Subtle crisp 1px inner edge
GLYPH_WHITE = (255, 255, 255, 255)        # High-contrast white mark


def _draw_horizontal_arrow(
    draw: ImageDraw.ImageDraw,
    x_left: int,
    x_right: int,
    y_center: int,
    direction: str,
    shaft_half: int,
    head_len: int,
    head_half: int,
    fill: tuple[int, int, int, int],
) -> None:
    """Draw a crisp geometric horizontal arrow with sharp corners and thick strokes."""
    if direction == "right":
        shaft_end = x_right - head_len + int(shaft_half * 0.45)
        draw.rounded_rectangle(
            [x_left, y_center - shaft_half, shaft_end, y_center + shaft_half],
            radius=shaft_half // 2,
            fill=fill,
        )
        head_poly = [
            (x_right, y_center),
            (x_right - head_len, y_center - head_half),
            (x_right - head_len, y_center + head_half),
        ]
        draw.polygon(head_poly, fill=fill)
    else:
        shaft_start = x_left + head_len - int(shaft_half * 0.45)
        draw.rounded_rectangle(
            [shaft_start, y_center - shaft_half, x_right, y_center + shaft_half],
            radius=shaft_half // 2,
            fill=fill,
        )
        head_poly = [
            (x_left, y_center),
            (x_left + head_len, y_center - head_half),
            (x_left + head_len, y_center + head_half),
        ]
        draw.polygon(head_poly, fill=fill)


def render_master_icon(size: int = 1024) -> Image.Image:
    """Render the high-resolution master icon at 1024x1024."""
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img, "RGBA")

    # Clean rounded square base with slight inset for crisp anti-aliased edges
    margin = int(size * 0.045)
    radius = int(size * 0.21)
    box = [margin, margin, size - margin - 1, size - margin - 1]
    draw.rounded_rectangle(box, radius=radius, fill=COBALT_FILL)

    # Subtle inner border (scales cleanly to ~1px at standard UI sizes)
    border_w = max(1, int(size * 0.018))
    inset = margin + border_w // 2
    border_box = [inset, inset, size - inset - 1, size - inset - 1]
    draw.rounded_rectangle(
        border_box,
        radius=max(1, radius - border_w // 2),
        outline=INNER_BORDER,
        width=border_w,
    )

    # Bold geometric bidirectional conversion mark (stays legible even at 16x16)
    x_left = int(size * 0.22)
    x_right = int(size * 0.78)
    shaft_half = int(size * 0.056)   # ~115px thick shaft at 1024px -> ~1.8px at 16px
    head_len = int(size * 0.20)
    head_half = int(size * 0.145)

    y_top = int(size * 0.365)
    y_bot = int(size * 0.635)

    # Top arrow pointing right (->)
    _draw_horizontal_arrow(
        draw,
        x_left=x_left,
        x_right=x_right,
        y_center=y_top,
        direction="right",
        shaft_half=shaft_half,
        head_len=head_len,
        head_half=head_half,
        fill=GLYPH_WHITE,
    )

    # Bottom arrow pointing left (<-)
    _draw_horizontal_arrow(
        draw,
        x_left=x_left,
        x_right=x_right,
        y_center=y_bot,
        direction="left",
        shaft_half=shaft_half,
        head_len=head_len,
        head_half=head_half,
        fill=GLYPH_WHITE,
    )

    return img


def generate_icon(output_path: Path = OUTPUT_ICO) -> Path:
    """Generate and save multi-resolution assets/quickconvert.ico."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    master = render_master_icon(1024)

    # Explicitly downsample every target frame from 1024x1024 using LANCZOS
    frames = [
        master.resize((dim, dim), Image.Resampling.LANCZOS)
        for dim in ICON_SIZES
    ]

    # Save largest frame as primary and embed all other pre-downsampled frames
    largest = frames[-1]
    rest = frames[:-1]
    largest.save(
        output_path,
        format="ICO",
        bitmap_format="bmp",
        append_images=rest,
    )
    return output_path


if __name__ == "__main__":
    out = generate_icon()
    with Image.open(out) as verify:
        embedded = sorted(verify.info.get("sizes", set()))
        print(f"Generated {out} ({out.stat().st_size} bytes) with frames: {embedded}")
