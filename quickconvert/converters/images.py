"""Universal offline image converter (Pillow + pillow-heif) with EXIF/ICC preservation."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from PIL import Image, ImageOps, ImageSequence

from quickconvert.config import AppConfig
from quickconvert.graph import (
    COST_DIRECT_NATIVE,
    COST_HIGH_FIDELITY,
    ConversionGraph,
    SubProgressCallback,
)
from quickconvert.utils import CancellationToken, ConversionError

try:
    import pillow_heif  # type: ignore[import-untyped]

    pillow_heif.register_heif_opener()
except Exception:
    pass


IMAGE_SOURCES: tuple[str, ...] = (
    "png",
    "jpg",
    "jpeg",
    "webp",
    "heic",
    "heif",
    "avif",
    "bmp",
    "ico",
    "tiff",
    "tif",
)
IMAGE_TARGETS: tuple[str, ...] = (
    "png",
    "jpg",
    "jpeg",
    "webp",
    "avif",
    "heic",
    "heif",
    "tiff",
    "tif",
    "bmp",
    "ico",
)
_HIDDEN_ALIAS_TARGETS = {"jpeg", "heif", "tif"}
_ICO_STANDARD_SIZES = [256, 128, 64, 48, 32, 16]


def _flatten_to_rgb(img: Image.Image) -> Image.Image:
    """Flatten transparency onto a solid white background and return an RGB image."""
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        rgba = img.convert("RGBA")
        background = Image.new("RGB", rgba.size, (255, 255, 255))
        background.paste(rgba, mask=rgba.split()[3])
        return background
    if img.mode != "RGB":
        return img.convert("RGB")
    return img


def _prepare_ico_canvas(img: Image.Image) -> tuple[Image.Image, list[tuple[int, int]]]:
    """Center non-square images onto a transparent square RGBA canvas for ICO export."""
    rgba = img.convert("RGBA")
    w, h = rgba.size
    side = max(w, h)
    if w != h:
        canvas = Image.new("RGBA", (side, side), (0, 0, 0, 0))
        offset = ((side - w) // 2, (side - h) // 2)
        canvas.paste(rgba, offset, rgba)
        square = canvas
    else:
        square = rgba

    max_dim = min(side, 256)
    if side > 256:
        square = square.resize((256, 256), Image.Resampling.LANCZOS)

    sizes = [(s, s) for s in _ICO_STANDARD_SIZES if s <= max_dim]
    if not sizes:
        sizes = [(max_dim, max_dim)]
    elif (max_dim, max_dim) not in sizes and max_dim < 256:
        sizes.insert(0, (max_dim, max_dim))

    return square, sizes


def _save_single_image(
    img: Image.Image,
    dst_path: Path,
    ext: str,
    config: AppConfig,
    icc_profile: bytes | None,
) -> None:
    kwargs: dict[str, Any] = {}
    if icc_profile and ext in ("png", "jpg", "jpeg", "webp", "tiff", "tif", "heic", "heif", "avif"):
        kwargs["icc_profile"] = icc_profile

    if ext == "ico":
        square, sizes = _prepare_ico_canvas(img)
        square.save(dst_path, format="ICO", sizes=sizes)
    elif ext in ("jpg", "jpeg"):
        rgb = _flatten_to_rgb(img)
        rgb.save(dst_path, format="JPEG", quality=config.jpeg_quality, optimize=True, **kwargs)
    elif ext == "png":
        img.save(dst_path, format="PNG", optimize=True, **kwargs)
    elif ext == "webp":
        img.save(dst_path, format="WEBP", quality=config.webp_quality, **kwargs)
    elif ext in ("heic", "heif"):
        rgb = _flatten_to_rgb(img) if img.mode not in ("RGB", "RGBA") else img
        rgb.save(dst_path, format="HEIF", quality=config.jpeg_quality, **kwargs)
    elif ext == "avif":
        rgb = img.convert("RGBA") if "A" in img.getbands() else _flatten_to_rgb(img)
        rgb.save(dst_path, format="AVIF", quality=config.webp_quality, **kwargs)
    elif ext == "bmp":
        rgb = _flatten_to_rgb(img)
        rgb.save(dst_path, format="BMP")
    elif ext in ("tiff", "tif"):
        img.save(dst_path, format="TIFF", compression="tiff_lzw", **kwargs)
    elif ext == "pdf":
        rgb = _flatten_to_rgb(img)
        rgb.save(dst_path, format="PDF", resolution=float(config.image_dpi))
    else:
        raise ConversionError(f"Unsupported image target format: .{ext}")


def _make_image_converter(src_ext: str, target_ext: str):
    ext = target_ext.lower().lstrip(".")

    def _convert(
        src_path: Path,
        dst_path: Path,
        config: AppConfig,
        cancel_token: CancellationToken,
        sub_progress_cb: SubProgressCallback | None = None,
    ) -> None:
        cancel_token.raise_if_cancelled()
        with Image.open(src_path) as opened:
            icc_profile = opened.info.get("icc_profile")
            n_frames = getattr(opened, "n_frames", 1)

            # Multi-page TIFF -> multi-page PDF
            if src_ext in ("tiff", "tif") and ext == "pdf" and n_frames > 1:
                rgb_pages: list[Image.Image] = []
                try:
                    for idx, frame in enumerate(ImageSequence.Iterator(opened), start=1):
                        cancel_token.raise_if_cancelled()
                        transposed = ImageOps.exif_transpose(frame.copy())
                        rgb_pages.append(_flatten_to_rgb(transposed))
                        if sub_progress_cb:
                            sub_progress_cb(idx / n_frames, f"Page {idx}/{n_frames}")
                    rgb_pages[0].save(
                        dst_path,
                        format="PDF",
                        save_all=True,
                        append_images=rgb_pages[1:],
                        resolution=float(config.image_dpi),
                    )
                finally:
                    for p in rgb_pages:
                        try:
                            p.close()
                        except Exception:
                            pass
                return

            # Multi-page TIFF -> per-page PNG / JPG folder
            if src_ext in ("tiff", "tif") and ext in ("png", "jpg", "jpeg") and n_frames > 1:
                dst_path.mkdir(parents=True, exist_ok=True)
                for idx, frame in enumerate(ImageSequence.Iterator(opened), start=1):
                    cancel_token.raise_if_cancelled()
                    transposed = ImageOps.exif_transpose(frame.copy())
                    page_out = dst_path / f"page_{idx:03d}.{ext}"
                    _save_single_image(transposed, page_out, ext, config, icc_profile)
                    if sub_progress_cb:
                        sub_progress_cb(idx / n_frames, f"Page {idx}/{n_frames}")
                return

            img = ImageOps.exif_transpose(opened)
            cancel_token.raise_if_cancelled()
            _save_single_image(img, dst_path, ext, config, icc_profile)

        if sub_progress_cb:
            sub_progress_cb(1.0, None)

    return _convert


def register_image_converters(graph: ConversionGraph) -> None:
    """Register all universal image-to-image and image-to-PDF edges (1-hop, non-chaining)."""
    for src in IMAGE_SOURCES:
        for dst in IMAGE_TARGETS:
            if src == dst:
                continue
            # Avoid offering same-format alias conversions in menu when already canonical
            is_tiff_per_page = src in ("tiff", "tif") and dst in ("png", "jpg")
            graph.add_edge(
                src=src,
                dst=dst,
                converter=_make_image_converter(src, dst),
                label=f"{src.upper()} to {dst.upper()}",
                fidelity_cost=COST_DIRECT_NATIVE if dst in ("png", "tiff", "bmp") else COST_HIGH_FIDELITY,
                may_output_directory=is_tiff_per_page,
                directory_suffix="_pages",
                allow_as_intermediate=False,
                menu_visible=(dst not in _HIDDEN_ALIAS_TARGETS),
                category="Image",
            )

        graph.add_edge(
            src=src,
            dst="pdf",
            converter=_make_image_converter(src, "pdf"),
            label=f"{src.upper()} to PDF",
            fidelity_cost=COST_HIGH_FIDELITY,
            allow_as_intermediate=False,
            menu_visible=True,
            category="Document",
        )
