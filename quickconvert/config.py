"""Configuration loading and defaults for QuickConvert."""

from __future__ import annotations

import json
from dataclasses import dataclass, fields
from pathlib import Path

from quickconvert.paths import get_app_dir, get_resource_path


PROJECT_ROOT = get_app_dir()
DEFAULT_CONFIG_PATH = get_resource_path("config.json")


@dataclass(frozen=True)
class AppConfig:
    """Runtime settings for QuickConvert converters and UI."""

    image_dpi: int = 200
    jpeg_quality: int = 90
    webp_quality: int = 85
    ffmpeg_preset: str = "medium"
    ffmpeg_crf: int = 23
    audio_bitrate: str = "192k"
    ffmpeg_path: str = ""
    soffice_path: str = ""
    auto_close_delay_ms: int = 1800


def load_config(config_path: Path | str | None = None) -> AppConfig:
    """Load AppConfig from JSON, falling back to defaults for missing keys."""
    path = Path(config_path) if config_path is not None else DEFAULT_CONFIG_PATH
    if not path.is_file():
        return AppConfig()

    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return AppConfig()

    if not isinstance(raw, dict):
        return AppConfig()

    defaults = AppConfig()
    valid_names = {f.name for f in fields(AppConfig)}
    merged: dict[str, object] = {}
    for name in valid_names:
        val = raw.get(name, getattr(defaults, name))
        default_val = getattr(defaults, name)
        try:
            if isinstance(default_val, int):
                merged[name] = int(val)
            elif isinstance(default_val, str):
                merged[name] = str(val)
            else:
                merged[name] = val
        except (TypeError, ValueError):
            merged[name] = default_val

    return AppConfig(**merged)  # type: ignore[arg-type]
