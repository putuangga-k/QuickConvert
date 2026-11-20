"""Runtime path helpers supporting both source execution and frozen PyInstaller builds."""

from __future__ import annotations

import sys
from pathlib import Path


SOURCE_ROOT = Path(__file__).resolve().parent.parent


def is_frozen() -> bool:
    """Return True when running inside a PyInstaller-frozen executable."""
    return bool(getattr(sys, "frozen", False))


def get_app_dir() -> Path:
    """Return the directory containing QuickConvert.exe (frozen) or the repo root (source)."""
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return SOURCE_ROOT


def get_resource_path(relative_path: str) -> Path:
    """Resolve a bundled resource file across installed app dir, PyInstaller _MEIPASS, and source."""
    app_candidate = get_app_dir() / relative_path
    if app_candidate.exists():
        return app_candidate.resolve()

    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        mei_candidate = Path(meipass) / relative_path
        if mei_candidate.exists():
            return mei_candidate.resolve()

    return (SOURCE_ROOT / relative_path).resolve()
