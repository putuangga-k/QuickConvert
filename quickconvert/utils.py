"""Shared utilities: cancellation, atomic output reservation, and cross-drive commits."""

from __future__ import annotations

import ctypes
import errno
import os
import shutil
import stat
import subprocess
import sys
import threading
import time
import uuid
from pathlib import Path


NO_WINDOW_FLAGS: int = getattr(subprocess, "CREATE_NO_WINDOW", 0) if sys.platform == "win32" else 0


class ConversionError(RuntimeError):
    """Raised when a conversion step fails."""


class ConversionCancelledError(ConversionError):
    """Raised when a conversion is cancelled by the user."""


class CancellationToken:
    """Thread-safe cancellation signal shared between the UI and converters."""

    def __init__(self) -> None:
        self._event = threading.Event()

    def cancel(self) -> None:
        self._event.set()

    @property
    def is_cancelled(self) -> bool:
        return self._event.is_set()

    def raise_if_cancelled(self) -> None:
        if self._event.is_set():
            raise ConversionCancelledError("Conversion cancelled by user.")


def normalize_ext(ext: str) -> str:
    """Normalize an extension string to lowercase without a leading dot."""
    return ext.strip().lower().lstrip(".")


def reserve_unique_file(directory: Path, stem: str, ext: str) -> Path:
    """Atomically reserve a non-colliding file path using O_CREAT | O_EXCL.

    Auto-increments `<stem>.<ext>` -> `<stem> (1).<ext>`, `<stem> (2).<ext>`, etc.
    """
    ext_clean = normalize_ext(ext)
    directory.mkdir(parents=True, exist_ok=True)
    counter = 0
    while True:
        filename = f"{stem}.{ext_clean}" if counter == 0 else f"{stem} ({counter}).{ext_clean}"
        candidate = directory / filename
        try:
            fd = os.open(str(candidate), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            return candidate
        except FileExistsError:
            counter += 1


def reserve_unique_dir(directory: Path, base_name: str) -> Path:
    """Atomically reserve a non-colliding directory path using os.mkdir.

    Auto-increments `<base_name>` -> `<base_name> (1)`, `<base_name> (2)`, etc.
    """
    directory.mkdir(parents=True, exist_ok=True)
    counter = 0
    while True:
        dirname = base_name if counter == 0 else f"{base_name} ({counter})"
        candidate = directory / dirname
        try:
            os.mkdir(candidate)
            return candidate
        except FileExistsError:
            counter += 1


def commit_file_atomic(staged_src: Path, reserved_dst: Path) -> Path:
    """Atomically replace `reserved_dst` with `staged_src`, handling cross-drive moves.

    When `staged_src` and `reserved_dst` reside on different volumes (e.g., C: TEMP
    and D: workspace), copies into a hidden sibling temporary file in the destination
    directory first and then atomically calls `os.replace()` onto `reserved_dst`.
    """
    try:
        os.replace(staged_src, reserved_dst)
        return reserved_dst
    except OSError as exc:
        # Windows raises WinError 17 (ERROR_NOT_SAME_DEVICE) / errno.EXDEV across drives.
        if exc.errno not in (errno.EXDEV,) and getattr(exc, "winerror", None) != 17:
            raise

    sibling_tmp = reserved_dst.parent / f".quickconvert_commit_{uuid.uuid4().hex}.tmp"
    try:
        shutil.copy2(staged_src, sibling_tmp)
        os.replace(sibling_tmp, reserved_dst)
        return reserved_dst
    finally:
        if sibling_tmp.exists():
            try:
                sibling_tmp.unlink()
            except OSError:
                pass


def commit_directory_contents(staged_dir: Path, reserved_dst_dir: Path) -> Path:
    """Commit all files from a staged output directory into the reserved destination folder."""
    entries = sorted(staged_dir.iterdir(), key=lambda p: p.name)
    for entry in entries:
        if entry.is_file():
            target_file = reserved_dst_dir / entry.name
            commit_file_atomic(entry, target_file)
    return reserved_dst_dir


def _clear_readonly_and_retry(func, path: str, _exc: BaseException) -> None:
    """Clear Windows read-only/hidden attributes (e.g. LibreOffice lockfiles) and retry removal."""
    try:
        if sys.platform == "win32":
            # FILE_ATTRIBUTE_NORMAL = 0x80 clears Read-Only, Hidden, and System flags
            ctypes.windll.kernel32.SetFileAttributesW(str(path), 0x80)
        os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
        func(path)
    except OSError:
        pass


def rmtree_robust(path: Path, retries: int = 4, delay: float = 0.15) -> None:
    """Reliably delete a directory tree on Windows, handling read-only files and brief locks."""
    if not path.exists():
        return
    for attempt in range(retries):
        try:
            shutil.rmtree(path, onexc=_clear_readonly_and_retry)
        except TypeError:
            # Fallback for Python < 3.12
            shutil.rmtree(path, onerror=lambda f, p, e: _clear_readonly_and_retry(f, p, e[1]))
        except OSError:
            pass
        if not path.exists():
            return
        if attempt + 1 < retries:
            time.sleep(delay * (attempt + 1))

