"""Mandatory post-conversion output integrity validator."""

from __future__ import annotations

import csv
import json
import subprocess
import zipfile
from pathlib import Path

import yaml
from PIL import Image

from quickconvert.config import AppConfig
from quickconvert.utils import NO_WINDOW_FLAGS, ConversionError, normalize_ext

try:
    import pymupdf as fitz  # type: ignore[import-untyped]
except ImportError:  # pragma: no cover
    import fitz  # type: ignore[import-untyped]

try:
    import pillow_heif  # type: ignore[import-untyped]

    pillow_heif.register_heif_opener()
except Exception:
    pass


_IMAGE_EXTS = {
    "png",
    "jpg",
    "jpeg",
    "webp",
    "bmp",
    "ico",
    "tiff",
    "tif",
    "heic",
    "heif",
    "avif",
    "gif",
}
_OOXML_REQUIRED_ENTRIES = {
    "docx": "word/document.xml",
    "xlsx": "xl/workbook.xml",
    "pptx": "ppt/presentation.xml",
}
_MEDIA_EXTS = {
    "mp4",
    "mkv",
    "mov",
    "webm",
    "avi",
    "m4v",
    "wmv",
    "mp3",
    "wav",
    "flac",
    "m4a",
    "aac",
    "ogg",
    "opus",
    "wma",
}
_TEXT_EXTS = {"txt", "html", "htm", "md"}


def _validate_image(path: Path) -> None:
    try:
        with Image.open(path) as img:
            img.verify()
    except Exception as exc:
        raise ConversionError(f"Output image integrity check failed for '{path.name}': {exc}") from exc


def _validate_pdf(path: Path) -> None:
    doc = None
    try:
        doc = fitz.open(str(path))
        if len(doc) <= 0:
            raise ConversionError(f"Output PDF '{path.name}' contains zero pages.")
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(f"Output PDF integrity check failed for '{path.name}': {exc}") from exc
    finally:
        if doc is not None:
            doc.close()


def _validate_ooxml(path: Path, ext: str) -> None:
    if not zipfile.is_zipfile(path):
        raise ConversionError(f"Output '.{ext}' file '{path.name}' is not a valid OOXML archive.")
    try:
        with zipfile.ZipFile(path, "r") as zf:
            bad = zf.testzip()
            if bad is not None:
                raise ConversionError(f"Corrupt OOXML archive member '{bad}' in '{path.name}'.")
            names = set(zf.namelist())
            if "[Content_Types].xml" not in names:
                raise ConversionError(f"Missing [Content_Types].xml in '{path.name}'.")
            required = _OOXML_REQUIRED_ENTRIES[ext]
            if required not in names:
                raise ConversionError(f"Missing '{required}' in '{path.name}'.")
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(f"OOXML structure check failed for '{path.name}': {exc}") from exc


def _validate_json(path: Path) -> None:
    try:
        text = path.read_text(encoding="utf-8")
        json.loads(text)
    except Exception as exc:
        raise ConversionError(f"Output JSON validation failed for '{path.name}': {exc}") from exc


def _validate_yaml(path: Path) -> None:
    try:
        text = path.read_text(encoding="utf-8")
        parsed = yaml.safe_load(text)
        if parsed is None:
            raise ConversionError(f"Output YAML '{path.name}' parsed as empty.")
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(f"Output YAML validation failed for '{path.name}': {exc}") from exc


def _validate_delimited(path: Path, delimiter: str) -> None:
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as fh:
            rows = list(csv.reader(fh, delimiter=delimiter))
        if not rows:
            raise ConversionError(f"Output delimited file '{path.name}' contains no rows.")
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(f"Output table validation failed for '{path.name}': {exc}") from exc


def _validate_text(path: Path) -> None:
    try:
        text = path.read_text(encoding="utf-8")
        if len(text) == 0:
            raise ConversionError(f"Output text file '{path.name}' is empty.")
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(f"Output UTF-8 text validation failed for '{path.name}': {exc}") from exc


def _validate_rtf(path: Path) -> None:
    try:
        raw = path.read_bytes()
        if not raw.lstrip().startswith(b"{\\rtf"):
            raise ConversionError(f"Output RTF file '{path.name}' lacks valid {{\\rtf header.")
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(f"Output RTF validation failed for '{path.name}': {exc}") from exc


def _validate_media(path: Path, config: AppConfig) -> None:
    from quickconvert.converters.media import find_ffmpeg_executable

    ffmpeg_bin = find_ffmpeg_executable(config)
    cmd = [
        ffmpeg_bin,
        "-hide_banner",
        "-v",
        "error",
        "-i",
        str(path),
        "-t",
        "1",
        "-f",
        "null",
        "-",
    ]
    res = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        creationflags=NO_WINDOW_FLAGS,
        check=False,
    )
    if res.returncode != 0:
        err = res.stderr.strip() or f"ffmpeg exit code {res.returncode}"
        raise ConversionError(f"Output media stream validation failed for '{path.name}': {err}")


def _validate_single_file(path: Path, ext: str, config: AppConfig) -> None:
    if not path.is_file() or path.stat().st_size <= 0:
        raise ConversionError(f"Output file '{path.name}' is missing or empty.")

    if ext in _IMAGE_EXTS:
        _validate_image(path)
    elif ext == "pdf":
        _validate_pdf(path)
    elif ext in _OOXML_REQUIRED_ENTRIES:
        _validate_ooxml(path, ext)
    elif ext == "json":
        _validate_json(path)
    elif ext in ("yaml", "yml"):
        _validate_yaml(path)
    elif ext == "csv":
        _validate_delimited(path, ",")
    elif ext == "tsv":
        _validate_delimited(path, "\t")
    elif ext in _TEXT_EXTS:
        _validate_text(path)
    elif ext == "rtf":
        _validate_rtf(path)
    elif ext in _MEDIA_EXTS:
        _validate_media(path, config)


def validate_output(output_path: Path, target_ext: str, config: AppConfig) -> None:
    """Validate staged conversion output before atomically committing to destination."""
    ext = normalize_ext(target_ext)
    if output_path.is_dir():
        files = sorted(p for p in output_path.iterdir() if p.is_file())
        if not files:
            raise ConversionError(f"Output folder '{output_path.name}' contains no files.")
        for file_path in files:
            file_ext = normalize_ext(file_path.suffix) or ext
            _validate_single_file(file_path, file_ext, config)
    else:
        _validate_single_file(output_path, ext, config)
