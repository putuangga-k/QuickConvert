"""Build script producing standalone dist/QuickConvertSetup.exe (no Python or Admin required).

Steps:
1. Ensure assets/quickconvert.ico exists.
2. Freeze main.py into a fast --onedir --windowed app bundle (dist/QuickConvert/QuickConvert.exe).
3. Compress dist/QuickConvert/ into a staged ZIP payload and export menu_specs.json.
4. Freeze installer_gui.py into a single-file dist/QuickConvertSetup.exe embedding the payload.
"""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from install_menu import build_registry_specs
from quickconvert.utils import rmtree_robust


PROJECT_ROOT = Path(__file__).resolve().parent
DIST_DIR = PROJECT_ROOT / "dist"
BUILD_DIR = PROJECT_ROOT / "build"
ICON_FILE = PROJECT_ROOT / "assets" / "quickconvert.ico"
CONFIG_FILE = PROJECT_ROOT / "config.json"


def _run(cmd: list[str]) -> None:
    print(">", " ".join(cmd))
    res = subprocess.run(cmd, cwd=str(PROJECT_ROOT), check=False)
    if res.returncode != 0:
        raise RuntimeError(f"Command failed with exit code {res.returncode}: {' '.join(cmd)}")


def _ensure_icon() -> None:
    if not ICON_FILE.is_file():
        _run([sys.executable, str(PROJECT_ROOT / "assets" / "generate_icon.py")])


def _build_app_bundle() -> Path:
    app_dist = DIST_DIR / "QuickConvert"
    if app_dist.exists():
        rmtree_robust(app_dist)

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onedir",
        "--windowed",
        "--name",
        "QuickConvert",
        "--icon",
        str(ICON_FILE),
        "--add-data",
        f"{ICON_FILE};assets",
        "--add-data",
        f"{CONFIG_FILE};.",
        "--collect-binaries",
        "imageio_ffmpeg",
        "--collect-data",
        "imageio_ffmpeg",
        "--collect-binaries",
        "pillow_heif",
        "--collect-submodules",
        "quickconvert",
        "--hidden-import",
        "win32com.client",
        "--hidden-import",
        "pythoncom",
        "--hidden-import",
        "pywintypes",
        "--hidden-import",
        "pdf2docx",
        "--hidden-import",
        "pymupdf",
        "--hidden-import",
        "pymupdf4llm",
        "--hidden-import",
        "PIL._tkinter_finder",
        "--exclude-module",
        "pytest",
        "--exclude-module",
        "IPython",
        str(PROJECT_ROOT / "main.py"),
    ]
    _run(cmd)

    # Also keep config.json & assets/quickconvert.ico directly accessible next to QuickConvert.exe
    shutil.copy2(CONFIG_FILE, app_dist / "config.json")
    (app_dist / "assets").mkdir(exist_ok=True)
    shutil.copy2(ICON_FILE, app_dist / "assets" / "quickconvert.ico")

    exe_path = app_dist / "QuickConvert.exe"
    if not exe_path.is_file():
        raise RuntimeError(f"Expected frozen executable not found: {exe_path}")
    return app_dist


def _export_menu_specs_json(dst_json: Path) -> None:
    placeholder_dir = Path(r"C:\__QC_INSTALL_DIR__")
    specs = build_registry_specs(
        executable_override=placeholder_dir / "QuickConvert.exe",
        icon_override=placeholder_dir / "assets" / "quickconvert.ico",
    )
    serialized = []
    for spec in specs:
        serialized.append(
            {
                "source_ext": spec.source_ext,
                "base_key_path": spec.base_key_path,
                "icon_path": spec.icon_path.replace(str(placeholder_dir), "{INSTALL_DIR}"),
                "actions": [
                    {
                        "target_ext": a.target_ext,
                        "subkey_name": a.subkey_name,
                        "label": a.label,
                        "command": a.command.replace(str(placeholder_dir), "{INSTALL_DIR}"),
                        "separator_before": a.separator_before,
                    }
                    for a in spec.actions
                ],
            }
        )
    dst_json.write_text(json.dumps(serialized, indent=2), encoding="utf-8")


def _zip_directory(src_dir: Path, zip_path: Path) -> None:
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for file_path in sorted(src_dir.rglob("*")):
            if file_path.is_file():
                zf.write(file_path, arcname=file_path.relative_to(src_dir))


def build_installer() -> Path:
    _ensure_icon()
    app_dist = _build_app_bundle()

    stage_dir = Path(tempfile.mkdtemp(prefix="quickconvert_build_stage_"))
    try:
        payload_zip = stage_dir / "QuickConvert_payload.zip"
        specs_json = stage_dir / "menu_specs.json"

        print("Compressing QuickConvert bundle into payload archive...")
        _zip_directory(app_dist, payload_zip)
        _export_menu_specs_json(specs_json)

        setup_exe = DIST_DIR / "QuickConvertSetup.exe"
        if setup_exe.exists():
            setup_exe.unlink()

        cmd = [
            sys.executable,
            "-m",
            "PyInstaller",
            "--noconfirm",
            "--clean",
            "--onefile",
            "--windowed",
            "--name",
            "QuickConvertSetup",
            "--icon",
            str(ICON_FILE),
            "--add-data",
            f"{payload_zip};.",
            "--add-data",
            f"{specs_json};.",
            "--add-data",
            f"{ICON_FILE};assets",
            "--exclude-module",
            "PIL",
            "--exclude-module",
            "numpy",
            "--exclude-module",
            "pymupdf",
            "--exclude-module",
            "cv2",
            "--exclude-module",
            "onnxruntime",
            str(PROJECT_ROOT / "installer_gui.py"),
        ]
        _run(cmd)

        if not setup_exe.is_file():
            raise RuntimeError(f"Expected installer executable not found: {setup_exe}")

        size_mb = setup_exe.stat().st_size / (1024 * 1024)
        print(f"\nSUCCESS: Built {setup_exe} ({size_mb:.1f} MB)")
        return setup_exe
    finally:
        rmtree_robust(stage_dir)


if __name__ == "__main__":
    build_installer()
