"""Cleanly remove every HKCU Registry key added by install_menu.py."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from install_menu import MANIFEST_PATH, MENU_KEY_NAME, REG_ROOT_PREFIX
from quickconvert.pipeline import build_default_graph

if sys.platform == "win32":
    import winreg


def collect_target_keys(manifest_override: Path | None = None) -> list[str]:
    """Gather all QuickConvert HKCU Registry paths from manifest and the conversion graph."""
    keys: set[str] = set()
    manifest_file = manifest_override if manifest_override is not None else MANIFEST_PATH

    if manifest_file.is_file():
        try:
            data = json.loads(manifest_file.read_text(encoding="utf-8"))
            for key_path in data.get("keys", []):
                if isinstance(key_path, str) and key_path.startswith(REG_ROOT_PREFIX + "\\"):
                    keys.add(key_path)
        except (OSError, json.JSONDecodeError):
            pass

    graph = build_default_graph()
    for ext in graph.get_supported_sources():
        keys.add(rf"{REG_ROOT_PREFIX}\.{ext}\shell\{MENU_KEY_NAME}")

    return sorted(keys)


def _delete_key_recursive(root_hive: int, subkey_path: str) -> bool:
    """Recursively delete `subkey_path` under `root_hive`. Returns True if deleted."""
    try:
        with winreg.OpenKey(root_hive, subkey_path, 0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
            child_names: list[str] = []
            idx = 0
            while True:
                try:
                    child_names.append(winreg.EnumKey(key, idx))
                    idx += 1
                except OSError:
                    break
            for child in child_names:
                _delete_key_recursive(root_hive, rf"{subkey_path}\{child}")
    except FileNotFoundError:
        return False

    winreg.DeleteKey(root_hive, subkey_path)
    return True


def _prune_if_empty(root_hive: int, subkey_path: str) -> bool:
    """Delete a parent key only if it has zero remaining subkeys and zero values."""
    try:
        with winreg.OpenKey(root_hive, subkey_path, 0, winreg.KEY_READ) as key:
            subkey_count, value_count, _ = winreg.QueryInfoKey(key)
    except FileNotFoundError:
        return False

    if subkey_count == 0 and value_count == 0:
        try:
            winreg.DeleteKey(root_hive, subkey_path)
            return True
        except OSError:
            return False
    return False


def uninstall_context_menus(
    dry_run: bool = False,
    manifest_override: Path | None = None,
) -> int:
    """Remove all QuickConvert keys under HKEY_CURRENT_USER\\Software\\Classes\\SystemFileAssociations."""
    manifest_file = manifest_override if manifest_override is not None else MANIFEST_PATH
    target_keys = collect_target_keys(manifest_override=manifest_file)

    if dry_run:
        for key_path in target_keys:
            print(rf"[DRY-RUN] Would delete HKCU\{key_path}")
        print(f"[DRY-RUN] Total keys targeted: {len(target_keys)}")
        return len(target_keys)

    if sys.platform != "win32":
        raise RuntimeError("Windows Registry uninstallation is only supported on Windows.")

    removed = 0
    for key_path in target_keys:
        if _delete_key_recursive(winreg.HKEY_CURRENT_USER, key_path):
            removed += 1
            # Clean up empty parent `shell` and `.<ext>` keys if QuickConvert was the sole entry.
            shell_parent = key_path.rsplit("\\", 1)[0]
            ext_parent = shell_parent.rsplit("\\", 1)[0]
            if _prune_if_empty(winreg.HKEY_CURRENT_USER, shell_parent):
                _prune_if_empty(winreg.HKEY_CURRENT_USER, ext_parent)

    if manifest_file.is_file():
        try:
            manifest_file.unlink()
        except OSError:
            pass

    print(f"Removed {removed} QuickConvert Registry key trees from HKEY_CURRENT_USER.")
    return removed


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Uninstall QuickConvert Windows right-click context menus.")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print Registry keys that would be removed without modifying the Windows Registry.",
    )
    args = parser.parse_args(argv)
    uninstall_context_menus(dry_run=args.dry_run)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
