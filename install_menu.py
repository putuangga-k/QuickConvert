"""Install QuickConvert cascading context menus under HKCU (no Admin privileges required)."""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass
from pathlib import Path

from quickconvert.pipeline import build_default_graph

if sys.platform == "win32":
    import winreg


PROJECT_ROOT = Path(__file__).resolve().parent
MAIN_SCRIPT = PROJECT_ROOT / "main.py"
ICON_PATH = (PROJECT_ROOT / "assets" / "quickconvert.ico").resolve()
MANIFEST_PATH = PROJECT_ROOT / "installed_registry_keys.json"
REG_ROOT_PREFIX = r"Software\Classes\SystemFileAssociations"
MENU_KEY_NAME = "QuickConvert"

# Windows Explorer truncates static Registry submenus beyond 16 items.
# Cap direct quick targets at 10 + 1 "More Formats..." picker = <= 11 items total.
MAX_QUICK_SUBMENU_TARGETS = 10
ECF_SEPARATORBEFORE = 0x20


@dataclass(frozen=True)
class RegistryActionSpec:
    """Single submenu entry under QuickConvert\\shell\\<subkey_name>."""

    target_ext: str
    subkey_name: str
    label: str
    command: str
    separator_before: bool = False


@dataclass(frozen=True)
class ExtensionMenuSpec:
    """Complete cascading menu registration spec for a single source extension."""

    source_ext: str
    base_key_path: str
    icon_path: str
    actions: tuple[RegistryActionSpec, ...]


def resolve_pythonw_executable() -> Path:
    """Locate pythonw.exe (preferring the project's .venv) so conversions run silently."""
    candidates = [
        PROJECT_ROOT / ".venv" / "Scripts" / "pythonw.exe",
        Path(sys.executable).resolve().parent / "pythonw.exe",
    ]
    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()
    raise FileNotFoundError(
        "Could not locate pythonw.exe in .venv\\Scripts or alongside sys.executable. "
        "QuickConvert requires pythonw.exe for silent context menu execution."
    )


def build_registry_specs(
    pythonw_path: Path | None = None,
    executable_override: Path | None = None,
    icon_override: Path | None = None,
) -> list[ExtensionMenuSpec]:
    """Build all HKCU\\Software\\Classes\\SystemFileAssociations menu specs from the graph.

    In source mode (default), launches main.py via pythonw.exe so no console flashes.
    When `executable_override` is provided (frozen QuickConvert.exe GUI app), invokes it directly.
    """
    icon_str = str((icon_override if icon_override is not None else ICON_PATH).resolve())
    graph = build_default_graph()

    if executable_override is not None:
        exe_resolved = executable_override.resolve()
        cmd_prefix = f'"{exe_resolved}"'
    else:
        pyw = pythonw_path if pythonw_path is not None else resolve_pythonw_executable()
        main_py = MAIN_SCRIPT.resolve()
        cmd_prefix = f'"{pyw}" "{main_py}"'

    specs: list[ExtensionMenuSpec] = []
    for src_ext in graph.get_supported_sources():
        quick_targets = graph.get_quick_targets(src_ext, max_items=MAX_QUICK_SUBMENU_TARGETS)
        if not quick_targets:
            continue

        base_key = rf"{REG_ROOT_PREFIX}\.{src_ext}\shell\{MENU_KEY_NAME}"
        action_list: list[RegistryActionSpec] = []
        for idx, opt in enumerate(quick_targets, start=1):
            action_list.append(
                RegistryActionSpec(
                    target_ext=opt.target_ext,
                    subkey_name=f"{idx:02d}_to_{opt.target_ext}",
                    label=opt.menu_label,
                    command=f'{cmd_prefix} --target {opt.target_ext} "%1"',
                    separator_before=False,
                )
            )

        # Always append "More Formats..." picker with a native separator line above it
        action_list.append(
            RegistryActionSpec(
                target_ext="__more__",
                subkey_name="99_more_formats",
                label="More Formats...",
                command=f'{cmd_prefix} --choose-target "%1"',
                separator_before=True,
            )
        )

        specs.append(
            ExtensionMenuSpec(
                source_ext=src_ext,
                base_key_path=base_key,
                icon_path=icon_str,
                actions=tuple(action_list),
            )
        )
    return specs


def _delete_key_recursive(root_hive: int, subkey_path: str) -> bool:
    """Recursively delete an existing Registry subkey tree so reinstalls never leave stale items."""
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


def install_context_menus(
    dry_run: bool = False,
    executable_override: Path | None = None,
    icon_override: Path | None = None,
    manifest_override: Path | None = None,
) -> list[ExtensionMenuSpec]:
    """Write all QuickConvert context menu keys strictly under HKEY_CURRENT_USER."""
    specs = build_registry_specs(
        executable_override=executable_override,
        icon_override=icon_override,
    )
    manifest_file = manifest_override if manifest_override is not None else MANIFEST_PATH

    if dry_run:
        for spec in specs:
            print(rf"[DRY-RUN] HKCU\{spec.base_key_path} (Icon={spec.icon_path}, Items={len(spec.actions)})")
            for action in spec.actions:
                sep = " [separator]" if action.separator_before else ""
                print(f"    -> [{action.subkey_name}] {action.label}{sep}: {action.command}")
        print(f"[DRY-RUN] Total extensions: {len(specs)}")
        return specs

    if sys.platform != "win32":
        raise RuntimeError("Windows Registry installation is only supported on Windows.")

    installed_paths: list[str] = []
    for spec in specs:
        # Remove any stale previous QuickConvert subkeys for this extension first
        _delete_key_recursive(winreg.HKEY_CURRENT_USER, spec.base_key_path)

        with winreg.CreateKeyEx(
            winreg.HKEY_CURRENT_USER,
            spec.base_key_path,
            0,
            winreg.KEY_WRITE,
        ) as menu_key:
            winreg.SetValueEx(menu_key, "MUIVerb", 0, winreg.REG_SZ, "QuickConvert")
            winreg.SetValueEx(menu_key, "SubCommands", 0, winreg.REG_SZ, "")
            winreg.SetValueEx(menu_key, "Icon", 0, winreg.REG_SZ, spec.icon_path)

        for action in spec.actions:
            action_key_path = rf"{spec.base_key_path}\shell\{action.subkey_name}"
            with winreg.CreateKeyEx(
                winreg.HKEY_CURRENT_USER,
                action_key_path,
                0,
                winreg.KEY_WRITE,
            ) as action_key:
                winreg.SetValueEx(action_key, "MUIVerb", 0, winreg.REG_SZ, action.label)
                if action.separator_before:
                    winreg.SetValueEx(action_key, "CommandFlags", 0, winreg.REG_DWORD, ECF_SEPARATORBEFORE)

            cmd_key_path = rf"{action_key_path}\command"
            with winreg.CreateKeyEx(
                winreg.HKEY_CURRENT_USER,
                cmd_key_path,
                0,
                winreg.KEY_WRITE,
            ) as cmd_key:
                winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, action.command)

        installed_paths.append(spec.base_key_path)

    manifest = {
        "hive": "HKEY_CURRENT_USER",
        "keys": installed_paths,
    }
    manifest_file.parent.mkdir(parents=True, exist_ok=True)
    manifest_file.write_text(json.dumps(manifest, indent=2), encoding="utf-8")

    print(f"Installed QuickConvert context menus for {len(specs)} file extensions under HKCU.")
    print(f"Manifest saved to: {manifest_file}")
    print("Note: On Windows 11, open 'Show more options' (or Shift+Right-Click) to see classic submenus.")
    return specs


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Install QuickConvert Windows right-click context menus.")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print Registry keys and commands without modifying the Windows Registry.",
    )
    args = parser.parse_args(argv)
    install_context_menus(dry_run=args.dry_run)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
