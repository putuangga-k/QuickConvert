"""Standalone Windows 11 Fluent GUI installer & uninstaller for QuickConvert (no Admin required).

Bundled into QuickConvertSetup.exe by build_installer.py. Installs QuickConvert into
%LOCALAPPDATA%\\Programs\\QuickConvert, registers cascading context menus exclusively under
HKEY_CURRENT_USER\\Software\\Classes\\SystemFileAssociations, and registers an HKCU entry in
Windows Installed Apps for one-click uninstallation.
"""

from __future__ import annotations

import argparse
import ctypes
import ctypes.wintypes
import json
import os
import queue
import shutil
import stat
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import zipfile
from dataclasses import dataclass
from pathlib import Path
from tkinter import font as tkfont

if sys.platform == "win32":
    import winreg


APP_NAME = "QuickConvert"
APP_VERSION = "1.0.0"
APP_PUBLISHER = "QuickConvert"
APP_USER_MODEL_ID = "QuickConvert.Setup.1"
REG_ROOT_PREFIX = r"Software\Classes\SystemFileAssociations"
UNINSTALL_REG_KEY = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\QuickConvert"
ECF_SEPARATORBEFORE = 0x20

DWMWA_USE_IMMERSIVE_DARK_MODE = 20
DWMWA_CAPTION_COLOR = 35
SPI_GETWORKAREA = 0x0030
NO_WINDOW_FLAGS = getattr(subprocess, "CREATE_NO_WINDOW", 0) if sys.platform == "win32" else 0
DETACHED_FLAGS = getattr(subprocess, "DETACHED_PROCESS", 0) if sys.platform == "win32" else 0


@dataclass(frozen=True)
class SetupTheme:
    is_light: bool
    bg: str
    surface: str
    border: str
    text_primary: str
    text_secondary: str
    accent: str
    accent_hover: str
    accent_text: str
    success: str
    error: str
    track: str
    btn_hover: str
    btn_active: str


LIGHT_THEME = SetupTheme(
    is_light=True,
    bg="#F3F3F3",
    surface="#FFFFFF",
    border="#E5E5E5",
    text_primary="#1B1B1B",
    text_secondary="#5D5D5D",
    accent="#0067C0",
    accent_hover="#005A9E",
    accent_text="#FFFFFF",
    success="#0F7B0F",
    error="#C42B1C",
    track="#E5E5E5",
    btn_hover="#F6F6F6",
    btn_active="#EBEBEB",
)

DARK_THEME = SetupTheme(
    is_light=False,
    bg="#202020",
    surface="#2B2B2B",
    border="#3A3A3A",
    text_primary="#F3F3F3",
    text_secondary="#A6A6A6",
    accent="#60CDFF",
    accent_hover="#4CC2FF",
    accent_text="#000000",
    success="#6CCB5F",
    error="#FF99A4",
    track="#3A3A3A",
    btn_hover="#353535",
    btn_active="#3E3E3E",
)


def _bundle_root() -> Path:
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        return Path(meipass)
    return Path(__file__).resolve().parent


def _default_install_dir() -> Path:
    local_app = os.environ.get("LOCALAPPDATA")
    if local_app:
        return Path(local_app) / "Programs" / APP_NAME
    return Path.home() / "AppData" / "Local" / "Programs" / APP_NAME


def _clear_readonly_and_retry(func, path: str, _exc: BaseException) -> None:
    try:
        if sys.platform == "win32":
            ctypes.windll.kernel32.SetFileAttributesW(str(path), 0x80)
        os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
        func(path)
    except OSError:
        pass


def _rmtree_robust(path: Path, retries: int = 4) -> None:
    if not path.exists():
        return
    for attempt in range(retries):
        try:
            shutil.rmtree(path, onexc=_clear_readonly_and_retry)
        except TypeError:
            shutil.rmtree(path, onerror=lambda f, p, e: _clear_readonly_and_retry(f, p, e[1]))
        except OSError:
            pass
        if not path.exists():
            return
        if attempt + 1 < retries:
            time.sleep(0.15 * (attempt + 1))


def _delete_reg_tree(root_hive: int, subkey_path: str) -> bool:
    try:
        with winreg.OpenKey(root_hive, subkey_path, 0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
            children: list[str] = []
            idx = 0
            while True:
                try:
                    children.append(winreg.EnumKey(key, idx))
                    idx += 1
                except OSError:
                    break
            for child in children:
                _delete_reg_tree(root_hive, rf"{subkey_path}\{child}")
    except FileNotFoundError:
        return False
    winreg.DeleteKey(root_hive, subkey_path)
    return True


def _prune_empty_reg_key(root_hive: int, subkey_path: str) -> bool:
    try:
        with winreg.OpenKey(root_hive, subkey_path, 0, winreg.KEY_READ) as key:
            subkeys, values, _ = winreg.QueryInfoKey(key)
    except FileNotFoundError:
        return False
    if subkeys == 0 and values == 0:
        try:
            winreg.DeleteKey(root_hive, subkey_path)
            return True
        except OSError:
            return False
    return False


def _load_menu_specs() -> list[dict]:
    specs_file = _bundle_root() / "menu_specs.json"
    if not specs_file.is_file():
        raise FileNotFoundError(f"Missing bundled menu_specs.json at {specs_file}")
    return json.loads(specs_file.read_text(encoding="utf-8"))


def _is_installed(install_dir: Path) -> bool:
    return (install_dir / "QuickConvert.exe").is_file() or (install_dir / "installed_registry_keys.json").is_file()


def perform_install(install_dir: Path, progress_cb=None) -> int:
    """Extract QuickConvert payload into HKCU user folder and register HKCU context menus."""
    if sys.platform != "win32":
        raise RuntimeError("QuickConvert Setup is only supported on Windows.")

    payload_zip = _bundle_root() / "QuickConvert_payload.zip"
    if not payload_zip.is_file():
        raise FileNotFoundError(f"Bundled application payload not found: {payload_zip}")

    specs = _load_menu_specs()
    stage_dir = Path(tempfile.mkdtemp(prefix="quickconvert_setup_stage_"))
    try:
        if progress_cb:
            progress_cb(5.0, "Extracting QuickConvert application files...")

        with zipfile.ZipFile(payload_zip, "r") as zf:
            members = zf.infolist()
            total_members = max(1, len(members))
            for idx, info in enumerate(members, start=1):
                zf.extract(info, stage_dir)
                if progress_cb and (idx % 12 == 0 or idx == total_members):
                    pct = 5.0 + (idx / total_members) * 65.0
                    progress_cb(pct, f"Extracting files ({idx}/{total_members})...")

        if progress_cb:
            progress_cb(74.0, f"Installing to {install_dir}...")

        # Preserve user config.json if upgrading an existing install
        existing_cfg_bytes: bytes | None = None
        existing_cfg = install_dir / "config.json"
        if existing_cfg.is_file():
            try:
                existing_cfg_bytes = existing_cfg.read_bytes()
            except OSError:
                existing_cfg_bytes = None

        install_dir.mkdir(parents=True, exist_ok=True)
        # Copy staged files into install_dir cleanly
        for item in stage_dir.iterdir():
            dst = install_dir / item.name
            if item.is_dir():
                if dst.exists():
                    _rmtree_robust(dst)
                shutil.copytree(item, dst)
            else:
                shutil.copy2(item, dst)

        if existing_cfg_bytes is not None:
            try:
                (install_dir / "config.json").write_bytes(existing_cfg_bytes)
            except OSError:
                pass

        # Copy self into install_dir so Windows Installed Apps can invoke --uninstall
        uninstaller_exe = install_dir / "QuickConvertSetup.exe"
        if getattr(sys, "frozen", False):
            running_exe = Path(sys.executable).resolve()
            if running_exe != uninstaller_exe.resolve():
                shutil.copy2(running_exe, uninstaller_exe)

        if progress_cb:
            progress_cb(84.0, "Registering Windows Explorer context menus (HKCU)...")

        install_str = str(install_dir.resolve())
        installed_keys: list[str] = []
        total_specs = max(1, len(specs))

        for idx, spec in enumerate(specs, start=1):
            base_key = spec["base_key_path"]
            icon_val = spec["icon_path"].replace("{INSTALL_DIR}", install_str)
            _delete_reg_tree(winreg.HKEY_CURRENT_USER, base_key)

            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, base_key, 0, winreg.KEY_WRITE) as menu_key:
                winreg.SetValueEx(menu_key, "MUIVerb", 0, winreg.REG_SZ, APP_NAME)
                winreg.SetValueEx(menu_key, "SubCommands", 0, winreg.REG_SZ, "")
                winreg.SetValueEx(menu_key, "Icon", 0, winreg.REG_SZ, icon_val)

            for action in spec["actions"]:
                action_key_path = rf"{base_key}\shell\{action['subkey_name']}"
                cmd_str = action["command"].replace("{INSTALL_DIR}", install_str)
                with winreg.CreateKeyEx(
                    winreg.HKEY_CURRENT_USER,
                    action_key_path,
                    0,
                    winreg.KEY_WRITE,
                ) as action_key:
                    winreg.SetValueEx(action_key, "MUIVerb", 0, winreg.REG_SZ, action["label"])
                    if action.get("separator_before"):
                        winreg.SetValueEx(action_key, "CommandFlags", 0, winreg.REG_DWORD, ECF_SEPARATORBEFORE)

                with winreg.CreateKeyEx(
                    winreg.HKEY_CURRENT_USER,
                    rf"{action_key_path}\command",
                    0,
                    winreg.KEY_WRITE,
                ) as cmd_key:
                    winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, cmd_str)

            installed_keys.append(base_key)
            if progress_cb and (idx % 5 == 0 or idx == total_specs):
                pct = 84.0 + (idx / total_specs) * 14.0
                progress_cb(pct, f"Registering context menus ({idx}/{total_specs} formats)...")

        manifest_path = install_dir / "installed_registry_keys.json"
        manifest_path.write_text(
            json.dumps({"hive": "HKEY_CURRENT_USER", "keys": installed_keys}, indent=2),
            encoding="utf-8",
        )

        # Register per-user Windows Installed Apps entry (no Admin required)
        icon_file = install_dir / "assets" / "quickconvert.ico"
        uninstall_cmd = (
            f'"{uninstaller_exe}" --uninstall'
            if uninstaller_exe.is_file()
            else f'"{sys.executable}" "{Path(__file__).resolve()}" --uninstall'
        )
        with winreg.CreateKeyEx(
            winreg.HKEY_CURRENT_USER,
            UNINSTALL_REG_KEY,
            0,
            winreg.KEY_WRITE,
        ) as uninst_key:
            winreg.SetValueEx(uninst_key, "DisplayName", 0, winreg.REG_SZ, APP_NAME)
            winreg.SetValueEx(uninst_key, "DisplayVersion", 0, winreg.REG_SZ, APP_VERSION)
            winreg.SetValueEx(uninst_key, "Publisher", 0, winreg.REG_SZ, APP_PUBLISHER)
            winreg.SetValueEx(uninst_key, "InstallLocation", 0, winreg.REG_SZ, install_str)
            winreg.SetValueEx(uninst_key, "DisplayIcon", 0, winreg.REG_SZ, str(icon_file))
            winreg.SetValueEx(uninst_key, "UninstallString", 0, winreg.REG_SZ, uninstall_cmd)
            winreg.SetValueEx(uninst_key, "NoModify", 0, winreg.REG_DWORD, 1)
            winreg.SetValueEx(uninst_key, "NoRepair", 0, winreg.REG_DWORD, 1)

        if progress_cb:
            progress_cb(100.0, f"Installed QuickConvert for {len(installed_keys)} file formats.")
        return len(installed_keys)
    finally:
        _rmtree_robust(stage_dir)


def perform_uninstall(install_dir: Path, progress_cb=None) -> int:
    """Cleanly remove every QuickConvert HKCU Registry key and application files."""
    if sys.platform != "win32":
        raise RuntimeError("QuickConvert Uninstaller is only supported on Windows.")

    if progress_cb:
        progress_cb(15.0, "Collecting registered QuickConvert context menu keys...")

    keys_to_remove: set[str] = set()
    manifest_file = install_dir / "installed_registry_keys.json"
    if manifest_file.is_file():
        try:
            data = json.loads(manifest_file.read_text(encoding="utf-8"))
            for k in data.get("keys", []):
                if isinstance(k, str) and k.startswith(REG_ROOT_PREFIX + "\\"):
                    keys_to_remove.add(k)
        except (OSError, json.JSONDecodeError):
            pass

    try:
        for spec in _load_menu_specs():
            k = spec.get("base_key_path", "")
            if k.startswith(REG_ROOT_PREFIX + "\\"):
                keys_to_remove.add(k)
    except Exception:
        pass

    sorted_keys = sorted(keys_to_remove)
    removed = 0
    total = max(1, len(sorted_keys))
    for idx, key_path in enumerate(sorted_keys, start=1):
        if _delete_reg_tree(winreg.HKEY_CURRENT_USER, key_path):
            removed += 1
            shell_parent = key_path.rsplit("\\", 1)[0]
            ext_parent = shell_parent.rsplit("\\", 1)[0]
            if _prune_empty_reg_key(winreg.HKEY_CURRENT_USER, shell_parent):
                _prune_empty_reg_key(winreg.HKEY_CURRENT_USER, ext_parent)
        if progress_cb and (idx % 5 == 0 or idx == total):
            pct = 15.0 + (idx / total) * 55.0
            progress_cb(pct, f"Removing Registry entries ({idx}/{total})...")

    _delete_reg_tree(winreg.HKEY_CURRENT_USER, UNINSTALL_REG_KEY)

    if progress_cb:
        progress_cb(80.0, "Removing installed application files...")

    running_exe = Path(sys.executable).resolve() if getattr(sys, "frozen", False) else None
    running_inside_install = (
        running_exe is not None
        and install_dir.exists()
        and install_dir.resolve() in running_exe.parents
    )

    if install_dir.exists():
        if running_inside_install:
            for item in list(install_dir.iterdir()):
                if running_exe is not None and item.resolve() == running_exe:
                    continue
                if item.is_dir():
                    _rmtree_robust(item)
                else:
                    try:
                        item.unlink()
                    except OSError:
                        pass
            # Silently delete the remaining directory 2 seconds after this process exits
            subprocess.Popen(
                [
                    "cmd.exe",
                    "/c",
                    f'ping 127.0.0.1 -n 3 > nul & rd /s /q "{install_dir.resolve()}"',
                ],
                creationflags=NO_WINDOW_FLAGS | DETACHED_FLAGS,
            )
        else:
            _rmtree_robust(install_dir)

    if progress_cb:
        progress_cb(100.0, f"Removed {removed} context menu entries and uninstalled QuickConvert.")
    return removed


def _init_windows_ui() -> None:
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def _detect_theme() -> SetupTheme:
    if sys.platform != "win32":
        return LIGHT_THEME
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
            0,
            winreg.KEY_READ,
        ) as key:
            val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return LIGHT_THEME if int(val) != 0 else DARK_THEME
    except OSError:
        return LIGHT_THEME


def _apply_dwm(root: tk.Tk, theme: SetupTheme) -> None:
    if sys.platform != "win32":
        return
    root.update_idletasks()
    hwnd = ctypes.windll.user32.GetParent(root.winfo_id()) or root.winfo_id()
    cleaned = theme.bg.lstrip("#")
    r, g, b = int(cleaned[0:2], 16), int(cleaned[2:4], 16), int(cleaned[4:6], 16)
    colorref = (b << 16) | (g << 8) | r
    try:
        dwmapi = ctypes.windll.dwmapi
        dark_val = ctypes.c_int(0 if theme.is_light else 1)
        dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ctypes.byref(dark_val), ctypes.sizeof(dark_val))
        cap_val = ctypes.c_int(colorref)
        dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_CAPTION_COLOR, ctypes.byref(cap_val), ctypes.sizeof(cap_val))
    except Exception:
        pass


class _SetupButton(tk.Frame):
    def __init__(self, master, text, theme: SetupTheme, font_spec, command, primary=False, pad_x=14, pad_y=6):
        border_col = theme.accent if primary else theme.border
        super().__init__(master, bg=border_col, padx=1, pady=1, bd=0, highlightthickness=0)
        self._theme = theme
        self._primary = primary
        self._command = command
        self._enabled = True

        self._bg_idle = theme.accent if primary else theme.surface
        self._bg_hover = theme.accent_hover if primary else theme.btn_hover
        self._fg_idle = theme.accent_text if primary else theme.text_primary

        self._label = tk.Label(
            self,
            text=text,
            font=font_spec,
            bg=self._bg_idle,
            fg=self._fg_idle,
            padx=pad_x,
            pady=pad_y,
            cursor="hand2",
            bd=0,
            highlightthickness=0,
        )
        self._label.pack(fill="both", expand=True)
        self._label.bind("<Enter>", lambda _e: self._enabled and self._label.configure(bg=self._bg_hover))
        self._label.bind("<Leave>", lambda _e: self._enabled and self._label.configure(bg=self._bg_idle))
        self._label.bind("<ButtonRelease-1>", self._on_click)

    def _on_click(self, event: tk.Event) -> None:
        if not self._enabled:
            return
        if 0 <= event.x <= self._label.winfo_width() and 0 <= event.y <= self._label.winfo_height():
            self._command()

    def set_enabled(self, enabled: bool) -> None:
        self._enabled = enabled
        self._label.configure(cursor="hand2" if enabled else "arrow")


class SetupWindow:
    LOGICAL_WIDTH = 460
    MIN_LOGICAL_HEIGHT = 230

    def __init__(self, install_dir: Path, start_uninstall: bool = False) -> None:
        _init_windows_ui()
        self._install_dir = install_dir
        self._start_uninstall = start_uninstall
        self._theme = _detect_theme()
        self._queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._busy = False

        self._root = tk.Tk()
        self._root.title(f"{APP_NAME} Setup")
        self._root.resizable(False, False)
        self._root.configure(bg=self._theme.bg)

        try:
            self._scale = max(1.0, self._root.winfo_fpixels("1i") / 96.0)
        except Exception:
            self._scale = 1.0

        self._win_w = self._px(self.LOGICAL_WIDTH)
        self._apply_icon()
        self._select_fonts()
        self._build_ui()
        self._fit_window()
        _apply_dwm(self._root, self._theme)

        if self._start_uninstall:
            self._root.after(120, self._begin_uninstall)

    def _px(self, val: float) -> int:
        return max(1, int(round(val * self._scale)))

    def _apply_icon(self) -> None:
        icon_file = _bundle_root() / "assets" / "quickconvert.ico"
        if icon_file.is_file():
            try:
                self._root.iconbitmap(str(icon_file))
            except Exception:
                pass

    def _select_fonts(self) -> None:
        installed = set(tkfont.families(self._root))
        base = "Segoe UI Variable Text" if "Segoe UI Variable Text" in installed else "Segoe UI"
        semi = (
            "Segoe UI Variable Text Semibold"
            if "Segoe UI Variable Text Semibold" in installed
            else ("Segoe UI Semibold" if "Segoe UI Semibold" in installed else base)
        )
        self._font_title = (semi, 13) if semi != base else (base, 13, "bold")
        self._font_body = (base, 9)
        self._font_small = (base, 8)
        self._font_btn = (base, 9)

    def _build_ui(self) -> None:
        t = self._theme
        outer = tk.Frame(self._root, bg=t.bg, padx=self._px(14), pady=self._px(14))
        outer.pack(fill="both", expand=True)

        border = tk.Frame(outer, bg=t.border, padx=1, pady=1)
        border.pack(fill="both", expand=True)

        card = tk.Frame(border, bg=t.surface, padx=self._px(18), pady=self._px(16))
        card.pack(fill="both", expand=True)

        tk.Label(
            card,
            text=f"{APP_NAME} {APP_VERSION}",
            font=self._font_title,
            fg=t.text_primary,
            bg=t.surface,
            anchor="w",
        ).pack(fill="x")

        tk.Label(
            card,
            text="Offline Windows right-click converter for documents, PDFs, images, video, audio, and data. Installs for your user account without Administrator privileges.",
            font=self._font_body,
            fg=t.text_secondary,
            bg=t.surface,
            anchor="w",
            justify="left",
            wraplength=self._win_w - self._px(68),
        ).pack(fill="x", pady=(self._px(6), self._px(10)))

        # Install location box
        loc_border = tk.Frame(card, bg=t.border, padx=1, pady=1)
        loc_border.pack(fill="x", pady=(0, self._px(12)))
        loc_inner = tk.Frame(loc_border, bg=t.bg, padx=self._px(10), pady=self._px(6))
        loc_inner.pack(fill="x")
        tk.Label(
            loc_inner,
            text=f"Install folder: {self._install_dir}",
            font=self._font_small,
            fg=t.text_secondary,
            bg=t.bg,
            anchor="w",
            wraplength=self._win_w - self._px(92),
            justify="left",
        ).pack(fill="x")

        self._status_var = tk.StringVar(
            value="Ready to uninstall QuickConvert."
            if self._start_uninstall
            else ("QuickConvert is already installed. You can reinstall/update or uninstall."
                  if _is_installed(self._install_dir)
                  else "Ready to install QuickConvert.")
        )
        self._status_label = tk.Label(
            card,
            textvariable=self._status_var,
            font=self._font_body,
            fg=t.text_secondary,
            bg=t.surface,
            anchor="w",
            justify="left",
            wraplength=self._win_w - self._px(68),
        )
        self._status_label.pack(fill="x", pady=(0, self._px(8)))

        self._bar_h = self._px(5)
        self._bar = tk.Canvas(card, height=self._bar_h + 4, bg=t.surface, bd=0, highlightthickness=0)
        self._bar.pack(fill="x")
        self._bar_pct = 0.0
        self._bar_color = t.accent
        self._bar.bind("<Configure>", lambda _e: self._draw_bar())

        self._footer = tk.Frame(card, bg=t.surface)
        self._footer.pack(fill="x", pady=(self._px(14), self._px(2)))
        self._render_initial_buttons()

    def _draw_bar(self) -> None:
        self._bar.delete("all")
        w = self._bar.winfo_width()
        h = self._bar.winfo_height()
        if w <= self._bar_h:
            return
        cy = h / 2.0
        r = self._bar_h / 2.0
        self._bar.create_line(r, cy, w - r, cy, fill=self._theme.track, width=self._bar_h, capstyle=tk.ROUND)
        if self._bar_pct > 0:
            x1 = r + (w - 2 * r) * (self._bar_pct / 100.0)
            self._bar.create_line(r, cy, max(r + 0.5, x1), cy, fill=self._bar_color, width=self._bar_h, capstyle=tk.ROUND)

    def _clear_footer(self) -> None:
        for child in self._footer.winfo_children():
            child.destroy()

    def _render_initial_buttons(self) -> None:
        self._clear_footer()
        already = _is_installed(self._install_dir)
        _SetupButton(
            self._footer,
            "Close",
            self._theme,
            self._font_btn,
            self._root.destroy,
            primary=False,
            pad_x=self._px(12),
            pad_y=self._px(5),
        ).pack(side="right")

        if already:
            _SetupButton(
                self._footer,
                "Uninstall",
                self._theme,
                self._font_btn,
                self._begin_uninstall,
                primary=False,
                pad_x=self._px(12),
                pad_y=self._px(5),
            ).pack(side="right", padx=(0, self._px(8)))

        _SetupButton(
            self._footer,
            "Reinstall / Update" if already else "Install QuickConvert",
            self._theme,
            self._font_btn,
            self._begin_install,
            primary=True,
            pad_x=self._px(14),
            pad_y=self._px(5),
        ).pack(side="right", padx=(0, self._px(8)))

    def _render_done_button(self) -> None:
        self._clear_footer()
        _SetupButton(
            self._footer,
            "Finish",
            self._theme,
            self._font_btn,
            self._root.destroy,
            primary=True,
            pad_x=self._px(14),
            pad_y=self._px(5),
        ).pack(side="right")
        self._fit_window()

    def _fit_window(self) -> None:
        self._root.update_idletasks()
        rect = ctypes.wintypes.RECT()
        if sys.platform == "win32" and ctypes.windll.user32.SystemParametersInfoW(
            SPI_GETWORKAREA, 0, ctypes.byref(rect), 0
        ):
            wl, wt, wr, wb = rect.left, rect.top, rect.right, rect.bottom
        else:
            wl, wt, wr, wb = 0, 0, self._root.winfo_screenwidth(), self._root.winfo_screenheight() - 48
        win_h = max(self._px(self.MIN_LOGICAL_HEIGHT), self._root.winfo_reqheight() + self._px(4))
        x = wl + max(0, (wr - wl - self._win_w) // 2)
        y = wt + max(0, (wb - wt - win_h) // 2)
        self._root.geometry(f"{self._win_w}x{win_h}+{x}+{y}")

    def _begin_install(self) -> None:
        if self._busy:
            return
        self._busy = True
        self._clear_footer()
        self._bar_color = self._theme.accent
        self._status_label.configure(fg=self._theme.text_secondary)

        def _work() -> None:
            try:
                count = perform_install(
                    self._install_dir,
                    progress_cb=lambda p, m: self._queue.put(("progress", (p, m))),
                )
                self._queue.put(
                    (
                        "done",
                        f"Installation complete ({count} file formats registered). "
                        "Right-click any supported file -> Show more options -> QuickConvert.",
                    )
                )
            except Exception as exc:
                self._queue.put(("error", str(exc)))

        threading.Thread(target=_work, daemon=True).start()
        self._root.after(40, self._poll)

    def _begin_uninstall(self) -> None:
        if self._busy:
            return
        self._busy = True
        self._clear_footer()
        self._bar_color = self._theme.accent
        self._status_label.configure(fg=self._theme.text_secondary)

        def _work() -> None:
            try:
                removed = perform_uninstall(
                    self._install_dir,
                    progress_cb=lambda p, m: self._queue.put(("progress", (p, m))),
                )
                self._queue.put(
                    (
                        "done",
                        f"QuickConvert was cleanly uninstalled ({removed} Registry entries removed).",
                    )
                )
            except Exception as exc:
                self._queue.put(("error", str(exc)))

        threading.Thread(target=_work, daemon=True).start()
        self._root.after(40, self._poll)

    def _poll(self) -> None:
        try:
            while True:
                kind, data = self._queue.get_nowait()
                if kind == "progress":
                    pct, msg = data  # type: ignore[misc]
                    self._bar_pct = float(pct)
                    self._status_var.set(str(msg))
                    self._draw_bar()
                    self._fit_window()
                elif kind == "done":
                    self._busy = False
                    self._bar_pct = 100.0
                    self._bar_color = self._theme.success
                    self._status_label.configure(fg=self._theme.success)
                    self._status_var.set(str(data))
                    self._draw_bar()
                    self._render_done_button()
                    return
                elif kind == "error":
                    self._busy = False
                    self._bar_pct = 100.0
                    self._bar_color = self._theme.error
                    self._status_label.configure(fg=self._theme.error)
                    self._status_var.set(f"Setup failed: {data}")
                    self._draw_bar()
                    self._render_done_button()
                    return
        except queue.Empty:
            pass

        if self._busy:
            self._root.after(40, self._poll)

    def run(self) -> int:
        self._root.mainloop()
        return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="QuickConvert Windows Setup (No Admin required).")
    parser.add_argument("--uninstall", action="store_true", help="Uninstall QuickConvert and remove HKCU keys.")
    parser.add_argument("--silent", action="store_true", help="Run install/uninstall headlessly without UI.")
    parser.add_argument("--install-dir", default=None, help="Override target installation folder.")
    args = parser.parse_args(argv)

    install_dir = Path(args.install_dir).resolve() if args.install_dir else _default_install_dir()

    if args.silent:
        if args.uninstall:
            perform_uninstall(install_dir)
        else:
            perform_install(install_dir)
        return 0

    return SetupWindow(install_dir=install_dir, start_uninstall=args.uninstall).run()


if __name__ == "__main__":
    raise SystemExit(main())
