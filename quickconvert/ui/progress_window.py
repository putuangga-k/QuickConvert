"""Native Windows 11 Fluent-style floating progress window for QuickConvert."""

from __future__ import annotations

import ctypes
import ctypes.wintypes
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import font as tkfont

from PIL import Image, ImageTk

from quickconvert.config import AppConfig
from quickconvert.graph import ConversionGraph
from quickconvert.paths import get_app_dir, get_resource_path
from quickconvert.pipeline import convert_file
from quickconvert.utils import (
    NO_WINDOW_FLAGS,
    CancellationToken,
    ConversionCancelledError,
    normalize_ext,
)

if sys.platform == "win32":
    import winreg


PROJECT_ROOT = get_app_dir()
APP_ICON_PATH = get_resource_path("assets/quickconvert.ico")
APP_USER_MODEL_ID = "QuickConvert.FileConverter.1"

DWMWA_USE_IMMERSIVE_DARK_MODE = 20
DWMWA_CAPTION_COLOR = 35
SPI_GETWORKAREA = 0x0030


@dataclass(frozen=True)
class FluentTheme:
    """Windows 11 Fluent color tokens synced to the OS theme."""

    is_light: bool
    bg: str
    surface: str
    border: str
    text_primary: str
    text_secondary: str
    accent: str
    success: str
    error: str
    track: str
    btn_hover: str
    btn_active: str
    tag_bg: str


LIGHT_THEME = FluentTheme(
    is_light=True,
    bg="#F3F3F3",
    surface="#FFFFFF",
    border="#E5E5E5",
    text_primary="#1B1B1B",
    text_secondary="#5D5D5D",
    accent="#0067C0",
    success="#0F7B0F",
    error="#C42B1C",
    track="#E5E5E5",
    btn_hover="#F6F6F6",
    btn_active="#EBEBEB",
    tag_bg="#F3F3F3",
)

DARK_THEME = FluentTheme(
    is_light=False,
    bg="#202020",
    surface="#2B2B2B",
    border="#3A3A3A",
    text_primary="#F3F3F3",
    text_secondary="#A6A6A6",
    accent="#60CDFF",
    success="#6CCB5F",
    error="#FF99A4",
    track="#3A3A3A",
    btn_hover="#353535",
    btn_active="#3E3E3E",
    tag_bg="#202020",
)


def _init_windows_process_environment() -> None:
    """Configure AppUserModelID and High-DPI awareness BEFORE creating tk.Tk()."""
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


def detect_windows_theme() -> FluentTheme:
    """Read HKCU AppsUseLightTheme via winreg and return matching Fluent tokens."""
    if sys.platform != "win32":
        return LIGHT_THEME

    reg_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ) as key:
            val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return LIGHT_THEME if int(val) != 0 else DARK_THEME
    except OSError:
        return LIGHT_THEME


def _hex_to_colorref(hex_color: str) -> int:
    """Convert `#RRGGBB` string into a Win32 COLORREF (`0x00BBGGRR`)."""
    cleaned = hex_color.lstrip("#")
    r = int(cleaned[0:2], 16)
    g = int(cleaned[2:4], 16)
    b = int(cleaned[4:6], 16)
    return (b << 16) | (g << 8) | r


def _apply_dwm_titlebar(root: tk.Tk, theme: FluentTheme) -> None:
    """Match native Windows 11 DWM titlebar theme and caption color."""
    if sys.platform != "win32":
        return

    root.update_idletasks()
    hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
    if not hwnd:
        hwnd = root.winfo_id()

    try:
        dwmapi = ctypes.windll.dwmapi
        dark_flag = ctypes.c_int(0 if theme.is_light else 1)
        dwmapi.DwmSetWindowAttribute(
            hwnd,
            DWMWA_USE_IMMERSIVE_DARK_MODE,
            ctypes.byref(dark_flag),
            ctypes.sizeof(dark_flag),
        )
        caption_color = ctypes.c_int(_hex_to_colorref(theme.bg))
        dwmapi.DwmSetWindowAttribute(
            hwnd,
            DWMWA_CAPTION_COLOR,
            ctypes.byref(caption_color),
            ctypes.sizeof(caption_color),
        )
    except Exception:
        pass


def _get_work_area(root: tk.Tk) -> tuple[int, int, int, int]:
    """Return usable screen work area (left, top, right, bottom) excluding the taskbar."""
    if sys.platform == "win32":
        rect = ctypes.wintypes.RECT()
        try:
            if ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(rect), 0):
                return rect.left, rect.top, rect.right, rect.bottom
        except Exception:
            pass
    return 0, 0, root.winfo_screenwidth(), root.winfo_screenheight() - 48


def _truncate_name(name: str, max_len: int = 36) -> str:
    if len(name) <= max_len:
        return name
    head = max_len // 2 - 2
    tail = max_len - head - 3
    return f"{name[:head]}...{name[-tail:]}"


def _clamp_error_text(text: str, max_chars: int = 210, max_lines: int = 4) -> str:
    """Keep error messages cleanly readable (up to ~4 wrapped lines) so footer controls never clip."""
    cleaned_lines = [re.sub(r"\s+", " ", line).strip() for line in text.splitlines() if line.strip()]
    if not cleaned_lines:
        return "An unknown error occurred."
    if len(cleaned_lines) > max_lines:
        cleaned_lines = [*cleaned_lines[: max_lines - 1], cleaned_lines[max_lines - 1] + "..."]
    joined = "\n".join(cleaned_lines)
    if len(joined) > max_chars:
        joined = joined[: max_chars - 3].rstrip() + "..."
    return joined


def _open_in_explorer(target_path: Path) -> None:
    """Open the output folder (or select the generated file) in Windows Explorer."""
    resolved = target_path.resolve()
    if sys.platform != "win32":
        return
    try:
        if resolved.is_file():
            subprocess.Popen(
                ["explorer.exe", f"/select,{resolved}"],
                creationflags=NO_WINDOW_FLAGS,
            )
        elif resolved.is_dir():
            os.startfile(str(resolved))  # type: ignore[attr-defined]
        elif resolved.parent.is_dir():
            os.startfile(str(resolved.parent))  # type: ignore[attr-defined]
    except OSError:
        pass


class _FluentButton(tk.Frame):
    """Flat, 1px hairline-bordered Windows 11 button with subtle hover feedback."""

    def __init__(
        self,
        master: tk.Misc,
        text: str,
        theme: FluentTheme,
        font_spec: tuple[str, int] | tuple[str, int, str],
        command,
        pad_x: int = 12,
        pad_y: int = 5,
    ) -> None:
        super().__init__(
            master,
            bg=theme.border,
            bd=0,
            highlightthickness=0,
            padx=1,
            pady=1,
        )
        self._theme = theme
        self._command = command
        self._enabled = True

        self._label = tk.Label(
            self,
            text=text,
            font=font_spec,
            bg=theme.surface,
            fg=theme.text_primary,
            padx=pad_x,
            pady=pad_y,
            cursor="hand2",
            bd=0,
            highlightthickness=0,
        )
        self._label.pack(fill="both", expand=True)

        self._label.bind("<Enter>", self._on_enter)
        self._label.bind("<Leave>", self._on_leave)
        self._label.bind("<ButtonPress-1>", self._on_press)
        self._label.bind("<ButtonRelease-1>", self._on_release)

    def _on_enter(self, _event: tk.Event) -> None:
        if self._enabled:
            self._label.configure(bg=self._theme.btn_hover)

    def _on_leave(self, _event: tk.Event) -> None:
        if self._enabled:
            self._label.configure(bg=self._theme.surface)

    def _on_press(self, _event: tk.Event) -> None:
        if self._enabled:
            self._label.configure(bg=self._theme.btn_active)

    def _on_release(self, event: tk.Event) -> None:
        if not self._enabled:
            return
        self._label.configure(bg=self._theme.btn_hover)
        w = self._label.winfo_width()
        h = self._label.winfo_height()
        if 0 <= event.x <= w and 0 <= event.y <= h and self._command:
            self._command()

    def set_enabled(self, enabled: bool) -> None:
        self._enabled = enabled
        if enabled:
            self._label.configure(fg=self._theme.text_primary, bg=self._theme.surface, cursor="hand2")
        else:
            self._label.configure(fg=self._theme.text_secondary, bg=self._theme.surface, cursor="arrow")


class _FluentProgressBar(tk.Canvas):
    """DPI-scaled rounded-end flat progress bar with smooth determinate and indeterminate modes."""

    def __init__(
        self,
        master: tk.Misc,
        theme: FluentTheme,
        bar_height: int,
        canvas_bg: str,
    ) -> None:
        self._bar_h = max(4, bar_height)
        canvas_h = self._bar_h + 4
        super().__init__(
            master,
            height=canvas_h,
            bg=canvas_bg,
            bd=0,
            highlightthickness=0,
        )
        self._theme = theme
        self._mode = "determinate"
        self._percent = 0.0
        self._fill_color = theme.accent
        self._phase = 0.0
        self._anim_id: str | None = None
        self.bind("<Configure>", lambda _e: self._redraw())

    def set_determinate(self, percent: float, color: str | None = None) -> None:
        self._stop_anim()
        self._mode = "determinate"
        self._percent = max(0.0, min(100.0, percent))
        self._fill_color = color or self._theme.accent
        self._redraw()

    def set_indeterminate(self, color: str | None = None) -> None:
        self._fill_color = color or self._theme.accent
        if self._mode != "indeterminate":
            self._mode = "indeterminate"
            self._phase = 0.0
            self._tick()
        else:
            self._redraw()

    def _stop_anim(self) -> None:
        if self._anim_id is not None:
            try:
                self.after_cancel(self._anim_id)
            except tk.TclError:
                pass
            self._anim_id = None

    def _tick(self) -> None:
        if self._mode != "indeterminate":
            return
        self._phase = (self._phase + 0.022) % 1.0
        self._redraw()
        self._anim_id = self.after(16, self._tick)

    def _redraw(self) -> None:
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w <= self._bar_h:
            return

        cy = h / 2.0
        cap_r = self._bar_h / 2.0
        x0 = cap_r
        x1 = w - cap_r

        # Track pill
        self.create_line(
            x0,
            cy,
            x1,
            cy,
            fill=self._theme.track,
            width=self._bar_h,
            capstyle=tk.ROUND,
        )

        span = x1 - x0
        if self._mode == "determinate":
            if self._percent <= 0.0:
                return
            fill_x = x0 + span * (self._percent / 100.0)
            self.create_line(
                x0,
                cy,
                max(x0 + 0.5, fill_x),
                cy,
                fill=self._fill_color,
                width=self._bar_h,
                capstyle=tk.ROUND,
            )
        else:
            pill_w = max(self._bar_h * 4.0, span * 0.28)
            travel = span + pill_w
            head = x0 - pill_w + travel * self._phase
            seg_start = max(x0, head)
            seg_end = min(x1, head + pill_w)
            if seg_end > seg_start:
                self.create_line(
                    seg_start,
                    cy,
                    seg_end,
                    cy,
                    fill=self._fill_color,
                    width=self._bar_h,
                    capstyle=tk.ROUND,
                )

    def stop(self) -> None:
        self._stop_anim()


class ProgressWindow:
    """Compact Windows 11 Fluent floating progress dialog with dynamic DPI-safe height fitting."""

    LOGICAL_WIDTH = 420
    MIN_LOGICAL_HEIGHT = 155

    def __init__(
        self,
        input_files: list[Path],
        target_ext: str,
        config: AppConfig,
        graph: ConversionGraph,
        output_dir: Path | None = None,
        preview_mode: bool = False,
    ) -> None:
        _init_windows_process_environment()

        self._input_files = input_files
        self._target_ext = normalize_ext(target_ext)
        self._config = config
        self._graph = graph
        self._output_dir = output_dir
        self._preview_mode = preview_mode

        self._cancel_token = CancellationToken()
        self._queue: queue.Queue[tuple[str, tuple[object, ...]]] = queue.Queue()
        self._finished = False
        self._exit_code = 0
        self._last_outputs: list[Path] = []
        self._auto_close_id: str | None = None
        self._auto_close_pending = False
        self._hovered = False
        self._icon_photo: ImageTk.PhotoImage | None = None

        self._theme = detect_windows_theme()

        self._root = tk.Tk()
        self._root.title("QuickConvert")
        self._root.resizable(False, False)
        self._root.attributes("-topmost", True)
        self._root.configure(bg=self._theme.bg)
        self._root.protocol("WM_DELETE_WINDOW", self._on_close_request)

        # Compute real display DPI scale factor (1.0, 1.25, 1.5, 2.0, etc.)
        try:
            self._dpi_scale = max(1.0, self._root.winfo_fpixels("1i") / 96.0)
        except Exception:
            self._dpi_scale = 1.0

        self._win_w = self._px(self.LOGICAL_WIDTH)

        self._apply_app_icon()
        self._select_fonts()
        self._build_widgets()
        self._fit_window()
        _apply_dwm_titlebar(self._root, self._theme)

        self._root.bind("<Enter>", self._on_window_enter)
        self._root.bind("<Leave>", self._on_window_leave)

    def _px(self, value: float) -> int:
        """Scale a logical 96-DPI pixel measurement to physical display pixels."""
        return max(1, int(round(value * self._dpi_scale)))

    def _apply_app_icon(self) -> None:
        if not APP_ICON_PATH.is_file():
            return
        try:
            self._root.iconbitmap(str(APP_ICON_PATH))
        except Exception:
            pass
        try:
            with Image.open(APP_ICON_PATH) as ico_img:
                icon_rgba = ico_img.convert("RGBA").resize((64, 64), Image.Resampling.LANCZOS)
            self._icon_photo = ImageTk.PhotoImage(icon_rgba, master=self._root)
            self._root.iconphoto(True, self._icon_photo)
        except Exception:
            pass

    def _select_fonts(self) -> None:
        installed = set(tkfont.families(self._root))
        base = "Segoe UI Variable Text" if "Segoe UI Variable Text" in installed else "Segoe UI"

        if "Segoe UI Variable Text Semibold" in installed:
            self._font_filename: tuple[str, int] | tuple[str, int, str] = ("Segoe UI Variable Text Semibold", 10)
        elif "Segoe UI Semibold" in installed:
            self._font_filename = ("Segoe UI Semibold", 10)
        else:
            self._font_filename = (base, 10, "bold")

        self._font_status: tuple[str, int] = (base, 9)
        self._font_tag: tuple[str, int, str] = (base, 8, "bold")
        self._font_button: tuple[str, int] = (base, 9)

    def _format_tag_text(self, file_path: Path) -> str:
        src = normalize_ext(file_path.suffix).upper() or "FILE"
        dst = self._target_ext.upper() or "OUT"
        return f"{src} -> {dst}"

    def _build_widgets(self) -> None:
        t = self._theme
        outer_pad = self._px(12)

        outer = tk.Frame(self._root, bg=t.bg, padx=outer_pad, pady=outer_pad)
        outer.pack(fill="both", expand=True)

        # Card surface with 1px hairline border
        card_border = tk.Frame(outer, bg=t.border, padx=1, pady=1)
        card_border.pack(fill="both", expand=True)

        card = tk.Frame(
            card_border,
            bg=t.surface,
            padx=self._px(15),
            pady=self._px(13),
        )
        card.pack(fill="both", expand=True)

        # Header row: filename (left, 10pt semibold) + compact uppercase format tag (right, 8pt bold)
        header_row = tk.Frame(card, bg=t.surface)
        header_row.pack(fill="x")

        initial_name = (
            "Quarterly_Report_2026.docx"
            if self._preview_mode
            else _truncate_name(self._input_files[0].name)
        )
        self._file_var = tk.StringVar(value=initial_name)
        tk.Label(
            header_row,
            textvariable=self._file_var,
            font=self._font_filename,
            fg=t.text_primary,
            bg=t.surface,
            anchor="w",
        ).pack(side="left", fill="x", expand=True)

        initial_tag = (
            "DOCX -> PNG"
            if self._preview_mode
            else self._format_tag_text(self._input_files[0])
        )
        self._tag_var = tk.StringVar(value=initial_tag)
        tag_border = tk.Frame(header_row, bg=t.border, padx=1, pady=1)
        tag_border.pack(side="right", padx=(self._px(8), 0))
        tk.Label(
            tag_border,
            textvariable=self._tag_var,
            font=self._font_tag,
            fg=t.text_secondary,
            bg=t.tag_bg,
            padx=self._px(6),
            pady=self._px(2),
        ).pack()

        # Step status line (9pt muted secondary text, wrapped cleanly within DPI-scaled width)
        self._step_var = tk.StringVar(value="Preparing conversion...")
        self._step_label = tk.Label(
            card,
            textvariable=self._step_var,
            font=self._font_status,
            fg=t.text_secondary,
            bg=t.surface,
            anchor="w",
            justify="left",
            wraplength=self._win_w - self._px(56),
        )
        self._step_label.pack(fill="x", pady=(self._px(8), self._px(10)))

        # Rounded-end flat progress bar on borderless tk.Canvas (5px logical thickness)
        self._bar = _FluentProgressBar(
            card,
            theme=t,
            bar_height=self._px(5),
            canvas_bg=t.surface,
        )
        self._bar.pack(fill="x")

        # Bottom action row (right-aligned flat hairline buttons + left status label)
        self._footer = tk.Frame(card, bg=t.surface)
        self._footer.pack(fill="x", pady=(self._px(14), self._px(2)))

        self._state_var = tk.StringVar(value="")
        self._state_label = tk.Label(
            self._footer,
            textvariable=self._state_var,
            font=self._font_status,
            fg=t.text_secondary,
            bg=t.surface,
            anchor="w",
        )
        self._state_label.pack(side="left", fill="x", expand=True, padx=(0, self._px(8)))

        self._btn_container = tk.Frame(self._footer, bg=t.surface)
        self._btn_container.pack(side="right")

        self._show_running_buttons(refit=False)

    def _fit_window(self) -> None:
        """Dynamically size window height to content and anchor cleanly above the taskbar."""
        try:
            self._root.update_idletasks()
            work_left, work_top, work_right, work_bottom = _get_work_area(self._root)
            margin = self._px(16)
            max_work_h = max(self._px(self.MIN_LOGICAL_HEIGHT), (work_bottom - work_top) - margin * 2)

            req_h = max(self._px(self.MIN_LOGICAL_HEIGHT), self._root.winfo_reqheight() + self._px(2))
            win_h = min(req_h, max_work_h)
            win_w = self._win_w

            pid_slot = os.getpid() % 6
            vertical_offset = pid_slot * self._px(34)

            x = max(work_left + margin, work_right - win_w - margin)
            y = max(work_top + margin, work_bottom - win_h - margin - vertical_offset)
            self._root.geometry(f"{win_w}x{win_h}+{x}+{y}")
        except tk.TclError:
            pass

    def _clear_buttons(self) -> None:
        for child in self._btn_container.winfo_children():
            child.destroy()

    def _show_running_buttons(self, refit: bool = True) -> None:
        self._clear_buttons()
        self._cancel_btn = _FluentButton(
            self._btn_container,
            text="Cancel",
            theme=self._theme,
            font_spec=self._font_button,
            command=self._on_cancel_click,
            pad_x=self._px(12),
            pad_y=self._px(5),
        )
        self._cancel_btn.pack(side="right")
        if refit:
            self._fit_window()

    def _show_completed_buttons(self, open_target: Path | None) -> None:
        self._clear_buttons()
        close_btn = _FluentButton(
            self._btn_container,
            text="Close",
            theme=self._theme,
            font_spec=self._font_button,
            command=self._close_window,
            pad_x=self._px(12),
            pad_y=self._px(5),
        )
        close_btn.pack(side="right")

        def _handle_open() -> None:
            if open_target is not None:
                _open_in_explorer(open_target)
            self._close_window()

        open_btn = _FluentButton(
            self._btn_container,
            text="Open Folder",
            theme=self._theme,
            font_spec=self._font_button,
            command=_handle_open,
            pad_x=self._px(12),
            pad_y=self._px(5),
        )
        open_btn.pack(side="right", padx=(0, self._px(8)))
        self._fit_window()

    def _show_close_only_button(self) -> None:
        self._clear_buttons()
        close_btn = _FluentButton(
            self._btn_container,
            text="Close",
            theme=self._theme,
            font_spec=self._font_button,
            command=self._close_window,
            pad_x=self._px(12),
            pad_y=self._px(5),
        )
        close_btn.pack(side="right")
        self._fit_window()

    def _schedule_auto_close(self, delay_ms: int) -> None:
        self._cancel_auto_close()
        self._auto_close_pending = True
        if not self._hovered:
            self._auto_close_id = self._root.after(max(400, delay_ms), self._close_window)

    def _cancel_auto_close(self) -> None:
        if self._auto_close_id is not None:
            try:
                self._root.after_cancel(self._auto_close_id)
            except tk.TclError:
                pass
            self._auto_close_id = None

    def _on_window_enter(self, _event: tk.Event) -> None:
        self._hovered = True
        if self._auto_close_pending:
            self._cancel_auto_close()

    def _on_window_leave(self, _event: tk.Event) -> None:
        try:
            px, py = self._root.winfo_pointerxy()
            rx = self._root.winfo_rootx()
            ry = self._root.winfo_rooty()
            rw = self._root.winfo_width()
            rh = self._root.winfo_height()
            if rx <= px <= rx + rw and ry <= py <= ry + rh:
                return
        except tk.TclError:
            return

        self._hovered = False
        if self._auto_close_pending and self._auto_close_id is None:
            self._auto_close_id = self._root.after(
                max(600, self._config.auto_close_delay_ms),
                self._close_window,
            )

    def _close_window(self) -> None:
        self._auto_close_pending = False
        self._cancel_auto_close()
        self._bar.stop()
        try:
            self._root.destroy()
        except tk.TclError:
            pass

    def _on_cancel_click(self) -> None:
        if self._finished or self._preview_mode:
            self._cancel_token.cancel()
            self._close_window()
            return
        self._cancel_token.cancel()
        self._step_var.set("Cancelling and cleaning up temporary files...")
        if hasattr(self, "_cancel_btn"):
            self._cancel_btn.set_enabled(False)
        self._fit_window()

    def _on_close_request(self) -> None:
        if self._finished or self._preview_mode:
            self._close_window()
        else:
            self._on_cancel_click()

    def _worker(self) -> None:
        outputs: list[Path] = []
        total_files = len(self._input_files)
        try:
            for file_idx, file_path in enumerate(self._input_files, start=1):
                self._cancel_token.raise_if_cancelled()
                prefix = f"[{file_idx}/{total_files}] " if total_files > 1 else ""
                self._queue.put(
                    (
                        "file",
                        (f"{prefix}{_truncate_name(file_path.name)}", self._format_tag_text(file_path)),
                    )
                )

                def _on_progress(
                    step_index: int,
                    total_steps: int,
                    message: str,
                    sub_progress: float | None,
                    _file_idx: int = file_idx,
                ) -> None:
                    if sub_progress is None:
                        self._queue.put(("progress", (message, None)))
                    else:
                        step_frac = ((step_index - 1) + sub_progress) / max(1, total_steps)
                        overall_pct = (((_file_idx - 1) + step_frac) / max(1, total_files)) * 100.0
                        self._queue.put(("progress", (message, overall_pct)))

                out = convert_file(
                    input_path=file_path,
                    target_ext=self._target_ext,
                    config=self._config,
                    graph=self._graph,
                    progress_cb=_on_progress,
                    cancel_token=self._cancel_token,
                    output_dir=self._output_dir,
                )
                outputs.append(out)

            self._queue.put(("done", (outputs,)))
        except ConversionCancelledError:
            self._queue.put(("cancelled", ()))
        except Exception as exc:
            self._queue.put(("error", (str(exc),)))

    def _poll_queue(self) -> None:
        needs_fit = False
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == "file":
                    name_text, tag_text = payload
                    self._file_var.set(str(name_text))
                    self._tag_var.set(str(tag_text))
                    needs_fit = True
                elif kind == "progress":
                    msg, pct = payload
                    msg_str = str(msg)
                    if self._step_var.get() != msg_str:
                        self._step_var.set(msg_str)
                        needs_fit = True
                    if pct is None:
                        self._bar.set_indeterminate(self._theme.accent)
                    else:
                        self._bar.set_determinate(float(pct), self._theme.accent)
                elif kind == "done":
                    outputs: list[Path] = payload[0]  # type: ignore[assignment]
                    self._last_outputs = outputs
                    self._finished = True
                    self._exit_code = 0
                    self._bar.set_determinate(100.0, self._theme.success)
                    if len(outputs) == 1:
                        label = outputs[0].name + ("/" if outputs[0].is_dir() else "")
                        self._step_var.set(f"Saved to {_truncate_name(label, 42)}")
                        open_target: Path | None = outputs[0]
                    else:
                        self._step_var.set(f"Converted {len(outputs)} files.")
                        open_target = outputs[0].parent if outputs else None
                    self._state_label.configure(fg=self._theme.success)
                    self._state_var.set("Completed")
                    self._show_completed_buttons(open_target)
                    self._schedule_auto_close(self._config.auto_close_delay_ms)
                    return
                elif kind == "cancelled":
                    self._finished = True
                    self._exit_code = 130
                    self._bar.set_determinate(0.0, self._theme.text_secondary)
                    self._step_var.set("Conversion cancelled.")
                    self._state_label.configure(fg=self._theme.text_secondary)
                    self._state_var.set("Cancelled")
                    self._show_close_only_button()
                    self._schedule_auto_close(900)
                    return
                elif kind == "error":
                    err_msg = _clamp_error_text(str(payload[0]))
                    self._finished = True
                    self._exit_code = 1
                    self._bar.set_determinate(100.0, self._theme.error)
                    self._step_var.set(err_msg)
                    self._state_label.configure(fg=self._theme.error)
                    self._state_var.set("Error")
                    self._show_close_only_button()
                    return
        except queue.Empty:
            pass

        if needs_fit:
            self._fit_window()

        if not self._finished:
            self._root.after(40, self._poll_queue)

    def _run_preview_sequence(self) -> None:
        """Cycle through realistic UI states (including multi-line error) for visual inspection."""
        self._file_var.set("Quarterly_Report_2026.docx")
        self._tag_var.set("DOCX -> PNG")
        self._state_label.configure(fg=self._theme.text_secondary)
        self._state_var.set("Preview 1/4: Indeterminate")
        self._step_var.set("Step 1/2: DOCX to PDF (Exporting via Microsoft Office...)")
        self._bar.set_indeterminate(self._theme.accent)
        self._show_running_buttons(refit=True)

        def _stage_determinate(pct: int = 0) -> None:
            if self._cancel_token.is_cancelled:
                self._close_window()
                return
            page = min(4, max(1, (pct // 25) + 1))
            self._state_var.set(f"Preview 2/4: Determinate ({pct}%)")
            new_step = f"Step 2/2: PDF to PNG (Per Page) (Page {page}/4)"
            if self._step_var.get() != new_step:
                self._step_var.set(new_step)
                self._fit_window()
            self._bar.set_determinate(float(pct), self._theme.accent)
            if pct < 100:
                self._root.after(70, lambda: _stage_determinate(min(100, pct + 4)))
            else:
                self._root.after(350, _stage_completed)

        def _stage_completed() -> None:
            if self._cancel_token.is_cancelled:
                self._close_window()
                return
            self._bar.set_determinate(100.0, self._theme.success)
            self._step_var.set("Saved to Quarterly_Report_2026_pages/")
            self._state_label.configure(fg=self._theme.success)
            self._state_var.set("Preview 3/4: Completed")
            self._show_completed_buttons(PROJECT_ROOT)
            self._root.after(2600, _stage_error)

        def _stage_error() -> None:
            if self._cancel_token.is_cancelled:
                self._close_window()
                return
            self._finished = True
            self._bar.set_determinate(100.0, self._theme.error)
            sample_long_err = (
                "Document conversion failed for 'Quarterly_Report_2026.docx'. "
                "Microsoft Office COM returned RPC_E_CALL_REJECTED after retries, "
                "and LibreOffice headless fallback could not locate soffice.exe. "
                "Original source file was preserved untouched and temporary files were removed."
            )
            self._step_var.set(_clamp_error_text(sample_long_err))
            self._state_label.configure(fg=self._theme.error)
            self._state_var.set("Preview 4/4: Error")
            self._show_close_only_button()

        self._root.after(2200, lambda: _stage_determinate(0))

    def run(self) -> int:
        """Launch either the conversion worker thread or the visual preview cycle."""
        if self._preview_mode:
            self._root.after(50, self._run_preview_sequence)
        else:
            worker = threading.Thread(target=self._worker, daemon=True)
            worker.start()
            self._root.after(40, self._poll_queue)
        self._root.mainloop()
        return self._exit_code
