"""Native Windows 11 Fluent Format Picker dialog for QuickConvert ('More Formats...')."""

from __future__ import annotations

import tkinter as tk
from collections import OrderedDict
from pathlib import Path
from tkinter import font as tkfont

from PIL import Image, ImageTk

from quickconvert.graph import CATEGORY_ORDER, ConversionGraph, TargetOption
from quickconvert.ui.progress_window import (
    APP_ICON_PATH,
    FluentTheme,
    _apply_dwm_titlebar,
    _FluentButton,
    _get_work_area,
    _init_windows_process_environment,
    _truncate_name,
    detect_windows_theme,
)
from quickconvert.utils import normalize_ext


def compute_common_targets(input_files: list[Path], graph: ConversionGraph) -> list[TargetOption]:
    """Return valid target options common to all selected files, ordered by category."""
    if not input_files:
        return []

    first_ext = normalize_ext(input_files[0].suffix)
    first_targets = graph.get_valid_targets(first_ext, include_hidden=False)
    if len(input_files) == 1:
        return first_targets

    common_exts = {opt.target_ext for opt in first_targets}
    for file_path in input_files[1:]:
        ext = normalize_ext(file_path.suffix)
        valid_exts = {opt.target_ext for opt in graph.get_valid_targets(ext, include_hidden=False)}
        common_exts &= valid_exts

    return [opt for opt in first_targets if opt.target_ext in common_exts]


class FormatPickerDialog:
    """Compact Windows 11 Fluent dialog displaying all valid targets grouped by category."""

    LOGICAL_WIDTH = 460
    MIN_LOGICAL_HEIGHT = 180

    def __init__(self, input_files: list[Path], graph: ConversionGraph) -> None:
        _init_windows_process_environment()

        self._input_files = input_files
        self._graph = graph
        self._options = compute_common_targets(input_files, graph)
        self._selected_target: str | None = None
        self._icon_photo: ImageTk.PhotoImage | None = None
        self._theme: FluentTheme = detect_windows_theme()

        self._root = tk.Tk()
        self._root.title("QuickConvert - Choose Format")
        self._root.resizable(False, False)
        self._root.attributes("-topmost", True)
        self._root.configure(bg=self._theme.bg)
        self._root.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self._root.bind("<Escape>", lambda _e: self._on_cancel())

        try:
            self._dpi_scale = max(1.0, self._root.winfo_fpixels("1i") / 96.0)
        except Exception:
            self._dpi_scale = 1.0

        self._win_w = self._px(self.LOGICAL_WIDTH)

        self._apply_icon()
        self._select_fonts()
        self._build_ui()
        self._fit_window()
        _apply_dwm_titlebar(self._root, self._theme)

    def _px(self, value: float) -> int:
        """Scale a logical 96-DPI pixel measurement to physical display pixels."""
        return max(1, int(round(value * self._dpi_scale)))

    def _apply_icon(self) -> None:
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
            self._font_title: tuple[str, int] | tuple[str, int, str] = ("Segoe UI Variable Text Semibold", 10)
        elif "Segoe UI Semibold" in installed:
            self._font_title = ("Segoe UI Semibold", 10)
        else:
            self._font_title = (base, 10, "bold")

        self._font_sub: tuple[str, int] = (base, 9)
        self._font_cat: tuple[str, int, str] = (base, 8, "bold")
        self._font_btn: tuple[str, int] = (base, 9)

    def _group_by_category(self) -> OrderedDict[str, list[TargetOption]]:
        grouped: dict[str, list[TargetOption]] = {}
        for opt in self._options:
            grouped.setdefault(opt.category, []).append(opt)
        ordered_keys = sorted(grouped.keys(), key=lambda c: (CATEGORY_ORDER.get(c, 99), c))
        return OrderedDict((k, grouped[k]) for k in ordered_keys)

    def _build_ui(self) -> None:
        t = self._theme
        pad = self._px(12)

        outer = tk.Frame(self._root, bg=t.bg, padx=pad, pady=pad)
        outer.pack(fill="both", expand=True)

        card_border = tk.Frame(outer, bg=t.border, padx=1, pady=1)
        card_border.pack(fill="both", expand=True)

        card = tk.Frame(card_border, bg=t.surface, padx=self._px(15), pady=self._px(13))
        card.pack(fill="both", expand=True)

        # Header
        if len(self._input_files) == 1:
            heading = _truncate_name(self._input_files[0].name, 42)
            src_ext = normalize_ext(self._input_files[0].suffix).upper() or "FILE"
            subtitle = f"Select target format for {src_ext} file"
        else:
            heading = f"{len(self._input_files)} files selected"
            subtitle = "Select a target format compatible with all selected files"

        tk.Label(
            card,
            text=heading,
            font=self._font_title,
            fg=t.text_primary,
            bg=t.surface,
            anchor="w",
        ).pack(fill="x")

        tk.Label(
            card,
            text=subtitle,
            font=self._font_sub,
            fg=t.text_secondary,
            bg=t.surface,
            anchor="w",
            wraplength=self._win_w - self._px(56),
        ).pack(fill="x", pady=(self._px(3), self._px(10)))

        # Categorized format grid
        groups = self._group_by_category()
        if not groups:
            tk.Label(
                card,
                text="No compatible conversion targets are available for the selected file(s).",
                font=self._font_sub,
                fg=t.error,
                bg=t.surface,
                anchor="w",
                justify="left",
                wraplength=self._win_w - self._px(56),
            ).pack(fill="x", pady=(self._px(6), self._px(12)))
        else:
            cols = 3
            for cat_name, opts in groups.items():
                tk.Label(
                    card,
                    text=cat_name.upper(),
                    font=self._font_cat,
                    fg=t.text_secondary,
                    bg=t.surface,
                    anchor="w",
                ).pack(fill="x", pady=(self._px(5), self._px(4)))

                grid_frame = tk.Frame(card, bg=t.surface)
                grid_frame.pack(fill="x", pady=(0, self._px(6)))
                for c in range(cols):
                    grid_frame.grid_columnconfigure(c, weight=1, uniform="fmt_col")

                for idx, opt in enumerate(opts):
                    r, c = divmod(idx, cols)
                    btn = _FluentButton(
                        grid_frame,
                        text=opt.short_label,
                        theme=t,
                        font_spec=self._font_btn,
                        command=lambda ext=opt.target_ext: self._on_select(ext),
                        pad_x=self._px(8),
                        pad_y=self._px(5),
                    )
                    btn.grid(
                        row=r,
                        column=c,
                        sticky="ew",
                        padx=(0 if c == 0 else self._px(4), 0 if c == cols - 1 else self._px(4)),
                        pady=self._px(3),
                    )

        # Footer
        footer = tk.Frame(card, bg=t.surface)
        footer.pack(fill="x", pady=(self._px(10), self._px(2)))

        _FluentButton(
            footer,
            text="Cancel",
            theme=t,
            font_spec=self._font_btn,
            command=self._on_cancel,
            pad_x=self._px(12),
            pad_y=self._px(5),
        ).pack(side="right")

    def _fit_window(self) -> None:
        """Dynamically size picker height to content and clamp cleanly inside the usable work area."""
        try:
            self._root.update_idletasks()
            work_left, work_top, work_right, work_bottom = _get_work_area(self._root)
            margin = self._px(16)
            max_w = max(self._win_w, (work_right - work_left) - margin * 2)
            max_h = max(self._px(self.MIN_LOGICAL_HEIGHT), (work_bottom - work_top) - margin * 2)

            win_w = min(max(self._win_w, self._root.winfo_reqwidth() + self._px(4)), max_w)
            win_h = min(max(self._px(self.MIN_LOGICAL_HEIGHT), self._root.winfo_reqheight() + self._px(4)), max_h)

            x = work_left + max(margin, (work_right - work_left - win_w) // 2)
            y = work_top + max(margin, (work_bottom - work_top - win_h) // 2)
            self._root.geometry(f"{win_w}x{win_h}+{x}+{y}")
        except tk.TclError:
            pass

    def _on_select(self, target_ext: str) -> None:
        self._selected_target = target_ext
        self._root.destroy()

    def _on_cancel(self) -> None:
        self._selected_target = None
        self._root.destroy()

    def run(self) -> str | None:
        """Show the modal picker and return the chosen target extension (or None)."""
        self._root.mainloop()
        return self._selected_target
