"""QuickConvert CLI & silent GUI entrypoint invoked by Windows context menus."""

from __future__ import annotations

import argparse
import os
import sys
import traceback
from datetime import datetime
from pathlib import Path

from quickconvert.config import load_config
from quickconvert.pipeline import build_default_graph, convert_file
from quickconvert.utils import CancellationToken, normalize_ext


def _log_error(message: str) -> None:
    """Append errors to a local log file when running silently under pythonw.exe."""
    local_app_data = os.environ.get("LOCALAPPDATA")
    log_dir = (
        Path(local_app_data) / "QuickConvert"
        if local_app_data
        else Path(__file__).resolve().parent
    )
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / "quickconvert.log"
        stamp = datetime.now().isoformat(timespec="seconds")
        with log_file.open("a", encoding="utf-8") as fh:
            fh.write(f"[{stamp}] {message}\n")
    except OSError:
        pass


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="QuickConvert",
        description="Convert documents, images, videos, and audio files silently.",
    )
    parser.add_argument(
        "--target",
        default=None,
        help="Target file extension (e.g. pdf, docx, png, jpg, webp, mp3, gif, webm).",
    )
    parser.add_argument(
        "--choose-target",
        action="store_true",
        help="Open the Fluent Format Picker dialog to choose from all valid target formats.",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Optional destination directory (defaults to each input file's folder).",
    )
    parser.add_argument(
        "--config",
        default=None,
        help="Optional path to config.json.",
    )
    parser.add_argument(
        "--no-ui",
        action="store_true",
        help="Run headlessly without launching the tkinter progress window.",
    )
    parser.add_argument(
        "--preview-ui",
        action="store_true",
        help="Open ProgressWindow and cycle through mock states for visual inspection.",
    )
    parser.add_argument(
        "files",
        nargs="*",
        help="One or more input file paths to convert.",
    )
    args = parser.parse_args(argv)
    if not args.preview_ui:
        if not args.files:
            parser.error("the following arguments are required: files")
        if not args.choose_target and not args.target:
            parser.error("the following arguments are required: --target (or --choose-target)")
    return args


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    config = load_config(args.config)
    graph = build_default_graph()

    if args.preview_ui:
        try:
            from quickconvert.ui.progress_window import ProgressWindow

            window = ProgressWindow(
                input_files=[],
                target_ext="png",
                config=config,
                graph=graph,
                preview_mode=True,
            )
            return window.run()
        except Exception as exc:
            _log_error(f"Preview UI startup error: {exc}\n{traceback.format_exc()}")
            if sys.stderr:
                print(f"ERROR: {exc}", file=sys.stderr)
            return 1

    input_files = [Path(f) for f in args.files]
    output_dir = Path(args.output_dir) if args.output_dir else None

    chosen_target = args.target
    if args.choose_target:
        try:
            from quickconvert.ui.format_picker import FormatPickerDialog

            picker = FormatPickerDialog(input_files=input_files, graph=graph)
            chosen_target = picker.run()
            if not chosen_target:
                return 0
        except Exception as exc:
            _log_error(f"Format picker error: {exc}\n{traceback.format_exc()}")
            if sys.stderr:
                print(f"ERROR: {exc}", file=sys.stderr)
            return 1

    target_ext = normalize_ext(str(chosen_target))

    if args.no_ui:
        token = CancellationToken()
        exit_code = 0
        for file_path in input_files:
            try:
                out = convert_file(
                    input_path=file_path,
                    target_ext=target_ext,
                    config=config,
                    graph=graph,
                    cancel_token=token,
                    output_dir=output_dir,
                )
                print(f"OK: {file_path} -> {out}")
            except Exception as exc:
                exit_code = 1
                print(f"ERROR: {file_path}: {exc}", file=sys.stderr)
                _log_error(f"{file_path} -> .{target_ext}: {exc}\n{traceback.format_exc()}")
        return exit_code

    try:
        from quickconvert.ui.progress_window import ProgressWindow

        window = ProgressWindow(
            input_files=input_files,
            target_ext=target_ext,
            config=config,
            graph=graph,
            output_dir=output_dir,
        )
        return window.run()
    except Exception as exc:
        _log_error(f"UI startup error: {exc}\n{traceback.format_exc()}")
        if sys.stderr:
            print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
