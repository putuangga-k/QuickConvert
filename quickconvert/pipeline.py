"""End-to-end conversion pipeline with Dijkstra routing, staging, integrity validation, and atomic commit."""

from __future__ import annotations

import tempfile
from pathlib import Path

from quickconvert.config import AppConfig, load_config
from quickconvert.converters.data_text import register_data_text_converters
from quickconvert.converters.documents import register_document_converters
from quickconvert.converters.images import register_image_converters
from quickconvert.converters.media import register_media_converters
from quickconvert.converters.pdf_converter import register_pdf_converters
from quickconvert.graph import ConversionGraph, ProgressCallback
from quickconvert.utils import (
    CancellationToken,
    ConversionError,
    commit_directory_contents,
    commit_file_atomic,
    normalize_ext,
    reserve_unique_dir,
    reserve_unique_file,
    rmtree_robust,
)
from quickconvert.validator import validate_output


def build_default_graph() -> ConversionGraph:
    """Build and return the universal offline QuickConvert conversion graph."""
    graph = ConversionGraph()
    register_document_converters(graph)
    register_pdf_converters(graph)
    register_data_text_converters(graph)
    register_image_converters(graph)
    register_media_converters(graph)
    return graph


def convert_file(
    input_path: Path | str,
    target_ext: str,
    config: AppConfig | None = None,
    graph: ConversionGraph | None = None,
    progress_cb: ProgressCallback | None = None,
    cancel_token: CancellationToken | None = None,
    output_dir: Path | str | None = None,
) -> Path:
    """Convert `input_path` to `target_ext` using weighted Dijkstra shortest-path routing.

    All conversions stage outputs in a dedicated temp directory, pass mandatory
    integrity validation (`validate_output`) before committing onto an atomically
    reserved destination path via `os.replace()`, and guarantee temp cleanup in `finally`.
    """
    src_file = Path(input_path).resolve()
    if not src_file.is_file():
        raise ConversionError(f"Input file does not exist: {src_file}")

    src_ext = normalize_ext(src_file.suffix)
    dst_ext = normalize_ext(target_ext)
    if not src_ext:
        raise ConversionError(f"Input file has no extension: {src_file.name}")

    cfg = config if config is not None else load_config()
    conv_graph = graph if graph is not None else build_default_graph()
    token = cancel_token if cancel_token is not None else CancellationToken()

    path_edges = conv_graph.find_path(src_ext, dst_ext)
    if not path_edges:
        raise ConversionError(f"No conversion path available from '.{src_ext}' to '.{dst_ext}'.")

    dest_parent = Path(output_dir).resolve() if output_dir is not None else src_file.parent
    total_steps = len(path_edges)

    temp_dir = Path(tempfile.mkdtemp(prefix="quickconvert_"))
    reserved_dst: Path | None = None
    committed = False

    try:
        current_input = src_file
        staged_output: Path | None = None
        last_edge = path_edges[-1]

        for idx, edge in enumerate(path_edges, start=1):
            token.raise_if_cancelled()
            step_header = f"Step {idx}/{total_steps}: {edge.label}"

            if progress_cb:
                progress_cb(idx, total_steps, step_header, 0.0)

            def _sub_progress(
                frac: float | None,
                detail: str | None,
                _idx: int = idx,
                _header: str = step_header,
            ) -> None:
                if progress_cb:
                    msg = f"{_header} ({detail})" if detail else _header
                    progress_cb(_idx, total_steps, msg, frac)

            if edge.outputs_directory:
                step_out = temp_dir / f"step_{idx}_{edge.dst}{edge.directory_suffix}"
                step_out.mkdir(parents=True, exist_ok=True)
            else:
                step_out = temp_dir / f"step_{idx}_output.{edge.dst}"

            edge.converter(current_input, step_out, cfg, token, _sub_progress)
            token.raise_if_cancelled()

            if step_out.is_dir():
                if not any(step_out.iterdir()):
                    raise ConversionError(f"Step '{edge.label}' produced an empty output folder.")
            elif not step_out.is_file() or step_out.stat().st_size == 0:
                raise ConversionError(f"Step '{edge.label}' did not produce a valid output file.")

            current_input = step_out
            staged_output = step_out

            if progress_cb:
                progress_cb(idx, total_steps, step_header, 1.0)

        assert staged_output is not None

        # Mandatory post-conversion integrity check before touching destination
        if progress_cb:
            progress_cb(total_steps, total_steps, "Validating output integrity...", None)
        validate_output(staged_output, dst_ext, cfg)
        token.raise_if_cancelled()

        if staged_output.is_dir():
            dir_suffix = last_edge.directory_suffix or "_pages"
            reserved_dst = reserve_unique_dir(dest_parent, f"{src_file.stem}{dir_suffix}")
            commit_directory_contents(staged_output, reserved_dst)
        else:
            reserved_dst = reserve_unique_file(dest_parent, src_file.stem, dst_ext)
            commit_file_atomic(staged_output, reserved_dst)

        committed = True
        return reserved_dst
    finally:
        rmtree_robust(temp_dir)
        if not committed and reserved_dst is not None and reserved_dst.exists():
            if reserved_dst.is_dir():
                rmtree_robust(reserved_dst)
            else:
                try:
                    reserved_dst.unlink()
                except OSError:
                    pass
