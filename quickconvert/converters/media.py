"""Universal offline video and audio converter using FFmpeg with container-matched codecs."""

from __future__ import annotations

import shutil
import subprocess
import time
from pathlib import Path

from quickconvert.config import AppConfig
from quickconvert.graph import (
    COST_DIRECT_NATIVE,
    COST_HIGH_FIDELITY,
    ConversionGraph,
    SubProgressCallback,
)
from quickconvert.utils import (
    NO_WINDOW_FLAGS,
    CancellationToken,
    ConversionError,
)


VIDEO_SOURCES: tuple[str, ...] = ("mp4", "mkv", "mov", "webm", "avi", "m4v", "wmv")
VIDEO_TARGETS: tuple[str, ...] = ("mp4", "mkv", "mov", "webm", "avi", "m4v")
AUDIO_SOURCES: tuple[str, ...] = ("mp3", "wav", "flac", "m4a", "aac", "ogg", "opus", "wma")
AUDIO_TARGETS: tuple[str, ...] = ("mp3", "wav", "flac", "m4a", "aac", "ogg", "opus")

H264_FASTSTART_TARGETS: set[str] = {"mp4", "mov", "m4v"}


def find_ffmpeg_executable(config: AppConfig) -> str:
    """Resolve FFmpeg binary from config, system PATH, or imageio-ffmpeg fallback."""
    if config.ffmpeg_path:
        cfg_path = Path(config.ffmpeg_path)
        if cfg_path.is_file():
            return str(cfg_path)

    on_path = shutil.which("ffmpeg") or shutil.which("ffmpeg.exe")
    if on_path:
        return on_path

    try:
        import imageio_ffmpeg  # type: ignore[import-untyped]

        bundled = imageio_ffmpeg.get_ffmpeg_exe()
        if bundled and Path(bundled).is_file():
            return bundled
    except Exception:
        pass

    raise ConversionError(
        "FFmpeg executable not found. Install FFmpeg on PATH, set 'ffmpeg_path' in config.json, "
        "or install 'imageio-ffmpeg'."
    )


def build_ffmpeg_command(
    ffmpeg_bin: str,
    src_path: Path,
    dst_path: Path,
    src_ext: str,
    dst_ext: str,
    config: AppConfig,
) -> list[str]:
    """Construct an FFmpeg argument list with container-matched codecs and scoped flags."""
    cmd = [ffmpeg_bin, "-y", "-hide_banner", "-loglevel", "error", "-i", str(src_path)]

    if dst_ext in H264_FASTSTART_TARGETS:
        cmd.extend(
            [
                "-vf",
                "scale=trunc(iw/2)*2:trunc(ih/2)*2",
                "-c:v",
                "libx264",
                "-preset",
                config.ffmpeg_preset,
                "-crf",
                str(config.ffmpeg_crf),
                "-pix_fmt",
                "yuv420p",
            ]
        )
        if src_ext == "gif":
            cmd.append("-an")
        else:
            cmd.extend(["-c:a", "aac", "-b:a", config.audio_bitrate])
        cmd.extend(["-movflags", "+faststart"])

    elif dst_ext == "mkv":
        cmd.extend(
            [
                "-vf",
                "scale=trunc(iw/2)*2:trunc(ih/2)*2",
                "-c:v",
                "libx264",
                "-preset",
                config.ffmpeg_preset,
                "-crf",
                str(config.ffmpeg_crf),
                "-pix_fmt",
                "yuv420p",
            ]
        )
        if src_ext == "gif":
            cmd.append("-an")
        else:
            cmd.extend(["-c:a", "aac", "-b:a", config.audio_bitrate])

    elif dst_ext == "avi":
        cmd.extend(
            [
                "-vf",
                "scale=trunc(iw/2)*2:trunc(ih/2)*2",
                "-c:v",
                "libx264",
                "-preset",
                config.ffmpeg_preset,
                "-crf",
                str(config.ffmpeg_crf),
                "-pix_fmt",
                "yuv420p",
            ]
        )
        if src_ext == "gif":
            cmd.append("-an")
        else:
            cmd.extend(["-c:a", "libmp3lame", "-b:a", config.audio_bitrate])

    elif dst_ext == "webm":
        # VP9 + Opus for WebM (omit H.264 -preset)
        cmd.extend(
            [
                "-c:v",
                "libvpx-vp9",
                "-crf",
                str(config.ffmpeg_crf),
                "-b:v",
                "0",
            ]
        )
        if src_ext == "gif":
            cmd.append("-an")
        else:
            cmd.extend(["-c:a", "libopus", "-b:a", config.audio_bitrate])

    elif dst_ext == "gif":
        # Two-step palette filtergraph; omit video preset/CRF flags and audio
        cmd.extend(
            [
                "-an",
                "-vf",
                "fps=12,scale='min(720,iw)':-1:flags=lanczos,split[s0][s1];"
                "[s0]palettegen[p];[s1][p]paletteuse",
            ]
        )

    elif dst_ext == "mp3":
        cmd.extend(["-vn", "-c:a", "libmp3lame", "-b:a", config.audio_bitrate])

    elif dst_ext == "wav":
        cmd.extend(["-vn", "-c:a", "pcm_s16le"])

    elif dst_ext == "flac":
        cmd.extend(["-vn", "-c:a", "flac"])

    elif dst_ext == "m4a":
        cmd.extend(["-vn", "-c:a", "aac", "-b:a", config.audio_bitrate, "-movflags", "+faststart"])

    elif dst_ext == "aac":
        cmd.extend(["-vn", "-c:a", "aac", "-b:a", config.audio_bitrate])

    elif dst_ext == "ogg":
        cmd.extend(["-vn", "-c:a", "libvorbis", "-b:a", config.audio_bitrate])

    elif dst_ext == "opus":
        cmd.extend(["-vn", "-c:a", "libopus", "-b:a", config.audio_bitrate])

    else:
        raise ConversionError(f"Unsupported media target format: .{dst_ext}")

    cmd.append(str(dst_path))
    return cmd


def _make_media_converter(src_ext: str, dst_ext: str):
    def _convert(
        src_path: Path,
        dst_path: Path,
        config: AppConfig,
        cancel_token: CancellationToken,
        sub_progress_cb: SubProgressCallback | None = None,
    ) -> None:
        cancel_token.raise_if_cancelled()
        ffmpeg_bin = find_ffmpeg_executable(config)
        cmd = build_ffmpeg_command(ffmpeg_bin, src_path, dst_path, src_ext, dst_ext, config)

        if sub_progress_cb:
            sub_progress_cb(None, f"Running FFmpeg ({src_ext.upper()} -> {dst_ext.upper()})...")

        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=NO_WINDOW_FLAGS,
        )
        try:
            while proc.poll() is None:
                if cancel_token.is_cancelled:
                    proc.terminate()
                    try:
                        proc.wait(timeout=5)
                    except subprocess.TimeoutExpired:
                        proc.kill()
                    cancel_token.raise_if_cancelled()
                time.sleep(0.1)

            _, stderr = proc.communicate()
            if proc.returncode != 0:
                err_text = stderr.decode("utf-8", errors="replace").strip()
                raise ConversionError(
                    f"FFmpeg failed ({src_ext.upper()} -> {dst_ext.upper()}): "
                    f"{err_text or f'exit code {proc.returncode}'}"
                )
        finally:
            if proc.poll() is None:
                proc.kill()

        if sub_progress_cb:
            sub_progress_cb(1.0, None)

    return _convert


def register_media_converters(graph: ConversionGraph) -> None:
    """Register direct 1-hop video and audio conversion edges (non-chaining)."""
    # Video -> Video, Video -> GIF, Video -> Audio
    for src in VIDEO_SOURCES:
        for dst in (*VIDEO_TARGETS, "gif"):
            if src == dst:
                continue
            graph.add_edge(
                src=src,
                dst=dst,
                converter=_make_media_converter(src, dst),
                label=f"{src.upper()} to {dst.upper()}",
                fidelity_cost=COST_HIGH_FIDELITY,
                allow_as_intermediate=False,
                category="Video",
            )
        for dst in AUDIO_TARGETS:
            graph.add_edge(
                src=src,
                dst=dst,
                converter=_make_media_converter(src, dst),
                label=f"{src.upper()} to {dst.upper()}",
                fidelity_cost=COST_DIRECT_NATIVE if dst in ("wav", "flac") else COST_HIGH_FIDELITY,
                allow_as_intermediate=False,
                category="Audio",
            )

    # GIF -> Video (no audio stream, never chains to audio)
    for dst in VIDEO_TARGETS:
        graph.add_edge(
            src="gif",
            dst=dst,
            converter=_make_media_converter("gif", dst),
            label=f"GIF to {dst.upper()}",
            fidelity_cost=COST_HIGH_FIDELITY,
            allow_as_intermediate=False,
            category="Video",
        )

    # Audio -> Audio
    for src in AUDIO_SOURCES:
        for dst in AUDIO_TARGETS:
            if src == dst:
                continue
            graph.add_edge(
                src=src,
                dst=dst,
                converter=_make_media_converter(src, dst),
                label=f"{src.upper()} to {dst.upper()}",
                fidelity_cost=COST_DIRECT_NATIVE if dst in ("wav", "flac") else COST_HIGH_FIDELITY,
                allow_as_intermediate=False,
                category="Audio",
            )
