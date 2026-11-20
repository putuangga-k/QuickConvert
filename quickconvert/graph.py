"""Weighted directed conversion graph (Dijkstra) with fidelity costs and submenu prioritization."""

from __future__ import annotations

import heapq
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from quickconvert.config import AppConfig
from quickconvert.utils import CancellationToken, normalize_ext


# Progress callback signature:
# (step_index: int, total_steps: int, message: str, sub_progress: float | None)
ProgressCallback = Callable[[int, int, str, float | None], None]

# Converter function signature:
# (src_path, dst_path, config, cancel_token, sub_progress_cb)
SubProgressCallback = Callable[[float | None, str | None], None]
ConverterFn = Callable[
    [Path, Path, AppConfig, CancellationToken, SubProgressCallback | None],
    None,
]

# Fidelity cost tiers
COST_DIRECT_NATIVE = 1      # Lossless or direct native export
COST_HIGH_FIDELITY = 2      # High-fidelity render / transcode
COST_APPROXIMATION = 5      # Layout reconstruction / lossy extraction

CATEGORY_ORDER: dict[str, int] = {
    "Document": 0,
    "Image": 1,
    "Data & Text": 2,
    "Video": 3,
    "Audio": 4,
    "Other": 5,
}

# Prioritized direct targets per source extension (capped to <= 10 + "More Formats...")
QUICK_TARGET_PRIORITIES: dict[str, tuple[str, ...]] = {
    # PDF
    "pdf": ("docx", "png", "jpg", "webp", "tiff", "pptx", "txt", "html", "md"),
    # Word / Rich Text Documents
    "docx": ("pdf", "png", "jpg", "webp", "tiff", "txt", "html", "rtf"),
    "doc": ("pdf", "docx", "png", "jpg", "txt", "html", "rtf"),
    "rtf": ("pdf", "docx", "txt", "html", "png", "jpg"),
    "odt": ("pdf", "docx", "txt", "html", "rtf", "png", "jpg"),
    # Presentations
    "pptx": ("pdf", "png", "jpg", "webp", "tiff"),
    "ppt": ("pdf", "pptx", "png", "jpg", "webp", "tiff"),
    "odp": ("pdf", "pptx", "png", "jpg", "webp", "tiff"),
    # Spreadsheets
    "xlsx": ("pdf", "csv", "tsv", "json", "png", "jpg", "webp", "tiff"),
    "xls": ("pdf", "xlsx", "csv", "tsv", "json", "png", "jpg"),
    "ods": ("pdf", "xlsx", "csv", "tsv", "json", "png", "jpg"),
    # Markup & Plain Text
    "txt": ("pdf", "docx", "html", "md"),
    "md": ("pdf", "docx", "html", "txt"),
    "html": ("pdf", "docx", "md", "txt"),
    "htm": ("pdf", "docx", "md", "txt"),
    # Structured / Tabular Data
    "csv": ("xlsx", "tsv", "json", "yaml"),
    "tsv": ("xlsx", "csv", "json", "yaml"),
    "json": ("yaml", "xlsx", "csv", "tsv"),
    "yaml": ("json", "xlsx", "csv", "tsv"),
    "yml": ("json", "yaml", "xlsx", "csv", "tsv"),
    # Images
    "png": ("jpg", "webp", "avif", "heic", "tiff", "bmp", "ico", "pdf"),
    "jpg": ("png", "webp", "avif", "heic", "tiff", "bmp", "ico", "pdf"),
    "jpeg": ("png", "jpg", "webp", "avif", "heic", "tiff", "bmp", "ico", "pdf"),
    "webp": ("png", "jpg", "avif", "heic", "tiff", "bmp", "ico", "pdf"),
    "heic": ("jpg", "png", "webp", "avif", "tiff", "bmp", "ico", "pdf"),
    "heif": ("jpg", "png", "webp", "avif", "heic", "tiff", "bmp", "ico", "pdf"),
    "avif": ("png", "jpg", "webp", "heic", "tiff", "bmp", "ico", "pdf"),
    "bmp": ("png", "jpg", "webp", "avif", "heic", "tiff", "ico", "pdf"),
    "ico": ("png", "jpg", "webp", "avif", "bmp", "tiff", "pdf"),
    "tiff": ("pdf", "png", "jpg", "webp", "avif", "heic", "bmp", "ico"),
    "tif": ("pdf", "png", "jpg", "webp", "avif", "heic", "tiff", "bmp", "ico"),
    # Video
    "mp4": ("mkv", "mov", "webm", "gif", "mp3", "wav", "flac", "m4a", "aac", "avi"),
    "mkv": ("mp4", "mov", "webm", "gif", "mp3", "wav", "flac", "m4a", "aac", "avi"),
    "mov": ("mp4", "mkv", "webm", "gif", "mp3", "wav", "flac", "m4a", "aac", "avi"),
    "webm": ("mp4", "mkv", "mov", "gif", "mp3", "wav", "flac", "m4a", "ogg", "opus"),
    "avi": ("mp4", "mkv", "mov", "webm", "gif", "mp3", "wav", "flac", "m4a", "aac"),
    "m4v": ("mp4", "mkv", "mov", "webm", "gif", "mp3", "wav", "flac", "m4a", "aac"),
    "wmv": ("mp4", "mkv", "mov", "webm", "gif", "mp3", "wav", "flac", "m4a", "aac"),
    "gif": ("mp4", "mkv", "mov", "webm", "avi", "m4v"),
    # Audio
    "mp3": ("wav", "flac", "m4a", "aac", "ogg", "opus"),
    "wav": ("mp3", "flac", "m4a", "aac", "ogg", "opus"),
    "flac": ("mp3", "wav", "m4a", "aac", "ogg", "opus"),
    "m4a": ("mp3", "wav", "flac", "aac", "ogg", "opus"),
    "aac": ("mp3", "wav", "flac", "m4a", "ogg", "opus"),
    "ogg": ("mp3", "wav", "flac", "m4a", "aac", "opus"),
    "opus": ("mp3", "wav", "flac", "m4a", "aac", "ogg"),
    "wma": ("mp3", "wav", "flac", "m4a", "aac", "ogg", "opus"),
}


@dataclass(frozen=True)
class ConversionEdge:
    """A single weighted directed conversion step from `src` to `dst`."""

    src: str
    dst: str
    converter: ConverterFn
    label: str
    fidelity_cost: int = COST_DIRECT_NATIVE
    outputs_directory: bool = False
    may_output_directory: bool = False
    directory_suffix: str = "_pages"
    allow_as_intermediate: bool = True
    menu_visible: bool = True
    category: str = "Other"
    menu_label_override: str | None = None


@dataclass(frozen=True)
class TargetOption:
    """A reachable conversion target exposed to callers, submenus, and the Format Picker."""

    target_ext: str
    menu_label: str
    short_label: str
    category: str
    outputs_directory: bool
    total_cost: int
    hops: int
    path: tuple[ConversionEdge, ...]


class ConversionGraph:
    """Weighted directed graph of file extensions (nodes) and converter functions (edges)."""

    def __init__(self) -> None:
        self._adj: dict[str, list[ConversionEdge]] = {}

    def add_edge(
        self,
        src: str,
        dst: str,
        converter: ConverterFn,
        label: str | None = None,
        fidelity_cost: int = COST_DIRECT_NATIVE,
        outputs_directory: bool = False,
        may_output_directory: bool = False,
        directory_suffix: str = "_pages",
        allow_as_intermediate: bool = True,
        menu_visible: bool = True,
        category: str = "Other",
        menu_label_override: str | None = None,
    ) -> None:
        src_norm = normalize_ext(src)
        dst_norm = normalize_ext(dst)
        if src_norm == dst_norm:
            return
        step_label = label or f"{src_norm.upper()} to {dst_norm.upper()}"
        can_chain = False if (outputs_directory or may_output_directory) else allow_as_intermediate
        edge = ConversionEdge(
            src=src_norm,
            dst=dst_norm,
            converter=converter,
            label=step_label,
            fidelity_cost=max(1, int(fidelity_cost)),
            outputs_directory=outputs_directory,
            may_output_directory=may_output_directory,
            directory_suffix=directory_suffix,
            allow_as_intermediate=can_chain,
            menu_visible=menu_visible,
            category=category,
            menu_label_override=menu_label_override,
        )
        self._adj.setdefault(src_norm, []).append(edge)

    def get_supported_sources(self) -> list[str]:
        """Return all registered source extensions in deterministic sorted order."""
        return sorted(self._adj.keys())

    def _dijkstra_all(self, src_norm: str) -> dict[str, tuple[int, int, list[ConversionEdge]]]:
        """Run weighted shortest-path search from `src_norm` respecting `allow_as_intermediate`.

        Returns mapping `dst_ext -> (total_cost, hops, path_edges)`.
        """
        if src_norm not in self._adj:
            return {}

        # Best settled path to each reachable target (final destination)
        best_target: dict[str, tuple[int, int, list[ConversionEdge]]] = {}
        # Best cost at which a node was expanded as an intermediate
        best_expand: dict[str, tuple[int, int]] = {src_norm: (0, 0)}

        counter = 0
        # Heap entries: (total_cost, hops, tie_breaker, current_node, path_edges)
        heap: list[tuple[int, int, int, str, list[ConversionEdge]]] = [(0, 0, counter, src_norm, [])]

        while heap:
            cost, hops, _, current, path = heapq.heappop(heap)
            prev_expand = best_expand.get(current)
            if prev_expand is not None and (cost, hops) > prev_expand:
                continue

            visited_nodes = {src_norm} | {e.dst for e in path}
            for edge in self._adj.get(current, []):
                dst = edge.dst
                if dst == src_norm or dst in visited_nodes:
                    # Block round-trip cycles
                    continue

                new_cost = cost + edge.fidelity_cost
                new_hops = hops + 1
                new_path = [*path, edge]

                existing = best_target.get(dst)
                if existing is None or (new_cost, new_hops) < (existing[0], existing[1]):
                    best_target[dst] = (new_cost, new_hops, new_path)

                if edge.allow_as_intermediate and not edge.outputs_directory and not edge.may_output_directory:
                    prev_dst_expand = best_expand.get(dst)
                    if prev_dst_expand is None or (new_cost, new_hops) < prev_dst_expand:
                        best_expand[dst] = (new_cost, new_hops)
                        counter += 1
                        heapq.heappush(heap, (new_cost, new_hops, counter, dst, new_path))

        return best_target

    def find_path(self, src: str, dst: str) -> list[ConversionEdge] | None:
        """Find the minimum-cost valid conversion path from `src` to `dst`."""
        src_norm = normalize_ext(src)
        dst_norm = normalize_ext(dst)
        if src_norm == dst_norm:
            return None
        all_paths = self._dijkstra_all(src_norm)
        found = all_paths.get(dst_norm)
        return list(found[2]) if found is not None else None

    def _build_target_option(
        self,
        dst_ext: str,
        total_cost: int,
        hops: int,
        path: list[ConversionEdge],
    ) -> TargetOption:
        last = path[-1]
        if last.menu_label_override:
            menu_label = last.menu_label_override
            short_label = menu_label.removeprefix("Convert to ").strip()
        elif last.outputs_directory:
            menu_label = f"Convert to {dst_ext.upper()} (Per Page)"
            short_label = f"{dst_ext.upper()} (Per Page)"
        else:
            menu_label = f"Convert to {dst_ext.upper()}"
            short_label = dst_ext.upper()

        return TargetOption(
            target_ext=dst_ext,
            menu_label=menu_label,
            short_label=short_label,
            category=last.category,
            outputs_directory=last.outputs_directory,
            total_cost=total_cost,
            hops=hops,
            path=tuple(path),
        )

    def get_valid_targets(
        self,
        src: str,
        include_hidden: bool = False,
    ) -> list[TargetOption]:
        """Enumerate all reachable valid target formats from `src` sorted cleanly."""
        src_norm = normalize_ext(src)
        all_paths = self._dijkstra_all(src_norm)
        priority_list = QUICK_TARGET_PRIORITIES.get(src_norm, ())
        priority_rank = {ext: idx for idx, ext in enumerate(priority_list)}

        options: list[TargetOption] = []
        for dst_ext, (total_cost, hops, path) in all_paths.items():
            last = path[-1]
            if not include_hidden and not last.menu_visible:
                continue
            options.append(self._build_target_option(dst_ext, total_cost, hops, path))

        options.sort(
            key=lambda opt: (
                priority_rank.get(opt.target_ext, 999),
                CATEGORY_ORDER.get(opt.category, 99),
                opt.total_cost,
                opt.hops,
                opt.target_ext,
            )
        )
        return options

    def get_quick_targets(self, src: str, max_items: int = 10) -> list[TargetOption]:
        """Return up to `max_items` prioritized direct submenu targets for Explorer."""
        src_norm = normalize_ext(src)
        all_valid = self.get_valid_targets(src_norm, include_hidden=False)
        by_ext = {opt.target_ext: opt for opt in all_valid}

        selected: list[TargetOption] = []
        seen: set[str] = set()

        for preferred in QUICK_TARGET_PRIORITIES.get(src_norm, ()):
            if preferred in by_ext and preferred not in seen:
                selected.append(by_ext[preferred])
                seen.add(preferred)
                if len(selected) >= max_items:
                    return selected

        for opt in all_valid:
            if opt.target_ext not in seen:
                selected.append(opt)
                seen.add(opt.target_ext)
                if len(selected) >= max_items:
                    break

        return selected
