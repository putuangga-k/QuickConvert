"""Text, markup, spreadsheet, and structured data converters (100% offline)."""

from __future__ import annotations

import csv
import html as html_lib
import json
import re
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

import markdown as md_lib
import openpyxl
import yaml
from docx import Document

from quickconvert.config import AppConfig
from quickconvert.graph import (
    COST_DIRECT_NATIVE,
    COST_HIGH_FIDELITY,
    ConversionGraph,
    SubProgressCallback,
)
from quickconvert.utils import CancellationToken, ConversionError

try:
    import pymupdf as fitz  # type: ignore[import-untyped]
except ImportError:  # pragma: no cover
    import fitz  # type: ignore[import-untyped]


_SANITIZE_SHEET_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]+')


def _safe_sheet_filename(sheet_name: str, fallback_idx: int) -> str:
    cleaned = _SANITIZE_SHEET_RE.sub("_", sheet_name).strip(" .")
    return cleaned or f"sheet_{fallback_idx}"


# ---------------------------------------------------------------------------
# Text & Markup (.txt, .md, .html, .htm) -> PDF, DOCX, HTML, MD, TXT
# ---------------------------------------------------------------------------

class _SimpleHTMLToBlocks(HTMLParser):
    """Lightweight HTML parser converting headings, paragraphs, and list items to structured blocks."""

    def __init__(self) -> None:
        super().__init__()
        self.blocks: list[tuple[str, str]] = []
        self._current_tag: str = "p"
        self._buf: list[str] = []
        self._skip_depth = 0

    def _flush(self) -> None:
        text = "".join(self._buf).strip()
        if text:
            self.blocks.append((self._current_tag, re.sub(r"\s+", " ", text)))
        self._buf = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        t = tag.lower()
        if t in ("script", "style", "head"):
            self._skip_depth += 1
            return
        if self._skip_depth > 0:
            return
        if t in ("h1", "h2", "h3", "h4", "p", "li", "pre", "blockquote", "tr"):
            self._flush()
            self._current_tag = t
        elif t == "br":
            self._buf.append("\n")

    def handle_endtag(self, tag: str) -> None:
        t = tag.lower()
        if t in ("script", "style", "head") and self._skip_depth > 0:
            self._skip_depth -= 1
            return
        if self._skip_depth > 0:
            return
        if t in ("h1", "h2", "h3", "h4", "p", "li", "pre", "blockquote", "tr"):
            self._flush()
            self._current_tag = "p"

    def handle_data(self, data: str) -> None:
        if self._skip_depth == 0:
            self._buf.append(data)

    def finish(self) -> list[tuple[str, str]]:
        self._flush()
        return self.blocks


def _read_utf8_text(src_path: Path) -> str:
    try:
        return src_path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        return src_path.read_text(encoding="latin-1")


def _source_to_html_document(src_path: Path, src_ext: str) -> str:
    raw = _read_utf8_text(src_path)
    if src_ext in ("html", "htm"):
        return raw
    if src_ext == "md":
        body = md_lib.markdown(raw, extensions=["extra", "tables", "fenced_code", "nl2br"])
    else:
        escaped = html_lib.escape(raw)
        paragraphs = [
            f"<p>{line.replace(chr(10), '<br/>')}</p>"
            for line in escaped.split("\n\n")
            if line.strip()
        ]
        body = "\n".join(paragraphs) if paragraphs else f"<pre>{escaped}</pre>"

    return (
        "<!DOCTYPE html>\n"
        '<html lang="en"><head><meta charset="utf-8">'
        f"<title>{html_lib.escape(src_path.stem)}</title>"
        "<style>"
        "body{font-family:sans-serif;font-size:11pt;line-height:1.5;color:#1b1b1b;}"
        "h1{font-size:18pt;margin:0.8em 0 0.4em;}h2{font-size:14pt;margin:0.7em 0 0.35em;}"
        "h3{font-size:12pt;}pre,code{font-family:monospace;background:#f3f3f3;padding:2px 4px;}"
        "table{border-collapse:collapse;width:100%;margin:10px 0;}"
        "th,td{border:1px solid #d0d0d0;padding:6px 8px;text-align:left;}"
        "</style></head><body>\n"
        f"{body}\n</body></html>\n"
    )


def _html_to_pdf_story(html_content: str, dst_pdf: Path) -> None:
    """Render HTML cleanly into a multi-page PDF using PyMuPDF Story (100% offline)."""
    mediabox = fitz.paper_rect("letter")
    where = mediabox + (54, 54, -54, -54)
    story = fitz.Story(html=html_content)
    writer = fitz.DocumentWriter(str(dst_pdf))
    try:
        more = True
        while more:
            dev = writer.begin_page(mediabox)
            more, _ = story.place(where)
            story.draw(dev)
            writer.end_page()
    finally:
        writer.close()


def _make_markup_converter(src_ext: str, dst_ext: str):
    def _convert(
        src_path: Path,
        dst_path: Path,
        config: AppConfig,
        cancel_token: CancellationToken,
        sub_progress_cb: SubProgressCallback | None = None,
    ) -> None:
        cancel_token.raise_if_cancelled()

        if dst_ext == "pdf":
            html_doc = _source_to_html_document(src_path, src_ext)
            _html_to_pdf_story(html_doc, dst_path)

        elif dst_ext == "html":
            html_doc = _source_to_html_document(src_path, src_ext)
            dst_path.write_text(html_doc, encoding="utf-8")

        elif dst_ext == "txt":
            if src_ext in ("html", "htm"):
                parser = _SimpleHTMLToBlocks()
                parser.feed(_read_utf8_text(src_path))
                blocks = parser.finish()
                text = "\n\n".join(text for _, text in blocks).strip()
            elif src_ext == "md":
                html_doc = _source_to_html_document(src_path, "md")
                parser = _SimpleHTMLToBlocks()
                parser.feed(html_doc)
                blocks = parser.finish()
                text = "\n\n".join(text for _, text in blocks).strip()
            else:
                text = _read_utf8_text(src_path).strip()
            dst_path.write_text((text or " ") + "\n", encoding="utf-8")

        elif dst_ext == "md":
            if src_ext in ("html", "htm"):
                parser = _SimpleHTMLToBlocks()
                parser.feed(_read_utf8_text(src_path))
                lines: list[str] = []
                for tag, content in parser.finish():
                    if tag == "h1":
                        lines.append(f"# {content}")
                    elif tag == "h2":
                        lines.append(f"## {content}")
                    elif tag == "h3":
                        lines.append(f"### {content}")
                    elif tag == "h4":
                        lines.append(f"#### {content}")
                    elif tag == "li":
                        lines.append(f"- {content}")
                    elif tag == "blockquote":
                        lines.append(f"> {content}")
                    else:
                        lines.append(content)
                dst_path.write_text("\n\n".join(lines).strip() + "\n", encoding="utf-8")
            else:
                dst_path.write_text(_read_utf8_text(src_path).strip() + "\n", encoding="utf-8")

        elif dst_ext == "docx":
            doc = Document()
            if src_ext == "txt":
                for para in _read_utf8_text(src_path).splitlines():
                    doc.add_paragraph(para)
            else:
                html_doc = _source_to_html_document(src_path, src_ext)
                parser = _SimpleHTMLToBlocks()
                parser.feed(html_doc)
                blocks = parser.finish()
                if not blocks:
                    doc.add_paragraph(_read_utf8_text(src_path))
                for tag, content in blocks:
                    if tag in ("h1", "h2", "h3", "h4"):
                        doc.add_heading(content, level=int(tag[1]))
                    elif tag == "li":
                        doc.add_paragraph(content, style="List Bullet")
                    else:
                        doc.add_paragraph(content)
            doc.save(str(dst_path))
        else:
            raise ConversionError(f"Unsupported markup conversion: .{src_ext} -> .{dst_ext}")

        if sub_progress_cb:
            sub_progress_cb(1.0, None)

    return _convert


# ---------------------------------------------------------------------------
# Tabular & Structured Data (.csv, .tsv, .json, .yaml, .yml, .xlsx)
# ---------------------------------------------------------------------------

def _normalize_cell(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    return json.dumps(value, ensure_ascii=False, default=str)


def _read_delimited_rows(src_path: Path, delimiter: str) -> list[list[str]]:
    with src_path.open("r", encoding="utf-8-sig", newline="") as fh:
        rows = [list(row) for row in csv.reader(fh, delimiter=delimiter)]
    if not rows:
        raise ConversionError(f"Tabular file '{src_path.name}' is empty.")
    return rows


def _rows_to_records(rows: list[list[Any]]) -> list[dict[str, Any]]:
    if not rows:
        return []
    raw_headers = [str(h).strip() if h is not None and str(h).strip() else f"col_{i + 1}" for i, h in enumerate(rows[0])]
    # Deduplicate headers cleanly
    seen: dict[str, int] = {}
    headers: list[str] = []
    for h in raw_headers:
        count = seen.get(h, 0)
        seen[h] = count + 1
        headers.append(h if count == 0 else f"{h}_{count + 1}")

    records: list[dict[str, Any]] = []
    for row in rows[1:]:
        padded = list(row) + [""] * max(0, len(headers) - len(row))
        records.append({headers[i]: padded[i] for i in range(len(headers))})
    return records


def _json_obj_to_sheets(data: Any) -> dict[str, list[list[Any]]]:
    """Validate and convert structured JSON/YAML data into rectangular sheet tables."""
    if isinstance(data, list) and data:
        return {"Sheet1": _list_to_rows(data)}

    if isinstance(data, dict) and data:
        # Multi-sheet format: {"SheetA": [...], "SheetB": [...]}
        if all(isinstance(v, list) and len(v) > 0 for v in data.values()):
            return {str(name): _list_to_rows(rows) for name, rows in data.items()}
        # Single record dict -> 1-row table
        if all(not isinstance(v, (dict, list)) for v in data.values()):
            headers = list(data.keys())
            return {"Sheet1": [headers, [_normalize_cell(data[h]) for h in headers]]}

    raise ConversionError(
        "Structured data is not tabular (expected a non-empty list of objects/rows, "
        "a flat object, or a mapping of sheet names to row lists)."
    )


def _list_to_rows(items: list[Any]) -> list[list[Any]]:
    if all(isinstance(item, dict) for item in items):
        headers: list[str] = []
        seen: set[str] = set()
        for item in items:
            for key in item.keys():
                skey = str(key)
                if skey not in seen:
                    seen.add(skey)
                    headers.append(skey)
        if not headers:
            raise ConversionError("JSON objects contain no keys for tabular export.")
        rows: list[list[Any]] = [headers]
        for item in items:
            rows.append([_normalize_cell(item.get(h, "")) for h in headers])
        return rows

    if all(isinstance(item, list) for item in items):
        max_len = max((len(r) for r in items), default=0)
        if max_len == 0:
            raise ConversionError("Rows are empty.")
        return [[_normalize_cell(c) for c in r] + [""] * (max_len - len(r)) for r in items]

    raise ConversionError("List elements must be uniform objects or row lists for tabular export.")


def _write_delimited(path: Path, rows: list[list[Any]], delimiter: str) -> None:
    with path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh, delimiter=delimiter)
        for row in rows:
            writer.writerow(["" if c is None else c for c in row])


def _read_xlsx_sheets(src_path: Path) -> dict[str, list[list[Any]]]:
    wb = openpyxl.load_workbook(str(src_path), read_only=True, data_only=True)
    try:
        sheets: dict[str, list[list[Any]]] = {}
        for name in wb.sheetnames:
            ws = wb[name]
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            # Strip trailing completely empty rows
            while rows and all(c is None or c == "" for c in rows[-1]):
                rows.pop()
            if rows:
                sheets[name] = rows
        if not sheets:
            raise ConversionError(f"Workbook '{src_path.name}' contains no data.")
        return sheets
    finally:
        wb.close()


def _write_xlsx_sheets(dst_path: Path, sheets: dict[str, list[list[Any]]]) -> None:
    wb = openpyxl.Workbook()
    try:
        first = True
        used_titles: set[str] = set()
        for raw_title, rows in sheets.items():
            title = re.sub(r"[\[\]:*?/\\]", "_", raw_title)[:31] or "Sheet"
            base = title
            idx = 1
            while title in used_titles:
                suffix = f"_{idx}"
                title = f"{base[:31 - len(suffix)]}{suffix}"
                idx += 1
            used_titles.add(title)

            if first:
                ws = wb.active
                ws.title = title
                first = False
            else:
                ws = wb.create_sheet(title=title)

            for row in rows:
                ws.append([_normalize_cell(c) for c in row])
        wb.save(str(dst_path))
    finally:
        wb.close()


def _load_structured(src_path: Path, src_ext: str) -> Any:
    text = _read_utf8_text(src_path)
    if src_ext == "json":
        try:
            return json.loads(text)
        except json.JSONDecodeError as exc:
            raise ConversionError(f"Invalid JSON input '{src_path.name}': {exc}") from exc
    try:
        parsed = yaml.safe_load(text)
    except yaml.YAMLError as exc:
        raise ConversionError(f"Invalid YAML input '{src_path.name}': {exc}") from exc
    if parsed is None:
        raise ConversionError(f"YAML input '{src_path.name}' is empty.")
    return parsed


def _make_data_converter(src_ext: str, dst_ext: str):
    def _convert(
        src_path: Path,
        dst_path: Path,
        config: AppConfig,
        cancel_token: CancellationToken,
        sub_progress_cb: SubProgressCallback | None = None,
    ) -> None:
        cancel_token.raise_if_cancelled()

        # JSON <-> YAML/YML
        if src_ext in ("json", "yaml", "yml") and dst_ext in ("json", "yaml", "yml"):
            data = _load_structured(src_path, src_ext)
            if dst_ext == "json":
                dst_path.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
            else:
                dst_path.write_text(
                    yaml.safe_dump(data, sort_keys=False, allow_unicode=True),
                    encoding="utf-8",
                )
            if sub_progress_cb:
                sub_progress_cb(1.0, None)
            return

        # XLSX -> CSV / TSV (single sheet => file; multi-sheet => directory so zero sheets are lost)
        if src_ext == "xlsx" and dst_ext in ("csv", "tsv"):
            delimiter = "," if dst_ext == "csv" else "\t"
            sheets = _read_xlsx_sheets(src_path)
            if len(sheets) == 1:
                only_rows = next(iter(sheets.values()))
                _write_delimited(dst_path, only_rows, delimiter)
            else:
                dst_path.mkdir(parents=True, exist_ok=True)
                used_names: set[str] = set()
                for idx, (sheet_name, rows) in enumerate(sheets.items(), start=1):
                    cancel_token.raise_if_cancelled()
                    safe = _safe_sheet_filename(sheet_name, idx)
                    candidate = safe
                    dup = 1
                    while candidate.lower() in used_names:
                        candidate = f"{safe}_{dup}"
                        dup += 1
                    used_names.add(candidate.lower())
                    _write_delimited(dst_path / f"{candidate}.{dst_ext}", rows, delimiter)
            if sub_progress_cb:
                sub_progress_cb(1.0, None)
            return

        # XLSX -> JSON
        if src_ext == "xlsx" and dst_ext == "json":
            sheets = _read_xlsx_sheets(src_path)
            if len(sheets) == 1:
                payload: Any = _rows_to_records(next(iter(sheets.values())))
            else:
                payload = {name: _rows_to_records(rows) for name, rows in sheets.items()}
            dst_path.write_text(
                json.dumps(payload, indent=2, ensure_ascii=False, default=str) + "\n",
                encoding="utf-8",
            )
            if sub_progress_cb:
                sub_progress_cb(1.0, None)
            return

        # CSV / TSV -> CSV / TSV / JSON / XLSX
        if src_ext in ("csv", "tsv"):
            src_delim = "," if src_ext == "csv" else "\t"
            rows = _read_delimited_rows(src_path, src_delim)
            if dst_ext in ("csv", "tsv"):
                _write_delimited(dst_path, rows, "," if dst_ext == "csv" else "\t")
            elif dst_ext == "json":
                dst_path.write_text(
                    json.dumps(_rows_to_records(rows), indent=2, ensure_ascii=False) + "\n",
                    encoding="utf-8",
                )
            elif dst_ext == "xlsx":
                _write_xlsx_sheets(dst_path, {"Sheet1": rows})
            else:
                raise ConversionError(f"Unsupported tabular target: .{dst_ext}")
            if sub_progress_cb:
                sub_progress_cb(1.0, None)
            return

        # JSON / YAML -> XLSX / CSV / TSV
        if src_ext in ("json", "yaml", "yml") and dst_ext in ("xlsx", "csv", "tsv"):
            data = _load_structured(src_path, src_ext)
            sheets = _json_obj_to_sheets(data)
            if dst_ext == "xlsx":
                _write_xlsx_sheets(dst_path, sheets)
            else:
                delimiter = "," if dst_ext == "csv" else "\t"
                if len(sheets) == 1:
                    _write_delimited(dst_path, next(iter(sheets.values())), delimiter)
                else:
                    dst_path.mkdir(parents=True, exist_ok=True)
                    for idx, (sheet_name, rows) in enumerate(sheets.items(), start=1):
                        safe = _safe_sheet_filename(sheet_name, idx)
                        _write_delimited(dst_path / f"{safe}.{dst_ext}", rows, delimiter)
            if sub_progress_cb:
                sub_progress_cb(1.0, None)
            return

        raise ConversionError(f"Unsupported data conversion: .{src_ext} -> .{dst_ext}")

    return _convert


def register_data_text_converters(graph: ConversionGraph) -> None:
    """Register text, markup, spreadsheet-data, and structured data edges."""
    # Markup & Text (.txt, .md, .html, .htm)
    markup_sources = ("txt", "md", "html", "htm")
    markup_targets: list[tuple[str, int, bool, str]] = [
        ("pdf", COST_HIGH_FIDELITY, True, "Document"),
        ("docx", COST_HIGH_FIDELITY, False, "Document"),
        ("html", COST_DIRECT_NATIVE, False, "Data & Text"),
        ("md", COST_HIGH_FIDELITY, False, "Data & Text"),
        ("txt", COST_HIGH_FIDELITY, False, "Data & Text"),
    ]
    for src in markup_sources:
        for dst, cost, can_chain, cat in markup_targets:
            if src == dst or (src == "htm" and dst == "html"):
                continue
            graph.add_edge(
                src=src,
                dst=dst,
                converter=_make_markup_converter(src, dst),
                label=f"{src.upper()} to {dst.upper()}",
                fidelity_cost=cost,
                allow_as_intermediate=can_chain,
                category=cat,
            )

    # XLSX -> CSV / TSV / JSON via openpyxl (multi-sheet CSV/TSV exports to <stem>_sheets/)
    for dst in ("csv", "tsv"):
        graph.add_edge(
            src="xlsx",
            dst=dst,
            converter=_make_data_converter("xlsx", dst),
            label=f"XLSX to {dst.upper()}",
            fidelity_cost=COST_DIRECT_NATIVE,
            may_output_directory=True,
            directory_suffix="_sheets",
            allow_as_intermediate=False,
            category="Data & Text",
        )
    graph.add_edge(
        src="xlsx",
        dst="json",
        converter=_make_data_converter("xlsx", "json"),
        label="XLSX to JSON",
        fidelity_cost=COST_DIRECT_NATIVE,
        allow_as_intermediate=False,
        category="Data & Text",
    )

    # CSV <-> TSV <-> JSON <-> XLSX
    for src in ("csv", "tsv"):
        for dst in ("csv", "tsv", "json", "xlsx"):
            if src == dst:
                continue
            graph.add_edge(
                src=src,
                dst=dst,
                converter=_make_data_converter(src, dst),
                label=f"{src.upper()} to {dst.upper()}",
                fidelity_cost=COST_DIRECT_NATIVE,
                allow_as_intermediate=(dst == "json"),
                category="Data & Text",
            )

    # JSON <-> YAML / YML, plus JSON/YAML -> XLSX/CSV/TSV
    for src in ("json", "yaml", "yml"):
        for dst in ("json", "yaml", "yml"):
            if src == dst:
                continue
            if src == "yaml" and dst == "yml":
                continue
            graph.add_edge(
                src=src,
                dst=dst,
                converter=_make_data_converter(src, dst),
                label=f"{src.upper()} to {dst.upper()}",
                fidelity_cost=COST_DIRECT_NATIVE,
                allow_as_intermediate=(dst == "json"),
                menu_visible=(dst != "yml"),
                category="Data & Text",
            )
        for dst in ("xlsx", "csv", "tsv"):
            graph.add_edge(
                src=src,
                dst=dst,
                converter=_make_data_converter(src, dst),
                label=f"{src.upper()} to {dst.upper()}",
                fidelity_cost=COST_HIGH_FIDELITY,
                may_output_directory=(dst in ("csv", "tsv")),
                directory_suffix="_sheets",
                allow_as_intermediate=False,
                category="Data & Text",
            )
