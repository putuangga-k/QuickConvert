"""Universal offline PDF converter: DOCX, PNG/JPG/WEBP (per page), TIFF (multi-page), PPTX, TXT, HTML, MD."""

from __future__ import annotations

import io
import warnings
from pathlib import Path

from PIL import Image
from pptx import Presentation

from quickconvert.config import AppConfig
from quickconvert.converters.documents import _word_convert_com
from quickconvert.graph import (
    COST_APPROXIMATION,
    COST_HIGH_FIDELITY,
    ConversionGraph,
    SubProgressCallback,
)
from quickconvert.utils import CancellationToken, ConversionError

try:
    import pymupdf as fitz  # type: ignore[import-untyped]
except ImportError:  # pragma: no cover
    import fitz  # type: ignore[import-untyped]


# PowerPoint max slide dimension is 56 inches = 51,206,400 EMU
_EMU_PER_PT = 12700
_MAX_SLIDE_EMU = 56 * 914400


def _convert_pdf_to_docx(
    src_pdf: Path,
    dst_docx: Path,
    config: AppConfig,
    cancel_token: CancellationToken,
    sub_progress_cb: SubProgressCallback | None = None,
) -> None:
    """Convert .pdf -> .docx via Microsoft Word COM first, falling back to pdf2docx."""
    cancel_token.raise_if_cancelled()

    if sub_progress_cb:
        sub_progress_cb(None, "Reconstructing DOCX via Microsoft Word...")

    com_error: Exception | None = None
    try:
        _word_convert_com(src_pdf, dst_docx, "docx", cancel_token)
        return
    except Exception as exc:
        if cancel_token.is_cancelled:
            raise
        com_error = exc

    if sub_progress_cb:
        sub_progress_cb(None, "Falling back to pdf2docx...")

    cv = None
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", FutureWarning)
            from pdf2docx import Converter  # type: ignore[import-untyped]

        cv = Converter(str(src_pdf.resolve()))
        cv.convert(str(dst_docx.resolve()), start=0, end=None)
    except Exception as pdf2docx_exc:
        if cancel_token.is_cancelled:
            raise
        raise ConversionError(
            f"PDF to DOCX failed. Word COM error: {com_error}; pdf2docx error: {pdf2docx_exc}"
        ) from pdf2docx_exc
    finally:
        if cv is not None:
            try:
                cv.close()
            except Exception:
                pass


def _make_pdf_to_per_page_images(target_ext: str):
    ext = target_ext.lower().lstrip(".")

    def _convert(
        src_pdf: Path,
        dst_dir: Path,
        config: AppConfig,
        cancel_token: CancellationToken,
        sub_progress_cb: SubProgressCallback | None = None,
    ) -> None:
        cancel_token.raise_if_cancelled()
        dst_dir.mkdir(parents=True, exist_ok=True)

        zoom = config.image_dpi / 72.0
        matrix = fitz.Matrix(zoom, zoom)

        doc = fitz.open(str(src_pdf))
        try:
            page_count = len(doc)
            if page_count == 0:
                raise ConversionError(f"PDF '{src_pdf.name}' contains no pages.")

            for idx in range(page_count):
                cancel_token.raise_if_cancelled()
                page = doc.load_page(idx)
                pix = page.get_pixmap(matrix=matrix, alpha=False)
                out_file = dst_dir / f"page_{idx + 1:03d}.{ext}"

                if ext == "png":
                    pix.save(str(out_file))
                elif ext in ("jpg", "jpeg"):
                    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                    try:
                        img.save(str(out_file), format="JPEG", quality=config.jpeg_quality, optimize=True)
                    finally:
                        img.close()
                elif ext == "webp":
                    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                    try:
                        img.save(str(out_file), format="WEBP", quality=config.webp_quality)
                    finally:
                        img.close()
                else:
                    raise ConversionError(f"Unsupported PDF page image format: .{ext}")

                if sub_progress_cb:
                    sub_progress_cb((idx + 1) / page_count, f"Page {idx + 1}/{page_count}")
        finally:
            doc.close()

    return _convert


def _convert_pdf_to_multipage_tiff(
    src_pdf: Path,
    dst_tiff: Path,
    config: AppConfig,
    cancel_token: CancellationToken,
    sub_progress_cb: SubProgressCallback | None = None,
) -> None:
    """Render all PDF pages into a single multi-page LZW-compressed TIFF file."""
    cancel_token.raise_if_cancelled()
    zoom = config.image_dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)

    doc = fitz.open(str(src_pdf))
    frames: list[Image.Image] = []
    try:
        page_count = len(doc)
        if page_count == 0:
            raise ConversionError(f"PDF '{src_pdf.name}' contains no pages.")

        for idx in range(page_count):
            cancel_token.raise_if_cancelled()
            page = doc.load_page(idx)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            frames.append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
            if sub_progress_cb:
                sub_progress_cb((idx + 1) / page_count, f"Page {idx + 1}/{page_count}")

        frames[0].save(
            dst_tiff,
            format="TIFF",
            save_all=True,
            append_images=frames[1:],
            compression="tiff_lzw",
            dpi=(config.image_dpi, config.image_dpi),
        )
    finally:
        for frame in frames:
            try:
                frame.close()
            except Exception:
                pass
        doc.close()


def _convert_pdf_to_txt(
    src_pdf: Path,
    dst_txt: Path,
    config: AppConfig,
    cancel_token: CancellationToken,
    sub_progress_cb: SubProgressCallback | None = None,
) -> None:
    """Extract clean UTF-8 text from PDF using PyMuPDF."""
    cancel_token.raise_if_cancelled()
    doc = fitz.open(str(src_pdf))
    try:
        page_count = len(doc)
        if page_count == 0:
            raise ConversionError(f"PDF '{src_pdf.name}' contains no pages.")

        chunks: list[str] = []
        for idx in range(page_count):
            cancel_token.raise_if_cancelled()
            text = doc.load_page(idx).get_text("text").strip()
            if text:
                chunks.append(text)
            if sub_progress_cb:
                sub_progress_cb((idx + 1) / page_count, f"Page {idx + 1}/{page_count}")

        output = "\n\n".join(chunks).strip()
        if not output:
            raise ConversionError(
                f"PDF '{src_pdf.name}' contains no extractable text (it may be a scanned image)."
            )
        dst_txt.write_text(output + "\n", encoding="utf-8")
    finally:
        doc.close()


def _convert_pdf_to_html(
    src_pdf: Path,
    dst_html: Path,
    config: AppConfig,
    cancel_token: CancellationToken,
    sub_progress_cb: SubProgressCallback | None = None,
) -> None:
    """Extract structured HTML from PDF using PyMuPDF."""
    cancel_token.raise_if_cancelled()
    doc = fitz.open(str(src_pdf))
    try:
        page_count = len(doc)
        if page_count == 0:
            raise ConversionError(f"PDF '{src_pdf.name}' contains no pages.")

        sections: list[str] = []
        for idx in range(page_count):
            cancel_token.raise_if_cancelled()
            page_html = doc.load_page(idx).get_text("html")
            sections.append(f'<section class="pdf-page" data-page="{idx + 1}">\n{page_html}\n</section>')
            if sub_progress_cb:
                sub_progress_cb((idx + 1) / page_count, f"Page {idx + 1}/{page_count}")

        full_html = (
            "<!DOCTYPE html>\n"
            '<html lang="en">\n<head>\n<meta charset="utf-8">\n'
            f"<title>{src_pdf.stem}</title>\n"
            "<style>body{font-family:Segoe UI,Arial,sans-serif;max-width:900px;margin:24px auto;"
            "padding:0 16px;background:#fff;color:#1b1b1b;}section.pdf-page{margin-bottom:32px;"
            "padding-bottom:24px;border-bottom:1px solid #e5e5e5;}</style>\n"
            "</head>\n<body>\n"
            + "\n".join(sections)
            + "\n</body>\n</html>\n"
        )
        dst_html.write_text(full_html, encoding="utf-8")
    finally:
        doc.close()


def _convert_pdf_to_md(
    src_pdf: Path,
    dst_md: Path,
    config: AppConfig,
    cancel_token: CancellationToken,
    sub_progress_cb: SubProgressCallback | None = None,
) -> None:
    """Extract structured Markdown from PDF via pymupdf4llm or structured PyMuPDF blocks."""
    cancel_token.raise_if_cancelled()
    if sub_progress_cb:
        sub_progress_cb(None, "Extracting structured Markdown...")

    md_text = ""
    try:
        import pymupdf4llm  # type: ignore[import-untyped]

        md_text = str(pymupdf4llm.to_markdown(str(src_pdf))).strip()
    except Exception:
        md_text = ""

    if not md_text:
        doc = fitz.open(str(src_pdf))
        try:
            page_count = len(doc)
            if page_count == 0:
                raise ConversionError(f"PDF '{src_pdf.name}' contains no pages.")
            parts: list[str] = []
            for idx in range(page_count):
                cancel_token.raise_if_cancelled()
                page_dict = doc.load_page(idx).get_text("dict")
                page_lines: list[str] = []
                for block in page_dict.get("blocks", []):
                    if block.get("type") != 0:
                        continue
                    lines_text: list[str] = []
                    max_size = 0.0
                    for line in block.get("lines", []):
                        spans = line.get("spans", [])
                        line_str = "".join(s.get("text", "") for s in spans).strip()
                        if line_str:
                            lines_text.append(line_str)
                            for s in spans:
                                max_size = max(max_size, float(s.get("size", 0.0)))
                    para = " ".join(lines_text).strip()
                    if not para:
                        continue
                    if max_size >= 18.0:
                        page_lines.append(f"# {para}")
                    elif max_size >= 14.0:
                        page_lines.append(f"## {para}")
                    else:
                        page_lines.append(para)
                if page_lines:
                    parts.append("\n\n".join(page_lines))
                if sub_progress_cb:
                    sub_progress_cb((idx + 1) / page_count, f"Page {idx + 1}/{page_count}")
            md_text = "\n\n---\n\n".join(parts).strip()
        finally:
            doc.close()

    if not md_text:
        raise ConversionError(f"PDF '{src_pdf.name}' contains no extractable text for Markdown.")

    dst_md.write_text(md_text + "\n", encoding="utf-8")


def _convert_pdf_to_pptx(
    src_pdf: Path,
    dst_pptx: Path,
    config: AppConfig,
    cancel_token: CancellationToken,
    sub_progress_cb: SubProgressCallback | None = None,
) -> None:
    """Render each PDF page as a crisp high-DPI slide preserving PDF page aspect ratio."""
    cancel_token.raise_if_cancelled()
    zoom = config.image_dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)

    doc = fitz.open(str(src_pdf))
    try:
        page_count = len(doc)
        if page_count == 0:
            raise ConversionError(f"PDF '{src_pdf.name}' contains no pages.")

        first_rect = doc.load_page(0).rect
        slide_w = int(first_rect.width * _EMU_PER_PT)
        slide_h = int(first_rect.height * _EMU_PER_PT)
        scale_down = max(slide_w / _MAX_SLIDE_EMU, slide_h / _MAX_SLIDE_EMU, 1.0)
        slide_w = max(914400, int(slide_w / scale_down))
        slide_h = max(914400, int(slide_h / scale_down))

        prs = Presentation()
        prs.slide_width = slide_w
        prs.slide_height = slide_h
        blank_layout = prs.slide_layouts[6]

        for idx in range(page_count):
            cancel_token.raise_if_cancelled()
            page = doc.load_page(idx)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            png_bytes = pix.tobytes("png")

            slide = prs.slides.add_slide(blank_layout)
            # Fit image cleanly onto slide preserving exact page aspect ratio
            img_ratio = pix.width / max(1, pix.height)
            slide_ratio = slide_w / max(1, slide_h)
            if img_ratio >= slide_ratio:
                pic_w = slide_w
                pic_h = int(slide_w / img_ratio)
                left = 0
                top = (slide_h - pic_h) // 2
            else:
                pic_h = slide_h
                pic_w = int(slide_h * img_ratio)
                left = (slide_w - pic_w) // 2
                top = 0

            slide.shapes.add_picture(io.BytesIO(png_bytes), left, top, width=pic_w, height=pic_h)
            if sub_progress_cb:
                sub_progress_cb((idx + 1) / page_count, f"Slide {idx + 1}/{page_count}")

        prs.save(str(dst_pptx))
    finally:
        doc.close()


def register_pdf_converters(graph: ConversionGraph) -> None:
    """Register all universal PDF conversion edges (non-intermediate terminal targets)."""
    # PDF -> DOCX (layout reconstruction: cost 5, never intermediate)
    graph.add_edge(
        src="pdf",
        dst="docx",
        converter=_convert_pdf_to_docx,
        label="PDF to DOCX",
        fidelity_cost=COST_APPROXIMATION,
        outputs_directory=False,
        allow_as_intermediate=False,
        category="Document",
    )

    # PDF -> PNG / JPG / WEBP (per-page folder: cost 2, never intermediate)
    for target in ("png", "jpg", "webp"):
        graph.add_edge(
            src="pdf",
            dst=target,
            converter=_make_pdf_to_per_page_images(target),
            label=f"PDF to {target.upper()} (Per Page)",
            fidelity_cost=COST_HIGH_FIDELITY,
            outputs_directory=True,
            directory_suffix="_pages",
            allow_as_intermediate=False,
            category="Image",
            menu_label_override=f"Convert to {target.upper()} (Per Page)",
        )

    # PDF -> TIFF (single multi-page TIFF file: cost 2, never intermediate)
    graph.add_edge(
        src="pdf",
        dst="tiff",
        converter=_convert_pdf_to_multipage_tiff,
        label="PDF to TIFF (Multi-Page)",
        fidelity_cost=COST_HIGH_FIDELITY,
        outputs_directory=False,
        allow_as_intermediate=False,
        category="Image",
        menu_label_override="Convert to TIFF (Multi-Page)",
    )

    # PDF -> PPTX (slide deck render: cost 5, never intermediate)
    graph.add_edge(
        src="pdf",
        dst="pptx",
        converter=_convert_pdf_to_pptx,
        label="PDF to PPTX",
        fidelity_cost=COST_APPROXIMATION,
        outputs_directory=False,
        allow_as_intermediate=False,
        category="Document",
    )

    # PDF -> TXT / HTML / MD (layout/text extraction: cost 5, never intermediate)
    graph.add_edge(
        src="pdf",
        dst="txt",
        converter=_convert_pdf_to_txt,
        label="PDF to TXT",
        fidelity_cost=COST_APPROXIMATION,
        outputs_directory=False,
        allow_as_intermediate=False,
        category="Data & Text",
    )
    graph.add_edge(
        src="pdf",
        dst="html",
        converter=_convert_pdf_to_html,
        label="PDF to HTML",
        fidelity_cost=COST_APPROXIMATION,
        outputs_directory=False,
        allow_as_intermediate=False,
        category="Data & Text",
    )
    graph.add_edge(
        src="pdf",
        dst="md",
        converter=_convert_pdf_to_md,
        label="PDF to MD",
        fidelity_cost=COST_APPROXIMATION,
        outputs_directory=False,
        allow_as_intermediate=False,
        category="Data & Text",
    )
