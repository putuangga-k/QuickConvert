"""Expanded Office document converters: Microsoft Office COM primary, LibreOffice headless fallback."""

from __future__ import annotations

import contextlib
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Callable, Iterator, TypeVar

from quickconvert.config import AppConfig
from quickconvert.graph import (
    COST_APPROXIMATION,
    COST_DIRECT_NATIVE,
    COST_HIGH_FIDELITY,
    ConversionGraph,
    SubProgressCallback,
)
from quickconvert.utils import (
    NO_WINDOW_FLAGS,
    CancellationToken,
    ConversionError,
    rmtree_robust,
)

if sys.platform == "win32":
    import msvcrt

T = TypeVar("T")

# Standard Windows COM retryable HRESULTs
RPC_E_CALL_REJECTED = -2147418111  # 0x80010001
RPC_E_SERVERCALL_RETRYLATER = -2147417846  # 0x8001010A
_RETRYABLE_HRESULTS = {RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER}

# Word SaveAs2 / Export format constants
_WD_EXPORT_FORMAT_PDF = 17
_WD_FORMAT_DOCX = 16
_WD_FORMAT_RTF = 6
_WD_FORMAT_ENCODED_TEXT = 7
_WD_FORMAT_FILTERED_HTML = 10
_MSO_ENCODING_UTF8 = 65001

# PowerPoint SaveAs constants
_PP_SAVE_AS_PDF = 32
_PP_SAVE_AS_PPTX = 24

# Excel Export / SaveAs constants
_XL_TYPE_PDF = 0
_XL_OPEN_XML_WORKBOOK = 51

WORD_SOURCES: tuple[str, ...] = ("docx", "doc", "rtf", "odt")
PPT_SOURCES: tuple[str, ...] = ("pptx", "ppt", "odp")
EXCEL_SOURCES: tuple[str, ...] = ("xlsx", "xls", "ods")


@contextlib.contextmanager
def office_com_lock(
    cancel_token: CancellationToken,
    timeout: float = 180.0,
) -> Iterator[None]:
    """Serialize Office COM automation across concurrent right-click processes."""
    if sys.platform != "win32":
        yield
        return

    lock_path = Path(tempfile.gettempdir()) / "quickconvert_office_com.lock"
    fh = open(lock_path, "a+b")
    acquired = False
    deadline = time.monotonic() + timeout
    try:
        while True:
            cancel_token.raise_if_cancelled()
            try:
                fh.seek(0)
                msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
                acquired = True
                break
            except OSError:
                if time.monotonic() >= deadline:
                    raise ConversionError("Timed out waiting for Office COM lock.")
                time.sleep(0.15)
        yield
    finally:
        if acquired:
            try:
                fh.seek(0)
                msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)
            except OSError:
                pass
        fh.close()


def call_com_with_retry(
    fn: Callable[[], T],
    cancel_token: CancellationToken,
    retries: int = 10,
    base_delay: float = 0.35,
) -> T:
    """Retry COM calls when Office temporarily responds with RPC_E_CALL_REJECTED."""
    last_exc: Exception | None = None
    for attempt in range(retries):
        cancel_token.raise_if_cancelled()
        try:
            return fn()
        except Exception as exc:
            hresult = getattr(exc, "hresult", None)
            if hresult is None and getattr(exc, "args", None):
                first = exc.args[0]
                if isinstance(first, int):
                    hresult = first
            if hresult in _RETRYABLE_HRESULTS and attempt + 1 < retries:
                last_exc = exc
                time.sleep(base_delay * (attempt + 1))
                continue
            raise
    if last_exc is not None:
        raise last_exc
    raise ConversionError("COM call failed after retries.")


def _word_convert_com(
    src_path: Path,
    dst_path: Path,
    dst_ext: str,
    cancel_token: CancellationToken,
) -> None:
    """Convert Word-compatible document (or PDF) to target format via Word COM."""
    if sys.platform != "win32":
        raise ConversionError("Word COM automation is only supported on Windows.")

    try:
        import pythoncom  # type: ignore[import-untyped]
        import win32com.client  # type: ignore[import-untyped]
    except ImportError as exc:
        raise ConversionError("pywin32 is not installed.") from exc

    abs_in = str(Path(src_path).resolve())
    abs_out = str(Path(dst_path).resolve())

    with office_com_lock(cancel_token):
        pythoncom.CoInitialize()
        app = None
        doc = None
        try:
            app = call_com_with_retry(
                lambda: win32com.client.DispatchEx("Word.Application"),
                cancel_token,
            )
            app.Visible = False
            app.DisplayAlerts = 0
            doc = call_com_with_retry(
                lambda: app.Documents.Open(
                    abs_in,
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    Visible=False,
                ),
                cancel_token,
            )
            if dst_ext == "pdf":
                call_com_with_retry(
                    lambda: doc.ExportAsFixedFormat(
                        OutputFileName=abs_out,
                        ExportFormat=_WD_EXPORT_FORMAT_PDF,
                    ),
                    cancel_token,
                )
            elif dst_ext == "docx":
                call_com_with_retry(
                    lambda: doc.SaveAs2(abs_out, FileFormat=_WD_FORMAT_DOCX),
                    cancel_token,
                )
            elif dst_ext == "rtf":
                call_com_with_retry(
                    lambda: doc.SaveAs2(abs_out, FileFormat=_WD_FORMAT_RTF),
                    cancel_token,
                )
            elif dst_ext == "txt":
                call_com_with_retry(
                    lambda: doc.SaveAs2(
                        abs_out,
                        FileFormat=_WD_FORMAT_ENCODED_TEXT,
                        Encoding=_MSO_ENCODING_UTF8,
                    ),
                    cancel_token,
                )
            elif dst_ext == "html":
                call_com_with_retry(
                    lambda: doc.SaveAs2(abs_out, FileFormat=_WD_FORMAT_FILTERED_HTML),
                    cancel_token,
                )
            else:
                raise ConversionError(f"Unsupported Word COM target format: .{dst_ext}")
        finally:
            if doc is not None:
                try:
                    call_com_with_retry(lambda: doc.Close(False), CancellationToken(), retries=3, base_delay=0.2)
                except Exception:
                    pass
            if app is not None:
                try:
                    call_com_with_retry(lambda: app.Quit(False), CancellationToken(), retries=3, base_delay=0.2)
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    if not dst_path.is_file() or dst_path.stat().st_size == 0:
        raise ConversionError(f"Word COM did not produce a valid '.{dst_ext}' output.")


def _powerpoint_convert_com(
    src_path: Path,
    dst_path: Path,
    dst_ext: str,
    cancel_token: CancellationToken,
) -> None:
    """Convert PowerPoint-compatible file via single-instance PowerPoint COM safely."""
    if sys.platform != "win32":
        raise ConversionError("PowerPoint COM automation is only supported on Windows.")

    try:
        import pythoncom  # type: ignore[import-untyped]
        import win32com.client  # type: ignore[import-untyped]
    except ImportError as exc:
        raise ConversionError("pywin32 is not installed.") from exc

    abs_in = str(Path(src_path).resolve())
    abs_out = str(Path(dst_path).resolve())
    fmt_code = _PP_SAVE_AS_PDF if dst_ext == "pdf" else _PP_SAVE_AS_PPTX

    with office_com_lock(cancel_token):
        pythoncom.CoInitialize()
        app = None
        pres = None
        try:
            app = call_com_with_retry(
                lambda: win32com.client.Dispatch("PowerPoint.Application"),
                cancel_token,
            )
            pres = call_com_with_retry(
                lambda: app.Presentations.Open(
                    abs_in,
                    ReadOnly=True,
                    Untitled=False,
                    WithWindow=False,
                ),
                cancel_token,
            )
            call_com_with_retry(lambda: pres.SaveAs(abs_out, fmt_code), cancel_token)
        finally:
            if pres is not None:
                try:
                    call_com_with_retry(lambda: pres.Close(), CancellationToken(), retries=3, base_delay=0.2)
                except Exception:
                    pass
            # Single-instance COM server: only quit when no user presentations remain open.
            if app is not None:
                try:
                    remaining = call_com_with_retry(
                        lambda: int(app.Presentations.Count),
                        CancellationToken(),
                        retries=3,
                        base_delay=0.2,
                    )
                    if remaining == 0:
                        call_com_with_retry(lambda: app.Quit(), CancellationToken(), retries=3, base_delay=0.2)
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    if not dst_path.is_file() or dst_path.stat().st_size == 0:
        raise ConversionError(f"PowerPoint COM did not produce a valid '.{dst_ext}' output.")


def _excel_convert_com(
    src_path: Path,
    dst_path: Path,
    dst_ext: str,
    cancel_token: CancellationToken,
) -> None:
    """Convert Excel-compatible spreadsheet via isolated Excel COM."""
    if sys.platform != "win32":
        raise ConversionError("Excel COM automation is only supported on Windows.")

    try:
        import pythoncom  # type: ignore[import-untyped]
        import win32com.client  # type: ignore[import-untyped]
    except ImportError as exc:
        raise ConversionError("pywin32 is not installed.") from exc

    abs_in = str(Path(src_path).resolve())
    abs_out = str(Path(dst_path).resolve())

    with office_com_lock(cancel_token):
        pythoncom.CoInitialize()
        app = None
        wb = None
        try:
            app = call_com_with_retry(
                lambda: win32com.client.DispatchEx("Excel.Application"),
                cancel_token,
            )
            app.Visible = False
            app.DisplayAlerts = False
            wb = call_com_with_retry(
                lambda: app.Workbooks.Open(
                    abs_in,
                    UpdateLinks=0,
                    ReadOnly=True,
                    IgnoreReadOnlyRecommended=True,
                ),
                cancel_token,
            )
            if dst_ext == "pdf":
                call_com_with_retry(
                    lambda: wb.ExportAsFixedFormat(_XL_TYPE_PDF, abs_out),
                    cancel_token,
                )
            elif dst_ext == "xlsx":
                call_com_with_retry(
                    lambda: wb.SaveAs(abs_out, FileFormat=_XL_OPEN_XML_WORKBOOK),
                    cancel_token,
                )
            else:
                raise ConversionError(f"Unsupported Excel COM target format: .{dst_ext}")
        finally:
            if wb is not None:
                try:
                    call_com_with_retry(lambda: wb.Close(False), CancellationToken(), retries=3, base_delay=0.2)
                except Exception:
                    pass
            if app is not None:
                try:
                    call_com_with_retry(lambda: app.Quit(), CancellationToken(), retries=3, base_delay=0.2)
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    if not dst_path.is_file() or dst_path.stat().st_size == 0:
        raise ConversionError(f"Excel COM did not produce a valid '.{dst_ext}' output.")


def find_soffice_executable(config: AppConfig) -> str | None:
    """Locate LibreOffice `soffice` executable from config, PATH, or standard install paths."""
    if config.soffice_path:
        cfg_path = Path(config.soffice_path)
        if cfg_path.is_file():
            return str(cfg_path)

    on_path = shutil.which("soffice") or shutil.which("soffice.exe")
    if on_path:
        return on_path

    candidates = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ]
    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)
    return None


def _convert_via_libreoffice(
    src_path: Path,
    dst_path: Path,
    dst_ext: str,
    config: AppConfig,
    cancel_token: CancellationToken,
) -> None:
    """Convert document to `dst_ext` via LibreOffice headless fallback."""
    soffice = find_soffice_executable(config)
    if not soffice:
        raise ConversionError("LibreOffice (soffice) was not found.")

    convert_filter = "txt:Text (encoded):UTF8" if dst_ext == "txt" else dst_ext
    profile_dir = Path(tempfile.mkdtemp(prefix="quickconvert_lo_profile_"))
    out_dir = Path(tempfile.mkdtemp(prefix="quickconvert_lo_out_"))
    proc: subprocess.Popen[bytes] | None = None
    try:
        profile_uri = profile_dir.resolve().as_uri()
        cmd = [
            soffice,
            f"-env:UserInstallation={profile_uri}",
            "--headless",
            "--nologo",
            "--nodefault",
            "--norestore",
            "--convert-to",
            convert_filter,
            "--outdir",
            str(out_dir.resolve()),
            str(src_path.resolve()),
        ]
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=NO_WINDOW_FLAGS,
        )
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
            err_msg = stderr.decode("utf-8", errors="replace").strip()
            raise ConversionError(
                f"LibreOffice conversion failed: {err_msg or f'exit code {proc.returncode}'}"
            )

        produced = out_dir / f"{src_path.stem}.{dst_ext}"
        if not produced.is_file():
            matches = list(out_dir.glob(f"*.{dst_ext}"))
            if not matches:
                raise ConversionError(f"LibreOffice did not produce a '.{dst_ext}' output file.")
            produced = matches[0]

        os.replace(produced, dst_path)
    finally:
        rmtree_robust(profile_dir)
        rmtree_robust(out_dir)


def _docx_to_txt_python(src_path: Path, dst_path: Path) -> None:
    """Pure-Python fallback for .docx -> .txt using python-docx."""
    from docx import Document

    doc = Document(str(src_path))
    parts: list[str] = []
    for para in doc.paragraphs:
        if para.text:
            parts.append(para.text)
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
            if any(cells):
                parts.append("\t".join(cells))
    text = "\n".join(parts).strip() + "\n"
    dst_path.write_text(text, encoding="utf-8")


def _make_office_converter(family: str, dst_ext: str):
    def _convert(
        src_path: Path,
        dst_path: Path,
        config: AppConfig,
        cancel_token: CancellationToken,
        sub_progress_cb: SubProgressCallback | None = None,
    ) -> None:
        cancel_token.raise_if_cancelled()
        src_ext = src_path.suffix.lower().lstrip(".")

        if sub_progress_cb:
            sub_progress_cb(None, f"Exporting via Microsoft {family.title()}...")

        com_error: Exception | None = None
        try:
            if family == "word":
                _word_convert_com(src_path, dst_path, dst_ext, cancel_token)
            elif family == "powerpoint":
                _powerpoint_convert_com(src_path, dst_path, dst_ext, cancel_token)
            elif family == "excel":
                _excel_convert_com(src_path, dst_path, dst_ext, cancel_token)
            else:
                raise ConversionError(f"Unknown Office family: {family}")
            return
        except Exception as exc:
            if cancel_token.is_cancelled:
                raise
            com_error = exc

        if sub_progress_cb:
            sub_progress_cb(None, "Falling back to LibreOffice headless...")

        try:
            _convert_via_libreoffice(src_path, dst_path, dst_ext, config, cancel_token)
            return
        except Exception as lo_exc:
            if cancel_token.is_cancelled:
                raise
            if src_ext == "docx" and dst_ext == "txt":
                _docx_to_txt_python(src_path, dst_path)
                return
            raise ConversionError(
                f"Document conversion failed. Office COM error: {com_error}; "
                f"LibreOffice fallback error: {lo_exc}"
            ) from lo_exc

    return _convert


def register_document_converters(graph: ConversionGraph) -> None:
    """Register Word, PowerPoint, and Excel Office conversion edges."""
    # Word-family sources (.docx, .doc, .rtf, .odt)
    word_targets: list[tuple[str, int, bool, str]] = [
        ("pdf", COST_DIRECT_NATIVE, True, "Document"),
        ("docx", COST_DIRECT_NATIVE, True, "Document"),
        ("rtf", COST_HIGH_FIDELITY, False, "Document"),
        ("txt", COST_APPROXIMATION, False, "Data & Text"),
        ("html", COST_HIGH_FIDELITY, False, "Data & Text"),
    ]
    for src in WORD_SOURCES:
        for dst, cost, can_chain, cat in word_targets:
            if src == dst:
                continue
            graph.add_edge(
                src=src,
                dst=dst,
                converter=_make_office_converter("word", dst),
                label=f"{src.upper()} to {dst.upper()}",
                fidelity_cost=cost,
                allow_as_intermediate=can_chain,
                category=cat,
            )

    # PowerPoint-family sources (.pptx, .ppt, .odp)
    for src in PPT_SOURCES:
        graph.add_edge(
            src=src,
            dst="pdf",
            converter=_make_office_converter("powerpoint", "pdf"),
            label=f"{src.upper()} to PDF",
            fidelity_cost=COST_DIRECT_NATIVE,
            allow_as_intermediate=True,
            category="Document",
        )
        if src != "pptx":
            graph.add_edge(
                src=src,
                dst="pptx",
                converter=_make_office_converter("powerpoint", "pptx"),
                label=f"{src.upper()} to PPTX",
                fidelity_cost=COST_DIRECT_NATIVE,
                allow_as_intermediate=True,
                category="Document",
            )

    # Excel-family sources (.xlsx, .xls, .ods)
    for src in EXCEL_SOURCES:
        graph.add_edge(
            src=src,
            dst="pdf",
            converter=_make_office_converter("excel", "pdf"),
            label=f"{src.upper()} to PDF",
            fidelity_cost=COST_DIRECT_NATIVE,
            allow_as_intermediate=True,
            category="Document",
        )
        if src != "xlsx":
            graph.add_edge(
                src=src,
                dst="xlsx",
                converter=_make_office_converter("excel", "xlsx"),
                label=f"{src.upper()} to XLSX",
                fidelity_cost=COST_DIRECT_NATIVE,
                allow_as_intermediate=True,
                category="Data & Text",
            )
