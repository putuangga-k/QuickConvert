"""Comprehensive end-to-end verification suite for QuickConvert Universal Offline Converter."""

from __future__ import annotations

import csv
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

import openpyxl
import yaml
from docx import Document
from PIL import Image, ImageDraw
from pptx import Presentation

try:
    import pymupdf as fitz  # type: ignore[import-untyped]
except ImportError:  # pragma: no cover
    import fitz  # type: ignore[import-untyped]

try:
    import pillow_heif  # type: ignore[import-untyped]

    pillow_heif.register_heif_opener()
except Exception:
    pass

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from install_menu import ICON_PATH, REG_ROOT_PREFIX, build_registry_specs  # noqa: E402
from main import parse_args  # noqa: E402
from quickconvert.config import load_config  # noqa: E402
from quickconvert.converters.media import build_ffmpeg_command, find_ffmpeg_executable  # noqa: E402
from quickconvert.pipeline import build_default_graph  # noqa: E402
from quickconvert.ui.format_picker import compute_common_targets  # noqa: E402
from quickconvert.ui.progress_window import APP_ICON_PATH, detect_windows_theme  # noqa: E402
from quickconvert.utils import NO_WINDOW_FLAGS  # noqa: E402
from uninstall_menu import collect_target_keys  # noqa: E402


MAIN_PY = PROJECT_ROOT / "main.py"


def _list_quickconvert_temp_dirs() -> set[Path]:
    tmp = Path(tempfile.gettempdir())
    return {p for p in tmp.glob("quickconvert_*") if p.is_dir()}


def _run_cli(target: str, input_file: Path, output_dir: Path | None = None) -> subprocess.CompletedProcess[str]:
    cmd = [sys.executable, str(MAIN_PY), "--no-ui", "--target", target]
    if output_dir is not None:
        cmd.extend(["--output-dir", str(output_dir)])
    cmd.append(str(input_file))
    return subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        creationflags=NO_WINDOW_FLAGS,
        check=False,
    )


class TestQuickConvertUniversal(unittest.TestCase):
    """End-to-end verification of all document, PDF, data, image, media, and Registry flows."""

    work_dir: Path
    sample_pdf: Path
    sample_docx: Path
    sample_xlsx: Path
    sample_csv: Path
    sample_json: Path
    sample_md: Path
    sample_png: Path
    sample_mp4: Path

    @classmethod
    def setUpClass(cls) -> None:
        cls.work_dir = Path(tempfile.mkdtemp(prefix="qc_universal_suite_"))
        cls.sample_pdf = cls.work_dir / "sample_doc.pdf"
        cls.sample_docx = cls.work_dir / "sample_report.docx"
        cls.sample_xlsx = cls.work_dir / "sample_workbook.xlsx"
        cls.sample_csv = cls.work_dir / "sample_table.csv"
        cls.sample_json = cls.work_dir / "sample_config.json"
        cls.sample_md = cls.work_dir / "sample_notes.md"
        cls.sample_png = cls.work_dir / "sample_banner.png"
        cls.sample_mp4 = cls.work_dir / "sample_clip.mp4"

        # 1. Generate a crisp 2-page PDF fixture using PyMuPDF
        pdf_doc = fitz.open()
        try:
            p1 = pdf_doc.new_page(width=612, height=792)
            p1.insert_text((72, 90), "QuickConvert Verification Page 1", fontsize=20)
            p1.insert_text((72, 130), "Testing PDF to DOCX, PNG, JPG, WEBP, TIFF, PPTX, TXT, HTML, MD.", fontsize=12)
            p2 = pdf_doc.new_page(width=612, height=792)
            p2.insert_text((72, 90), "QuickConvert Verification Page 2", fontsize=20)
            p2.insert_text((72, 130), "Second page confirms multi-page TIFF, per-page images, and PPTX slides.", fontsize=12)
            pdf_doc.save(str(cls.sample_pdf))
        finally:
            pdf_doc.close()

        # 2. Generate a 2-page .docx fixture using python-docx
        doc = Document()
        doc.add_heading("QuickConvert Page 1", level=1)
        doc.add_paragraph("This is page one of the automated verification document.")
        doc.add_page_break()
        doc.add_heading("QuickConvert Page 2", level=1)
        doc.add_paragraph("This is page two, verifying multi-page DOCX to PDF to PNG chaining.")
        doc.save(str(cls.sample_docx))

        # 3. Generate a multi-sheet .xlsx workbook using openpyxl
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "Sales"
        ws1.append(["Region", "Quarter", "Revenue"])
        ws1.append(["North", "Q1", 12500])
        ws1.append(["South", "Q1", 9800])
        ws2 = wb.create_sheet("Inventory")
        ws2.append(["SKU", "Item", "Stock"])
        ws2.append(["A-100", "Adapter", 42])
        ws2.append(["B-200", "Cable", 115])
        wb.save(str(cls.sample_xlsx))
        wb.close()

        # 4. Generate sample .csv, .json, and .md fixtures
        with cls.sample_csv.open("w", encoding="utf-8", newline="") as fh:
            w = csv.writer(fh)
            w.writerow(["Name", "Role", "Score"])
            w.writerow(["Alice", "Engineer", "98"])
            w.writerow(["Bob", "Designer", "94"])

        cls.sample_json.write_text(
            json.dumps(
                {"app": "QuickConvert", "offline": True, "targets": ["pdf", "png", "mp3"]},
                indent=2,
            ),
            encoding="utf-8",
        )

        cls.sample_md.write_text(
            "# QuickConvert Notes\n\nThis is **offline** Markdown.\n\n- Fast\n- Private\n",
            encoding="utf-8",
        )

        # 5. Generate a non-square 320x160 RGBA .png fixture
        img = Image.new("RGBA", (320, 160), (20, 110, 220, 255))
        draw = ImageDraw.Draw(img)
        draw.rectangle([20, 20, 300, 140], fill=(240, 180, 40, 210))
        img.save(str(cls.sample_png), format="PNG")

        # 6. Generate a 2-second test .mp4 clip with video + sine audio using ffmpeg
        cfg = load_config()
        ffmpeg_bin = find_ffmpeg_executable(cfg)
        gen_cmd = [
            ffmpeg_bin,
            "-y",
            "-hide_banner",
            "-loglevel",
            "error",
            "-f",
            "lavfi",
            "-i",
            "testsrc=duration=2:size=320x240:rate=24",
            "-f",
            "lavfi",
            "-i",
            "sine=frequency=440:duration=2",
            "-c:v",
            "libx264",
            "-pix_fmt",
            "yuv420p",
            "-c:a",
            "aac",
            "-shortest",
            str(cls.sample_mp4),
        ]
        res = subprocess.run(gen_cmd, capture_output=True, text=True, creationflags=NO_WINDOW_FLAGS)
        if res.returncode != 0 or not cls.sample_mp4.is_file():
            raise RuntimeError(f"Failed to generate sample MP4 fixture: {res.stderr}")

    @classmethod
    def tearDownClass(cls) -> None:
        shutil.rmtree(cls.work_dir, ignore_errors=True)

    def test_01_pdf_universal_targets(self) -> None:
        """Verify .pdf -> .docx, .png/.jpg/.webp (per page), .tiff (multi-page), .txt, .html, .md, .pptx."""
        out_dir = self.work_dir / "out_pdf"
        out_dir.mkdir(exist_ok=True)

        # PDF -> DOCX
        r_docx = _run_cli("docx", self.sample_pdf, out_dir)
        self.assertEqual(r_docx.returncode, 0, msg=f"PDF->DOCX failed: {r_docx.stderr}")
        out_docx = out_dir / "sample_doc.docx"
        self.assertTrue(out_docx.is_file() and zipfile.is_zipfile(out_docx))

        # PDF -> PNG (Per Page)
        r_png = _run_cli("png", self.sample_pdf, out_dir)
        self.assertEqual(r_png.returncode, 0, msg=f"PDF->PNG failed: {r_png.stderr}")
        png_dir = out_dir / "sample_doc_pages"
        self.assertTrue((png_dir / "page_001.png").is_file())
        self.assertTrue((png_dir / "page_002.png").is_file())

        # PDF -> JPG (Per Page)
        r_jpg = _run_cli("jpg", self.sample_pdf, out_dir)
        self.assertEqual(r_jpg.returncode, 0, msg=f"PDF->JPG failed: {r_jpg.stderr}")
        jpg_dir = out_dir / "sample_doc_pages (1)"
        self.assertTrue((jpg_dir / "page_001.jpg").is_file())
        self.assertTrue((jpg_dir / "page_002.jpg").is_file())

        # PDF -> WEBP (Per Page)
        r_webp = _run_cli("webp", self.sample_pdf, out_dir)
        self.assertEqual(r_webp.returncode, 0, msg=f"PDF->WEBP failed: {r_webp.stderr}")
        webp_dir = out_dir / "sample_doc_pages (2)"
        self.assertTrue((webp_dir / "page_001.webp").is_file())
        self.assertTrue((webp_dir / "page_002.webp").is_file())

        # PDF -> TIFF (Multi-Page)
        r_tiff = _run_cli("tiff", self.sample_pdf, out_dir)
        self.assertEqual(r_tiff.returncode, 0, msg=f"PDF->TIFF failed: {r_tiff.stderr}")
        out_tiff = out_dir / "sample_doc.tiff"
        self.assertTrue(out_tiff.is_file())
        with Image.open(out_tiff) as tif:
            self.assertEqual(tif.format, "TIFF")
            self.assertEqual(getattr(tif, "n_frames", 1), 2)

        # Multi-page TIFF -> PDF and per-page PNG
        r_tif_pdf = _run_cli("pdf", out_tiff, out_dir)
        self.assertEqual(r_tif_pdf.returncode, 0, msg=f"Multi-TIFF->PDF failed: {r_tif_pdf.stderr}")
        with fitz.open(str(out_dir / "sample_doc.pdf")) as d:
            self.assertEqual(len(d), 2)

        tif_png_dir_out = self.work_dir / "out_tif_png"
        tif_png_dir_out.mkdir(exist_ok=True)
        r_tif_png = _run_cli("png", out_tiff, tif_png_dir_out)
        self.assertEqual(r_tif_png.returncode, 0, msg=f"Multi-TIFF->PNG failed: {r_tif_png.stderr}")
        self.assertTrue((tif_png_dir_out / "sample_doc_pages" / "page_002.png").is_file())

        # PDF -> TXT, HTML, MD
        for ext in ("txt", "html", "md"):
            res = _run_cli(ext, self.sample_pdf, out_dir)
            self.assertEqual(res.returncode, 0, msg=f"PDF->{ext.upper()} failed: {res.stderr}")
            text_out = (out_dir / f"sample_doc.{ext}").read_text(encoding="utf-8")
            self.assertIn("QuickConvert Verification", text_out)

        # PDF -> PPTX (2 slides)
        r_pptx = _run_cli("pptx", self.sample_pdf, out_dir)
        self.assertEqual(r_pptx.returncode, 0, msg=f"PDF->PPTX failed: {r_pptx.stderr}")
        prs = Presentation(str(out_dir / "sample_doc.pptx"))
        self.assertEqual(len(prs.slides), 2)

    def test_02_docx_and_markup_conversions(self) -> None:
        """Verify .docx -> .pdf, .png (2-hop chained), .txt, and .md -> .pdf / .docx."""
        before_temps = _list_quickconvert_temp_dirs()
        out_dir = self.work_dir / "out_docx"
        out_dir.mkdir(exist_ok=True)

        # DOCX -> PDF
        r_pdf = _run_cli("pdf", self.sample_docx, out_dir)
        self.assertEqual(r_pdf.returncode, 0, msg=f"DOCX->PDF failed: {r_pdf.stderr}")
        with fitz.open(str(out_dir / "sample_report.pdf")) as d:
            self.assertEqual(len(d), 2)

        # DOCX -> PNG (2-hop chained)
        r_png = _run_cli("png", self.sample_docx, out_dir)
        self.assertEqual(r_png.returncode, 0, msg=f"DOCX->PNG failed: {r_png.stderr}")
        pages_dir = out_dir / "sample_report_pages"
        self.assertTrue((pages_dir / "page_001.png").is_file())
        self.assertTrue((pages_dir / "page_002.png").is_file())

        # DOCX -> TXT
        r_txt = _run_cli("txt", self.sample_docx, out_dir)
        self.assertEqual(r_txt.returncode, 0, msg=f"DOCX->TXT failed: {r_txt.stderr}")
        self.assertIn("QuickConvert Page 1", (out_dir / "sample_report.txt").read_text(encoding="utf-8"))

        # MD -> PDF & DOCX
        r_md_pdf = _run_cli("pdf", self.sample_md, out_dir)
        self.assertEqual(r_md_pdf.returncode, 0, msg=f"MD->PDF failed: {r_md_pdf.stderr}")
        self.assertTrue((out_dir / "sample_notes.pdf").is_file())

        r_md_docx = _run_cli("docx", self.sample_md, out_dir)
        self.assertEqual(r_md_docx.returncode, 0, msg=f"MD->DOCX failed: {r_md_docx.stderr}")
        self.assertTrue((out_dir / "sample_notes.docx").is_file())

        after_temps = _list_quickconvert_temp_dirs()
        self.assertEqual(after_temps - before_temps, set(), msg="Temporary directories were leaked!")

    def test_03_spreadsheet_and_structured_data(self) -> None:
        """Verify multi-sheet .xlsx -> .csv (sheets folder), .json, .pdf; .csv -> .xlsx/.json; .json -> .yaml."""
        out_dir = self.work_dir / "out_data"
        out_dir.mkdir(exist_ok=True)

        # Multi-sheet XLSX -> CSV folder (<stem>_sheets/)
        r_csv = _run_cli("csv", self.sample_xlsx, out_dir)
        self.assertEqual(r_csv.returncode, 0, msg=f"XLSX->CSV failed: {r_csv.stderr}")
        sheets_dir = out_dir / "sample_workbook_sheets"
        self.assertTrue(sheets_dir.is_dir(), msg="Multi-sheet XLSX -> CSV must create _sheets folder.")
        self.assertTrue((sheets_dir / "Sales.csv").is_file())
        self.assertTrue((sheets_dir / "Inventory.csv").is_file())

        # Multi-sheet XLSX -> JSON
        r_json = _run_cli("json", self.sample_xlsx, out_dir)
        self.assertEqual(r_json.returncode, 0, msg=f"XLSX->JSON failed: {r_json.stderr}")
        parsed = json.loads((out_dir / "sample_workbook.json").read_text(encoding="utf-8"))
        self.assertIn("Sales", parsed)
        self.assertIn("Inventory", parsed)

        # XLSX -> PDF
        r_pdf = _run_cli("pdf", self.sample_xlsx, out_dir)
        self.assertEqual(r_pdf.returncode, 0, msg=f"XLSX->PDF failed: {r_pdf.stderr}")
        with fitz.open(str(out_dir / "sample_workbook.pdf")) as d:
            self.assertGreaterEqual(len(d), 1)

        # CSV -> XLSX & JSON
        r_c2x = _run_cli("xlsx", self.sample_csv, out_dir)
        self.assertEqual(r_c2x.returncode, 0, msg=f"CSV->XLSX failed: {r_c2x.stderr}")
        self.assertTrue((out_dir / "sample_table.xlsx").is_file())

        r_c2j = _run_cli("json", self.sample_csv, out_dir)
        self.assertEqual(r_c2j.returncode, 0, msg=f"CSV->JSON failed: {r_c2j.stderr}")
        rows = json.loads((out_dir / "sample_table.json").read_text(encoding="utf-8"))
        self.assertEqual(rows[0]["Name"], "Alice")

        # JSON -> YAML
        r_j2y = _run_cli("yaml", self.sample_json, out_dir)
        self.assertEqual(r_j2y.returncode, 0, msg=f"JSON->YAML failed: {r_j2y.stderr}")
        ydata = yaml.safe_load((out_dir / "sample_config.yaml").read_text(encoding="utf-8"))
        self.assertEqual(ydata["app"], "QuickConvert")

    def test_04_universal_images_heic_avif_webp_ico_pdf(self) -> None:
        """Verify .png -> .webp (with collision), .heic, .avif, .ico (square canvas), and .pdf."""
        out_dir = self.work_dir / "out_images"
        out_dir.mkdir(exist_ok=True)

        # PNG -> WEBP twice (collision safety)
        self.assertEqual(_run_cli("webp", self.sample_png, out_dir).returncode, 0)
        self.assertEqual(_run_cli("webp", self.sample_png, out_dir).returncode, 0)
        self.assertTrue((out_dir / "sample_banner.webp").is_file())
        self.assertTrue((out_dir / "sample_banner (1).webp").is_file())

        # PNG -> HEIC and AVIF
        r_heic = _run_cli("heic", self.sample_png, out_dir)
        self.assertEqual(r_heic.returncode, 0, msg=f"PNG->HEIC failed: {r_heic.stderr}")
        with Image.open(out_dir / "sample_banner.heic") as im:
            self.assertEqual(im.size, (320, 160))

        r_avif = _run_cli("avif", self.sample_png, out_dir)
        self.assertEqual(r_avif.returncode, 0, msg=f"PNG->AVIF failed: {r_avif.stderr}")
        with Image.open(out_dir / "sample_banner.avif") as im:
            self.assertEqual(im.size, (320, 160))

        # PNG -> ICO (square canvas with transparent top-left corner)
        r_ico = _run_cli("ico", self.sample_png, out_dir)
        self.assertEqual(r_ico.returncode, 0, msg=f"PNG->ICO failed: {r_ico.stderr}")
        with Image.open(out_dir / "sample_banner.ico") as im:
            self.assertEqual(im.size[0], im.size[1])
            self.assertEqual(im.convert("RGBA").getpixel((0, 0))[3], 0)

        # PNG -> PDF
        r_pdf = _run_cli("pdf", self.sample_png, out_dir)
        self.assertEqual(r_pdf.returncode, 0, msg=f"PNG->PDF failed: {r_pdf.stderr}")
        self.assertTrue((out_dir / "sample_banner.pdf").is_file())

    def test_05_universal_media_mp4_to_mkv_gif_mp3_m4a(self) -> None:
        """Verify .mp4 -> .mkv, .gif, .mp3, .m4a and scoped FFmpeg flags."""
        cfg = load_config()
        ffmpeg_bin = find_ffmpeg_executable(cfg)

        webm_cmd = build_ffmpeg_command(ffmpeg_bin, self.sample_mp4, Path("o.webm"), "mp4", "webm", cfg)
        self.assertIn("libvpx-vp9", webm_cmd)
        self.assertNotIn("-preset", webm_cmd)

        gif_cmd = build_ffmpeg_command(ffmpeg_bin, self.sample_mp4, Path("o.gif"), "mp4", "gif", cfg)
        self.assertNotIn("-preset", gif_cmd)
        self.assertNotIn("-crf", gif_cmd)

        out_dir = self.work_dir / "out_media"
        out_dir.mkdir(exist_ok=True)

        for target in ("mkv", "gif", "mp3", "m4a"):
            res = _run_cli(target, self.sample_mp4, out_dir)
            self.assertEqual(res.returncode, 0, msg=f"MP4->{target.upper()} failed: {res.stderr}")
            out_file = out_dir / f"sample_clip.{target}"
            self.assertTrue(out_file.is_file())
            self.assertGreater(out_file.stat().st_size, 500)

    def test_06_dijkstra_routing_guards_and_failure_cleanup(self) -> None:
        """Verify allow_as_intermediate blocks destructive/nonsense chains and failure cleans temps."""
        graph = build_default_graph()

        # Verified safe chains
        docx_png = graph.find_path("docx", "png")
        self.assertIsNotNone(docx_png)
        assert docx_png is not None
        self.assertEqual([(e.src, e.dst) for e in docx_png], [("docx", "pdf"), ("pdf", "png")])

        # Blocked destructive/cross-domain chains
        self.assertIsNone(graph.find_path("png", "docx"))
        self.assertIsNone(graph.find_path("png", "txt"))
        self.assertIsNone(graph.find_path("gif", "mp3"))
        self.assertIsNone(graph.find_path("pdf", "rtf"))

        # Failure cleanup test
        bad_mp4 = self.work_dir / "corrupted.mp4"
        bad_mp4.write_bytes(b"corrupt data")
        before_temps = _list_quickconvert_temp_dirs()
        out_fail = self.work_dir / "out_fail"
        out_fail.mkdir(exist_ok=True)

        res = _run_cli("mp3", bad_mp4, out_fail)
        self.assertNotEqual(res.returncode, 0)
        self.assertEqual(list(out_fail.iterdir()), [])
        self.assertEqual(_list_quickconvert_temp_dirs() - before_temps, set())

    def test_07_explorer_submenu_cap_pdf_menu_icon_and_uninstaller(self) -> None:
        """Verify <= 12 submenu items, 'More Formats...', full .pdf targets, custom icon, and uninstaller."""
        expected_ico = (PROJECT_ROOT / "assets" / "quickconvert.ico").resolve()
        self.assertTrue(expected_ico.is_file())
        self.assertEqual(ICON_PATH, expected_ico)
        self.assertEqual(APP_ICON_PATH.resolve(), expected_ico)

        with Image.open(expected_ico) as ico:
            self.assertEqual(
                sorted(ico.info.get("sizes", set())),
                [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
            )

        specs = build_registry_specs()
        self.assertGreaterEqual(len(specs), 35)
        spec_by_ext = {s.source_ext: s for s in specs}

        for spec in specs:
            self.assertTrue(spec.base_key_path.startswith(REG_ROOT_PREFIX + "\\."))
            self.assertEqual(spec.icon_path, str(expected_ico))
            self.assertLessEqual(len(spec.actions), 12, msg=f".{spec.source_ext} exceeds Explorer submenu cap!")
            self.assertEqual(spec.actions[-1].label, "More Formats...")
            self.assertIn("--choose-target", spec.actions[-1].command)
            self.assertTrue(spec.actions[-1].separator_before)
            for action in spec.actions:
                self.assertIn("pythonw.exe", action.command.lower())
                self.assertNotIn("python.exe", action.command.lower())
                self.assertNotIn(action.target_ext, ("jpeg", "heif", "tif", "yml"))

        # .pdf must expose DOCX, PNG, JPG, WEBP, TIFF, PPTX, TXT, HTML, MD + More Formats...
        pdf_targets = [a.target_ext for a in spec_by_ext["pdf"].actions]
        self.assertEqual(
            pdf_targets,
            ["docx", "png", "jpg", "webp", "tiff", "pptx", "txt", "html", "md", "__more__"],
        )

        # CLI flags and picker target intersection
        self.assertTrue(parse_args(["--preview-ui"]).preview_ui)
        self.assertTrue(parse_args(["--choose-target", "sample.pdf"]).choose_target)
        theme = detect_windows_theme()
        self.assertIn(theme.bg, ("#F3F3F3", "#202020"))
        common = compute_common_targets([self.sample_pdf], build_default_graph())
        self.assertGreaterEqual(len(common), 9)

        # Uninstaller covers 100% of installed Registry keys
        uninstall_keys = set(collect_target_keys())
        installed_keys = {s.base_key_path for s in specs}
        self.assertTrue(installed_keys.issubset(uninstall_keys))


if __name__ == "__main__":
    unittest.main(verbosity=2)
