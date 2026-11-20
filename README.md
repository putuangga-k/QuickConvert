# QuickConvert

**QuickConvert** is a 100% offline, silent Windows right-click context menu file converter written in Python. Right-click any supported document, PDF, spreadsheet, image, video, audio, or structured data file in Windows File Explorer to convert it in place — without opening heavy applications, uploading files to the cloud, flashing terminal windows, or needing Administrator privileges.

---

## Why QuickConvert?

- **100% Offline & Private:** Every conversion runs locally on your machine using Microsoft Office COM (with LibreOffice headless fallback), PyMuPDF, Pillow + `pillow-heif`, `openpyxl`, `python-pptx`, `pdf2docx`, and bundled FFmpeg.
- **No Administrator Privileges Required:** All Explorer context menus and per-user app files install exclusively under `HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations` and `%LOCALAPPDATA%\Programs\QuickConvert`.
- **Silent Execution:** Background conversion jobs launch via `pythonw.exe` (source mode) or the windowed `QuickConvert.exe` GUI binary (installer mode) with `CREATE_NO_WINDOW` on all child subprocesses so no console window ever flashes.
- **Smart Weighted Graph Routing (Dijkstra):** File extensions are modeled as nodes and converters as weighted edges (`fidelity_cost` + `allow_as_intermediate`). QuickConvert automatically chains verified high-fidelity multi-step bridges (e.g., `.docx → .pdf → .png` per page) while blocking lossy or cross-domain nonsense chains (e.g., `.png → .pdf → .docx` or `.gif → .mp4 → .mp3`).
- **File-Safe Atomic Commits & Integrity Validation:** Every conversion stages output inside a temporary directory, runs format-specific post-conversion integrity checks (`Image.verify()`, PDF page count, OOXML ZIP/XML structure, UTF-8/JSON/YAML/CSV parsing, FFmpeg stream check), and commits atomically via `os.replace()` with collision auto-increment (`file (1).ext`) so existing files are never overwritten. Temporary files are always cleaned up inside a `finally` block — even if a conversion fails or is cancelled.
- **Native Windows 11 Fluent UI:** High-DPI aware floating progress dialog and `"More Formats..."` picker with automatic OS Light/Dark theme sync, DWM titlebar matching, custom multi-resolution icon, hover-paused auto-close, and batch PID vertical offsetting.

---

## Architecture

```mermaid
flowchart LR
    EXP["Windows File Explorer<br/>HKCU Context Menu"] -->|"Quick Target (--target)"| MAIN["main.py / QuickConvert.exe"]
    EXP -->|"More Formats... (--choose-target)"| PICK["Fluent Format Picker"]
    PICK --> MAIN

    MAIN --> UI["Fluent ProgressWindow<br/>(Background Worker Thread)"]
    UI --> PIPE["Conversion Pipeline<br/>(Weighted Dijkstra Router)"]

    subgraph Converters["Offline Converters"]
        DOC["Office Documents<br/>Word / PowerPoint / Excel COM<br/>+ LibreOffice Fallback"]
        PDF["PDF Engine<br/>PyMuPDF + Word COM / pdf2docx<br/>+ python-pptx"]
        DATA["Text, Markup & Data<br/>openpyxl + PyYAML + Markdown<br/>+ PyMuPDF Story"]
        IMG["Universal Images<br/>Pillow + pillow-heif<br/>(EXIF & ICC Preserved)"]
        MED["Universal Media<br/>Bundled FFmpeg<br/>(Container-Matched Codecs)"]
    end

    PIPE --> Converters
    Converters --> STAGE["Temp Staging Dir<br/>(%TEMP%/quickconvert_*)"]
    STAGE --> VAL["Mandatory Output Validator"]
    VAL -->|"Atomic os.replace()"| OUT["Non-Colliding Destination<br/>(Auto-Incremented)"]
    STAGE -.->|"Guaranteed finally cleanup"| CLEAN["Temp Removed"]
```

---

## Supported Formats (47 Extensions)

| Category | Supported Input Extensions | Available Output Formats | Highlights |
| --- | --- | --- | --- |
| **PDF** | `.pdf` | `.docx`, `.png` *(Per Page)*, `.jpg` *(Per Page)*, `.webp` *(Per Page)*, `.tiff` *(Multi-Page)*, `.pptx`, `.txt`, `.html`, `.md` | Word COM primary for `.pdf → .docx` with `pdf2docx` fallback; 200 DPI configurable rendering; aspect-ratio-matched slides for `.pptx`. |
| **Word & Rich Text** | `.docx`, `.doc`, `.rtf`, `.odt` | `.pdf`, `.docx`, `.rtf`, `.txt`, `.html`, plus automatic chaining to `.png`/`.jpg`/`.webp` *(Per Page)* and `.tiff` *(Multi-Page)* | Microsoft Word COM primary (with cross-process lockfile & `RPC_E_CALL_REJECTED` retry) + LibreOffice headless fallback. |
| **Presentations** | `.pptx`, `.ppt`, `.odp` | `.pdf`, `.pptx`, plus automatic chaining to `.png`/`.jpg`/`.webp` *(Per Page)* and `.tiff` *(Multi-Page)* | Single-instance PowerPoint COM safety (`Quit()` only called when `Presentations.Count == 0` so open user decks stay untouched). |
| **Spreadsheets** | `.xlsx`, `.xls`, `.ods` | `.pdf`, `.xlsx`, `.csv`, `.tsv`, `.json`, plus automatic chaining to per-page images | Multi-sheet `.xlsx → .csv`/`.tsv` automatically exports one file per sheet inside `<name>_sheets/` so zero sheets are lost. |
| **Text & Markup** | `.txt`, `.md`, `.html`, `.htm` | `.pdf`, `.docx`, `.html`, `.md`, `.txt` | Pure offline HTML/Markdown → PDF rendering via `PyMuPDF` `Story`. |
| **Structured Data** | `.csv`, `.tsv`, `.json`, `.yaml`, `.yml` | `.xlsx`, `.csv`, `.tsv`, `.json`, `.yaml` | UTF-8 safe with rectangular validation before spreadsheet/CSV export. |
| **Images** | `.png`, `.jpg`, `.jpeg`, `.webp`, `.heic`, `.heif`, `.avif`, `.bmp`, `.ico`, `.tiff`, `.tif` | `.png`, `.jpg`, `.webp`, `.avif`, `.heic`, `.tiff`, `.bmp`, `.ico`, `.pdf` | Automatic EXIF orientation transpose, ICC color profile forwarding, transparent square-canvas `.ico` centering, and multi-page `.tiff` support. |
| **Video** | `.mp4`, `.mkv`, `.mov`, `.webm`, `.avi`, `.m4v`, `.wmv`, `.gif` | `.mp4`, `.mkv`, `.mov`, `.webm`, `.avi`, `.m4v`, `.gif`, `.mp3`, `.wav`, `.flac`, `.m4a`, `.aac`, `.ogg`, `.opus` | Container-matched codecs (`libx264` + `yuv420p` + `faststart`, `libvpx-vp9` + `libopus`, two-step palettegen GIF). `.gif` only exposes video targets. |
| **Audio** | `.mp3`, `.wav`, `.flac`, `.m4a`, `.aac`, `.ogg`, `.opus`, `.wma` | `.mp3`, `.wav`, `.flac`, `.m4a`, `.aac`, `.ogg`, `.opus` | Direct audio-only FFmpeg transcoding (`libmp3lame`, `pcm_s16le`, `flac`, `aac`, `libvorbis`, `libopus`). |

---

## How to Use

### Option 1: Standalone Windows Installer (Recommended — No Python Required)

1. Download **`QuickConvertSetup.exe`** from the **[Releases](../../releases)** page.
2. Double-click `QuickConvertSetup.exe` and click **Install QuickConvert**.
   - Installs to `%LOCALAPPDATA%\Programs\QuickConvert` (no UAC / Admin prompt).
   - Registers cascading `QuickConvert` submenus for all 47 file types under `HKCU\Software\Classes\SystemFileAssociations`.
3. **Convert any file:**
   - Right-click one or more files in File Explorer.
   - On **Windows 11**, click **Show more options** (or press **`Shift + Right-Click`**).
   - Hover over **QuickConvert** and choose a prioritized target format (up to 10 quick actions per type to stay below Explorer's 16-item cap), or click **More Formats...** to open the categorized Format Picker dialog.
   - The converted file (or `<name>_pages/` / `<name>_sheets/` folder) is saved right next to the original file without overwriting existing outputs.
4. **Uninstall cleanly anytime:**
   - Open **Windows Settings → Apps → Installed apps → QuickConvert → Uninstall**, or re-run `QuickConvertSetup.exe` and click **Uninstall**. Every HKCU Registry key and installed file is cleanly removed.

---

### Option 2: Run from Source (Developers)

#### Prerequisites
- **Windows 10 or 11** with **Python 3.12+** (including `tkinter` and `pythonw.exe`).
- *Optional (for Office docs):* Microsoft Office (Word / PowerPoint / Excel) and/or LibreOffice (`soffice`). FFmpeg is bundled automatically via `imageio-ffmpeg` (system `ffmpeg` on `PATH` is used first if installed).

#### Setup & Register Context Menus

```powershell
git clone https://github.com/<your-username>/<your-repo>.git
cd <your-repo>

python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt

# Preview all HKCU Registry entries without writing anything
.\.venv\Scripts\python.exe install_menu.py --dry-run

# Register cascading Windows right-click context menus under HKCU (no Admin required)
.\.venv\Scripts\python.exe install_menu.py
```

#### Uninstall Context Menus (Source Mode)

```powershell
.\.venv\Scripts\python.exe uninstall_menu.py
```

#### Command-Line Usage

```powershell
# Convert with the Windows 11 Fluent progress window (silent via pythonw.exe)
.\.venv\Scripts\pythonw.exe main.py --target png "C:\path\to\report.docx"

# Open the Fluent Format Picker dialog for one or more files
.\.venv\Scripts\pythonw.exe main.py --choose-target "C:\path\to\document.pdf"

# Headless CLI conversion (ideal for scripting & batch jobs)
.\.venv\Scripts\python.exe main.py --no-ui --target webp --output-dir "C:\out" "C:\path\to\photo.png"

# Preview all 4 Fluent ProgressWindow UI states (indeterminate, determinate, completed, error)
.\.venv\Scripts\pythonw.exe main.py --preview-ui
```

---

## Configuration (`config.json`)

Default settings live in `config.json` (in the repo root or `%LOCALAPPDATA%\Programs\QuickConvert\config.json` when installed):

| Key | Default | Description |
| --- | --- | --- |
| `image_dpi` | `200` | Render DPI for PDF → PNG/JPG/WEBP/TIFF/PPTX and Image → PDF. |
| `jpeg_quality` | `90` | JPEG/HEIC output quality (`1–100`). |
| `webp_quality` | `85` | WebP/AVIF output quality (`1–100`). |
| `ffmpeg_preset` | `"medium"` | H.264 encoding preset (`ultrafast` .. `veryslow`) for `.mp4`/`.mkv`/`.mov`/`.m4v`/`.avi`. |
| `ffmpeg_crf` | `23` | Constant Rate Factor for H.264 and VP9 (`.webm`) targets. |
| `audio_bitrate` | `"192k"` | Bitrate for lossy audio targets (`.mp3`, `.m4a`, `.aac`, `.ogg`, `.opus`). |
| `ffmpeg_path` | `""` | Optional explicit path to `ffmpeg.exe` (overrides `PATH` and bundled `imageio-ffmpeg`). |
| `soffice_path` | `""` | Optional explicit path to LibreOffice `soffice.exe`. |
| `auto_close_delay_ms` | `1800` | Delay before the progress window auto-closes on success (paused while hovered). |

---

## Building `QuickConvertSetup.exe` & Automated Releases

To build the standalone Windows installer locally:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-build.txt
.\.venv\Scripts\python.exe build_installer.py
```

This creates:
- `dist\QuickConvert\QuickConvert.exe` — fast `--onedir --windowed` standalone app bundle.
- `dist\QuickConvertSetup.exe` — single-file no-Admin GUI installer embedding the compressed bundle and pre-generated Registry manifest.

Pushing a Git tag matching `v*` (for example, `git tag v1.0.0 && git push origin v1.0.0`) automatically runs `.github/workflows/release.yml` on `windows-latest` to build and publish `QuickConvertSetup.exe` as a GitHub Release asset.

---

## Running Automated Tests

```powershell
.\.venv\Scripts\python.exe tests\test_pipeline.py
```

The end-to-end suite programmatically generates fixtures (`.pdf`, 2-page `.docx`, multi-sheet `.xlsx`, `.csv`, `.json`, `.md`, RGBA `.png`, and `.mp4` with audio), verifies single-hop and multi-hop conversions, collision auto-incrementing, post-conversion validation, guaranteed temp directory cleanup on success and failure, Dijkstra routing safety guards, custom 7-frame `.ico` integrity, Explorer's `<= 12` submenu cap, and full `uninstall_menu.py` coverage.

---

## Project Structure

```text
QuickConvert/
├── main.py                         # CLI & silent GUI entrypoint (--target, --choose-target, --no-ui, --preview-ui)
├── install_menu.py                 # HKCU-only cascading context menu installer (no Admin)
├── uninstall_menu.py               # Clean HKCU Registry uninstaller with empty-key pruning
├── installer_gui.py                # Standalone Windows 11 Fluent GUI setup & uninstaller
├── build_installer.py              # PyInstaller build script for QuickConvertSetup.exe
├── config.json                     # User-configurable conversion defaults
├── requirements.txt                # Runtime offline dependencies
├── requirements-build.txt          # Build dependencies (+ PyInstaller)
├── assets/
│   ├── generate_icon.py            # 1024px supersampled Fluent icon generator
│   └── quickconvert.ico            # Multi-resolution icon (16..256 px)
├── quickconvert/
│   ├── config.py                   # Config dataclass & JSON loader
│   ├── paths.py                    # Source & frozen-bundle path resolution
│   ├── utils.py                    # Cancellation, atomic reservation/commit, robust cleanup
│   ├── graph.py                    # Weighted Dijkstra router & Explorer priority caps
│   ├── validator.py                # Mandatory post-conversion output integrity checks
│   ├── pipeline.py                 # Staged multi-step pipeline with finally cleanup
│   ├── converters/
│   │   ├── documents.py            # Word / PowerPoint / Excel COM + LibreOffice fallback
│   │   ├── pdf_converter.py        # PDF -> DOCX, PNG/JPG/WEBP, TIFF, PPTX, TXT, HTML, MD
│   │   ├── data_text.py            # Markup/text, multi-sheet XLSX, CSV/TSV/JSON/YAML
│   │   ├── images.py               # Pillow + pillow-heif (HEIC/AVIF/EXIF/ICC/ICO/multi-TIFF)
│   │   └── media.py                # Universal FFmpeg video & audio converter
│   └── ui/
│       ├── progress_window.py      # Native Windows 11 Fluent floating progress window
│       └── format_picker.py        # Categorized "More Formats..." Fluent picker dialog
├── tests/
│   └── test_pipeline.py            # Automated end-to-end verification suite
└── .github/workflows/
    └── release.yml                 # Automated Windows Release build for QuickConvertSetup.exe
```
