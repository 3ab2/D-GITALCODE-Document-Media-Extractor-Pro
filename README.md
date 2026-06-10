<div align="center">

<img src="resources/icons/logo.png" alt="D-GITALCODE logo" width="120"/>

# D-GITALCODE Document Media Extractor Pro

**A professional desktop application that extracts every image from Word, PowerPoint and PDF documents — organized, deduplicated and fully reported.**

[![Version](https://img.shields.io/badge/version-2.0.0-2ecc71?style=flat-square)](CHANGELOG.md)
[![Python](https://img.shields.io/badge/python-3.11%2B-3776AB?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue?style=flat-square)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey?style=flat-square)]()
[![GUI](https://img.shields.io/badge/GUI-CustomTkinter-1f6feb?style=flat-square)](https://github.com/TomSchimansky/CustomTkinter)

[Features](#-features) • [Screenshots](#-screenshots) • [Installation](#-installation) • [Usage](#-usage) • [Architecture](#-architecture) • [Roadmap](#-roadmap)

</div>

---

## ✨ Features

| Feature | Description |
| ------- | ----------- |
| 🗂 **Multi-format extraction** | DOCX, DOC, PPTX, PPT and PDF — the right extractor is selected automatically |
| 🧲 **Drag & drop** | Drop documents (or whole selections) straight onto the window |
| 📁 **Per-document output** | Each source file gets its own self-contained results folder — images, reports, everything |
| 🧹 **Smart deduplication** | SHA-256 hashing detects duplicate images across the whole batch (move / skip / keep — your choice) |
| 🖼 **Image gallery** | Searchable thumbnail gallery of every extracted image, filterable by name, format or source |
| 📊 **Statistics dashboard** | Files processed, images extracted, duplicates, processing time and a live format chart |
| 📄 **Professional reports** | Branded TXT, PDF, Excel and image-metadata reports generated per document |
| 🕘 **Extraction history** | Every run is recorded locally and can be reopened in one click |
| ⚙️ **Settings panel** | Output folder, theme, language, duplicate handling and report formats — saved automatically |
| 🌍 **Multilingual** | English, Français, العربية |
| 🌗 **Dark / light themes** | Minimal, professional design system with one-click theme toggle |
| 📦 **One-click EXE build** | `build_exe.bat` produces a single windowed executable via PyInstaller |

## 📑 Supported Formats

| Type | Extension | Supported | Extraction method |
| ---- | --------- | :-------: | ----------------- |
| Word (modern) | `.docx` | ✅ | OOXML archive parsing |
| Word (legacy) | `.doc` | ✅ | Binary signature carving |
| PowerPoint (modern) | `.pptx` | ✅ | OOXML archive parsing |
| PowerPoint (legacy) | `.ppt` | ✅ | Binary signature carving |
| PDF | `.pdf` | ✅ | PyMuPDF image objects |

Extracted image formats: **PNG · JPG · GIF · BMP · TIFF · WEBP · EMF · WMF**

## 📸 Screenshots

| Dashboard | Extraction View |
| :-------: | :-------------: |
| ![Dashboard](docs/screenshots/dashboard.png) | ![Extraction](docs/screenshots/extraction.png) |

| Reports & Gallery | Settings |
| :---------------: | :------: |
| ![Reports](docs/screenshots/reports.png) | ![Settings](docs/screenshots/settings.png) |

## 🏛 Architecture

The project follows a strict layered architecture with full separation of concerns:

```text
extractor-images-from-docx/
├── main.py                  # Entry point
├── config.py                # Metadata, paths, formats, translations
├── gui/                     # ── GUI LAYER ──
│   ├── app.py               #    Main window, navigation, threading glue
│   ├── pages.py             #    Dashboard / Extract / History / Reports / Settings
│   ├── components.py        #    Reusable widgets (sidebar, cards, tables…)
│   └── theme.py             #    Design system: palettes, spacing, typography
├── core/                    # ── CORE LAYER ──
│   ├── extractor.py         #    BaseExtractor + Word / PowerPoint / PDF extractors
│   ├── image_utils.py       #    Validation, format detection, conversion
│   └── file_handler.py      #    Safe file I/O and directory traversal
├── services/                # ── SERVICES LAYER ──
│   ├── batch_service.py     #    Queued, threaded multi-file processing
│   ├── report_service.py    #    TXT / PDF / XLSX / metadata generation
│   ├── settings_service.py  #    Persistent user settings (settings.json)
│   ├── history_service.py   #    Local extraction history (history.json)
│   └── logging_service.py   #    Central logging (file + console + UI)
├── utils/                   # ── UTILITIES LAYER ──
│   ├── hash_utils.py        #    SHA-256 hashing & duplicate detection
│   ├── format_utils.py      #    Size / duration / timestamp formatting
│   └── path_utils.py        #    Safe paths, sanitization, unique names
└── resources/               # Logo, icons, language files (EN/FR/AR)
```

| Layer | Responsibility | Depends on |
| ----- | -------------- | ---------- |
| **GUI** | Presentation only — pages, widgets, theming | Services, Config |
| **Core** | Pure extraction logic — no UI imports | Utils, Config |
| **Services** | Orchestration: batching, reporting, persistence, logging | Core, Utils |
| **Utilities** | Stateless helpers: hashing, formatting, paths | Config |

## 🚀 Installation

> Requires **Python 3.11+**

```bash
# 1. Clone the repository
git clone https://github.com/3ab2/extractor-images-from-docx.git
cd extractor-images-from-docx

# 2. Create a virtual environment
python -m venv venv

# 3. Activate it
venv\Scripts\activate          # Windows
source venv/bin/activate       # macOS / Linux

# 4. Install dependencies
pip install -r requirements.txt

# 5. Launch
python main.py
```

### Build a standalone Windows executable

```bat
build_exe.bat
```

Produces `dist\D-GITALCODE-Document-Media-Extractor-Pro.exe` — single file, windowed, branded icon, no console.

## 📖 Usage

1. **Add documents** — drag & drop files onto the window, or use **Browse files** / **Select folder** on the *Extract Images* page.
2. **Review the queue** — the file table shows name, type and size for every document before extraction.
3. **Start extraction** — click **Start extraction**; progress, current file and a live activity log keep you informed.
4. **Explore the results** — the *Reports* page opens automatically with a searchable thumbnail gallery and direct buttons to each document's output folder.
5. **Check the dashboard** — totals, duplicates, processing time and the format distribution chart.
6. **Revisit any run** — the *History* page lists previous extractions and reopens their output folders.

## 🗃 Output Structure

Every source document gets its own self-contained folder — only the formats that were actually found are created:

```text
Extracted_Images_D-GITALCODE/
├── Annual_Report/                 # ← one folder per source document
│   ├── PNG/
│   │   ├── image_0001.png
│   │   └── image_0002.png
│   ├── JPG/
│   │   └── image_0003.jpg
│   ├── DUPLICATES/                # duplicates (when mode = separate)
│   ├── Extraction_Report.txt
│   ├── PDF_Report.pdf
│   ├── Excel_Report.xlsx
│   └── Metadata.xlsx
└── Product_Brief/
    ├── GIF/
    │   └── image_0001.gif
    ├── Extraction_Report.txt
    ├── PDF_Report.pdf
    ├── Excel_Report.xlsx
    └── Metadata.xlsx
```

## 📄 Reports

| Report | File | Contents |
| ------ | ---- | -------- |
| **TXT summary** | `Extraction_Report.txt` | Human-readable summary: totals, format breakdown, duplicates, timing |
| **PDF report** | `PDF_Report.pdf` | Branded, watermarked report suitable for sharing and archiving |
| **Excel report** | `Excel_Report.xlsx` | Structured workbook: summary sheet + per-image detail sheet |
| **Image metadata** | `Metadata.xlsx` | Source file, image name, format, width, height, size (KB), SHA-256 hash, output path |

Each report is generated **per document** inside its output folder, and every format can be toggled in *Settings → Reports*.

## 🗺 Roadmap

- [ ] Export gallery selection as ZIP archive
- [ ] OCR text layer detection for scanned PDFs
- [ ] Extraction of embedded video and audio media
- [ ] Command-line interface (headless batch mode)
- [ ] ODT / ODP (OpenDocument) support
- [ ] Cloud storage targets (Google Drive, OneDrive)
- [ ] Automatic update checks

## 🤝 Contributing

Contributions are welcome!

1. **Fork** the repository
2. **Create a branch** — `git checkout -b feat/amazing-feature`
3. **Code** — follow PEP 8, add type hints, keep layers decoupled (no GUI imports in `core/`)
4. **Test** — verify extraction on DOCX, PPTX and PDF samples
5. **Commit** — use conventional commits (`feat:`, `fix:`, `docs:`…)
6. **Open a Pull Request** with a clear description

Found a bug? [Open an issue](https://github.com/3ab2/extractor-images-from-docx/issues) with steps to reproduce.

## 📜 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.

## 👤 Author

<div align="center">

**ELKARCH ABDELHAMID**

Brand: **D-GITALCODE**

[![GitHub](https://img.shields.io/badge/GitHub-3ab2%2Fextractor--images--from--docx-181717?style=flat-square&logo=github)](https://github.com/3ab2/extractor-images-from-docx)

*D-GITALCODE © 2026 — Professional Document Media Extraction*

</div>
