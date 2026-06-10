# Changelog

All notable changes to **D-GITALCODE Document Media Extractor Pro** are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [2.0.0] - 2026-06-10

### Major Release — D-GITALCODE Document Media Extractor Pro

The application has been renamed from *Word Image Extractor Pro* to
**Document Media Extractor Pro** to reflect its expanded multi-format scope,
and fully refactored into a professional layered architecture.

### Added

- **DOC support** — legacy Word documents via binary signature carving
- **PPTX / PPT support** — modern and legacy PowerPoint presentations
- **PDF support** — image extraction powered by PyMuPDF
- **Automatic extractor detection** — `BaseExtractor` factory routes each file
  to `WordExtractor`, `PowerPointExtractor` or `PDFExtractor` by extension
- **Modular architecture** — clean separation into `gui/`, `core/`, `services/`
  and `utils/` layers with no GUI logic in core
- **Per-document reports** — TXT, PDF and Excel reports generated inside each
  source document's own output folder
- **Metadata export** — `Metadata.xlsx` with image name, format, dimensions,
  size (KB), SHA-256 hash and output path
- **Drag & drop** — drop documents directly onto the application window
- **Image gallery** — searchable thumbnail gallery (name, format, source)
- **Statistics dashboard** — stat cards and format-distribution chart
- **Extraction history** — local run history with one-click reopening
- **Settings panel** — output folder, theme, language, duplicate handling and
  report formats, persisted automatically
- **EXE build script** — `build_exe.bat` produces a single windowed executable
- **File preview table** — name, type, size and live status before extraction

### Improved

- **UI/UX** — complete redesign: sidebar navigation, dedicated pages, minimal
  professional design system, dark/light themes, multilingual UI (EN/FR/AR)
- **Performance** — threaded batch pipeline with a thread-safe UI queue keeps
  the interface responsive on large batches
- **Batch processing** — queued multi-file processing with per-file progress,
  cancellation and a detailed end-of-run summary
- **Reporting system** — branded, standardized reports with per-file format
  breakdown and selectable output formats

### Fixed

- **Empty format folders** — format subfolders (PNG/, JPG/, …) are now created
  only when images of that format are actually extracted
- **Output organization** — each source document gets a self-contained output
  folder; results from different documents are never mixed
- **Report isolation** — reports describe exactly one source document and live
  alongside its images instead of a shared global report

---

## [1.0.0] - 2024-06-09

### Initial Release

- Single-file desktop GUI for extracting images from `.docx` documents
- Automatic image format detection (JPG, PNG, GIF, BMP, TIFF, WEBP)
- Basic duplicate detection and text report generation
- Real-time progress bar

---

<div align="center">

**D-GITALCODE © 2026 — Developed by ELKARCH ABDELHAMID**

[Repository](https://github.com/3ab2/extractor-images-from-docx) · [Issues](https://github.com/3ab2/extractor-images-from-docx/issues)

</div>
