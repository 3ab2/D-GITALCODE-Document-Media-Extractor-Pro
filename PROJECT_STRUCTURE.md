# Project Structure

**D-GITALCODE Document Media Extractor Pro** — Modular Architecture (v2.0.0)

Author: **ELKARCH ABDELHAMID** | Brand: **D-GITALCODE**

```
extractor-images-from-docx/
│
├── main.py                      # Entry point: init config, logging, launch GUI
├── config.py                    # Metadata, paths, formats, Translator
│
├── build_exe.bat                # PyInstaller single-EXE build (icon, no console)
│
├── gui/                         # ── Presentation layer ──
│   ├── __init__.py
│   ├── app.py                   # Application shell: sidebar navigation, header,
│   │                            #   drag & drop, thread-safe UI queue, wiring
│   ├── components.py            # Widget library: Sidebar, Header, StatusBar,
│   │                            #   Card, DropZone, FileTable, ProgressPanel, Log
│   ├── pages.py                 # Dashboard, Extract, History, Reports, Settings
│   └── theme.py                 # Design system: palette, spacing, typography
│
├── core/                        # ── Business logic ──
│   ├── __init__.py
│   ├── extractor.py             # BaseExtractor + Word/PowerPoint/PDF extractors,
│   │                            #   create_extractor factory, result/summary models
│   ├── image_utils.py           # detect_image_format, is_valid_image, convert
│   └── file_handler.py          # FileHandler: safe I/O, traversal, validation
│
├── services/                    # ── Application services ──
│   ├── __init__.py
│   ├── report_service.py        # ReportService: TXT / PDF / XLSX / metadata
│   ├── logging_service.py       # LoggingService: file + console + UI handlers
│   ├── batch_service.py         # BatchService: threaded queue processing
│   ├── settings_service.py      # SettingsService: auto-saved preferences
│   └── history_service.py       # HistoryService: local extraction history
│
├── utils/                       # ── Cross-cutting helpers ──
│   ├── __init__.py
│   ├── hash_utils.py            # compute_hash, DuplicateDetector
│   ├── format_utils.py          # size / duration / timestamp formatting
│   └── path_utils.py            # ensure_directory, unique_path, sanitize, open
│
├── resources/
│   ├── icons/                   # Application icons
│   └── languages/               # en.json, fr.json, ar.json translations
│
├── output/                      # Default extraction destination
├── logs/                        # extraction.log
│
├── requirements.txt
├── README.md
├── CHANGELOG.md
└── LICENSE
```

## Layer Responsibilities

| Layer | Rule |
|---|---|
| `gui/` | Presentation only. Components expose callbacks and update methods; no business logic. Worker-thread events are marshalled via `Tk.after`. |
| `core/` | Pure extraction logic. No GUI imports, fully testable in isolation. |
| `services/` | Orchestration: logging, threading/queueing, report generation. |
| `utils/` | Stateless helpers shared by all layers. |
| `config.py` | Single source of truth for metadata, paths and supported formats. |

## Data Flow

```
MainApplication (gui/app.py)
        │  documents + callbacks
        ▼
BatchService (worker thread, queue.Queue)
        │  per document: create_extractor(file type)
        ▼
BaseExtractor ── WordExtractor (.docx/.doc)
        │     ├─ PowerPointExtractor (.pptx/.ppt)
        │     └─ PDFExtractor (.pdf, PyMuPDF)
        │
        ├─► DuplicateDetector (SHA-256, shared across batch)
        ▼
FileHandler ──► output/Extracted_Images_D-GITALCODE/<Document_Name>/<FORMAT>/
        │
        ▼
Per-file ExtractionSummary ──► ReportService ──► <Document_Name>/
        │                          Extraction_Report.txt · PDF_Report.pdf
        │                          Excel_Report.xlsx · Metadata.xlsx
        ▼
Batch ExtractionSummary ──► GUI (dashboard, results, history, log)
```
