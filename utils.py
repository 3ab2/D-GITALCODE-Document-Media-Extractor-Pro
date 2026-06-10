"""
================================================================================
D-GITALCODE Word Image Extractor Pro - Utilities Module
================================================================================

Author: ELKARCH ABDELHAMID
Brand: D-GITALCODE
Version: 2.0.0
Python: 3.11+

This module provides utility functions for hashing, image detection,
file operations, and common helpers.

================================================================================
"""

import hashlib
import logging
from pathlib import Path
from typing import Dict, Optional, Tuple
from PIL import Image
import io


# ==================== CONSTANTS ====================

SUPPORTED_FORMATS = {
    "JPG": (".jpg", ".jpeg"),
    "PNG": (".png",),
    "GIF": (".gif",),
    "BMP": (".bmp",),
    "TIFF": (".tiff", ".tif"),
    "WEBP": (".webp",),
    "SVG": (".svg",),
}

FORMAT_EXTENSIONS = {
    ".jpg": "JPG",
    ".jpeg": "JPG",
    ".png": "PNG",
    ".gif": "GIF",
    ".bmp": "BMP",
    ".tiff": "TIFF",
    ".tif": "TIFF",
    ".webp": "WEBP",
    ".svg": "SVG",
}

# Supported file extensions
SUPPORTED_DOC_FORMATS = (".docx", ".doc")

# UI Colors
COLORS = {
    "DARK_BG": "#1a1a1a",
    "SECONDARY_BG": "#2a2a2a",
    "ACCENT_COLOR": "#00a8e8",
    "TEXT_COLOR": "#ffffff",
    "SUCCESS_COLOR": "#00c853",
    "ERROR_COLOR": "#ff5252",
    "WARNING_COLOR": "#ffb300",
    "LIGHT_BG": "#f5f5f5",
    "LIGHT_SECONDARY": "#e0e0e0",
    "LIGHT_TEXT": "#1a1a1a",
}

# Language translations
LANGUAGES = {
    "EN": {
        "app_title": "D-GITALCODE Word Image Extractor Pro",
        "select_file": "Select Word Document",
        "select_folder": "Select Folder",
        "start_extraction": "Start Extraction",
        "batch_process": "Batch Process",
        "progress": "Progress",
        "status": "Status",
        "images_extracted": "Images Extracted",
        "total_images": "Total Images",
        "format": "Format",
        "duplicates": "Duplicates",
        "processing_time": "Processing Time",
        "output_folder": "Output Folder",
        "extraction_complete": "Extraction Complete",
        "generating_report": "Generating Report",
        "error": "Error",
        "success": "Success",
        "open_output": "Open Output Folder",
        "clear": "Clear",
        "dark_mode": "Dark Mode",
        "language": "Language",
        "no_file_selected": "No file selected",
        "file_selected": "File selected",
        "processing": "Processing",
        "ready": "Ready",
    },
    "FR": {
        "app_title": "D-GITALCODE Extracteur d'Images Word Pro",
        "select_file": "Sélectionner un Document Word",
        "select_folder": "Sélectionner un Dossier",
        "start_extraction": "Commencer l'Extraction",
        "batch_process": "Traitement par Lot",
        "progress": "Progrès",
        "status": "État",
        "images_extracted": "Images Extraites",
        "total_images": "Nombre Total d'Images",
        "format": "Format",
        "duplicates": "Doublons",
        "processing_time": "Temps de Traitement",
        "output_folder": "Dossier de Sortie",
        "extraction_complete": "Extraction Terminée",
        "generating_report": "Génération du Rapport",
        "error": "Erreur",
        "success": "Succès",
        "open_output": "Ouvrir le Dossier de Sortie",
        "clear": "Effacer",
        "dark_mode": "Mode Sombre",
        "language": "Langue",
        "no_file_selected": "Aucun fichier sélectionné",
        "file_selected": "Fichier sélectionné",
        "processing": "Traitement en cours",
        "ready": "Prêt",
    },
    "AR": {
        "app_title": "D-GITALCODE مستخرج صور Word Pro",
        "select_file": "اختر مستند Word",
        "select_folder": "اختر مجلد",
        "start_extraction": "ابدأ الاستخراج",
        "batch_process": "معالجة دفعية",
        "progress": "التقدم",
        "status": "الحالة",
        "images_extracted": "الصور المستخرجة",
        "total_images": "إجمالي الصور",
        "format": "الصيغة",
        "duplicates": "التكرارات",
        "processing_time": "وقت المعالجة",
        "output_folder": "مجلد الإخراج",
        "extraction_complete": "اكتمل الاستخراج",
        "generating_report": "توليد التقرير",
        "error": "خطأ",
        "success": "نجح",
        "open_output": "فتح مجلد الإخراج",
        "clear": "مسح",
        "dark_mode": "الوضع الليلي",
        "language": "اللغة",
        "no_file_selected": "لم يتم تحديد أي ملف",
        "file_selected": "تم تحديد الملف",
        "processing": "قيد المعالجة",
        "ready": "جاهز",
    },
}


# ==================== LOGGING ====================

class LoggerSetup:
    """Setup and manage application logging"""

    _logger = None

    @staticmethod
    def setup_logger(log_dir: Path) -> logging.Logger:
        """Initialize logger with file and console handlers"""
        if LoggerSetup._logger:
            return LoggerSetup._logger

        logger = logging.getLogger("D-GITALCODE")
        logger.setLevel(logging.DEBUG)

        # Create log directory
        log_dir = Path(log_dir)
        log_dir.mkdir(parents=True, exist_ok=True)

        # File handler
        log_file = log_dir / "extraction.log"
        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setLevel(logging.DEBUG)

        # Console handler
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)

        # Formatter
        formatter = logging.Formatter(
            "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        fh.setFormatter(formatter)
        ch.setFormatter(formatter)

        # Clear existing handlers
        logger.handlers.clear()

        logger.addHandler(fh)
        logger.addHandler(ch)

        LoggerSetup._logger = logger
        return logger

    @staticmethod
    def get_logger() -> logging.Logger:
        """Get existing logger"""
        if LoggerSetup._logger is None:
            return LoggerSetup.setup_logger(Path.home() / ".D-GITALCODE" / "logs")
        return LoggerSetup._logger


# ==================== HASHING & DEDUPLICATION ====================

class HashUtils:
    """Utility functions for image hashing and deduplication"""

    @staticmethod
    def calculate_hash(data: bytes, algorithm: str = "sha256") -> str:
        """
        Calculate hash of image data

        Args:
            data: Image bytes
            algorithm: Hash algorithm (sha256, md5)

        Returns:
            Hex digest of hash
        """
        if algorithm == "sha256":
            return hashlib.sha256(data).hexdigest()
        elif algorithm == "md5":
            return hashlib.md5(data).hexdigest()
        else:
            return hashlib.sha256(data).hexdigest()

    @staticmethod
    def is_duplicate(image_hash: str, known_hashes: set) -> bool:
        """Check if image hash is already in known hashes"""
        return image_hash in known_hashes

    @staticmethod
    def add_hash(image_hash: str, known_hashes: set):
        """Add hash to known hashes"""
        known_hashes.add(image_hash)


# ==================== IMAGE DETECTION ====================

class ImageDetector:
    """Detect and classify image formats"""

    @staticmethod
    def detect_format(image_data: bytes) -> Optional[Tuple[str, str]]:
        """
        Detect image format from bytes

        Args:
            image_data: Raw image bytes

        Returns:
            Tuple of (extension, format_folder) or None
        """
        try:
            img = Image.open(io.BytesIO(image_data))
            img_format = img.format

            format_map = {
                "JPEG": (".jpg", "JPG"),
                "PNG": (".png", "PNG"),
                "GIF": (".gif", "GIF"),
                "BMP": (".bmp", "BMP"),
                "TIFF": (".tiff", "TIFF"),
                "WEBP": (".webp", "WEBP"),
                "SVG": (".svg", "SVG"),
            }

            return format_map.get(img_format, None)

        except Exception as e:
            logger = LoggerSetup.get_logger()
            logger.warning(f"Could not detect image format: {str(e)}")
            return None

    @staticmethod
    def get_image_extension_by_format(pil_format: str) -> Optional[str]:
        """Get file extension from PIL format name"""
        format_map = {
            "JPEG": ".jpg",
            "PNG": ".png",
            "GIF": ".gif",
            "BMP": ".bmp",
            "TIFF": ".tiff",
            "WEBP": ".webp",
        }
        return format_map.get(pil_format, None)


# ==================== FILE OPERATIONS ====================

class FileUtils:
    """File system utilities"""

    @staticmethod
    def create_folder_structure(output_dir: Path) -> bool:
        """Create output folder structure with all format subfolders"""
        try:
            output_dir.mkdir(parents=True, exist_ok=True)

            folders = list(SUPPORTED_FORMATS.keys()) + ["DUPLICATES"]
            for folder in folders:
                (output_dir / folder).mkdir(parents=True, exist_ok=True)

            return True

        except Exception as e:
            logger = LoggerSetup.get_logger()
            logger.error(f"Failed to create folder structure: {str(e)}")
            return False

    @staticmethod
    def get_unique_filename(file_path: Path) -> Path:
        """Generate unique filename if file already exists"""
        if not file_path.exists():
            return file_path

        stem = file_path.stem
        suffix = file_path.suffix
        parent = file_path.parent

        counter = 1
        while True:
            new_name = f"{stem}_{counter}{suffix}"
            new_path = parent / new_name
            if not new_path.exists():
                return new_path
            counter += 1

    @staticmethod
    def get_file_size_mb(file_path: Path) -> float:
        """Get file size in MB"""
        return file_path.stat().st_size / (1024 * 1024)

    @staticmethod
    def get_folder_size_mb(folder_path: Path) -> float:
        """Get total folder size in MB"""
        total = 0
        for file_path in folder_path.rglob("*"):
            if file_path.is_file():
                total += file_path.stat().st_size
        return total / (1024 * 1024)


# ==================== TRANSLATION ====================

class Translator:
    """Handle application translations"""

    _current_language = "EN"

    @staticmethod
    def set_language(lang_code: str):
        """Set current language"""
        if lang_code in LANGUAGES:
            Translator._current_language = lang_code

    @staticmethod
    def get_text(key: str, language: str = None) -> str:
        """Get translated text"""
        lang = language or Translator._current_language
        if lang not in LANGUAGES:
            lang = "EN"

        return LANGUAGES[lang].get(key, key)

    @staticmethod
    def get_available_languages() -> list:
        """Get list of available languages"""
        return list(LANGUAGES.keys())


# ==================== STATISTICS ====================

class StatisticsCollector:
    """Collect and manage extraction statistics"""

    def __init__(self):
        """Initialize statistics"""
        self.stats = {
            "total_images": 0,
            "images_by_format": {fmt: 0 for fmt in SUPPORTED_FORMATS.keys()},
            "duplicates_found": 0,
            "files_processed": 0,
            "files_failed": 0,
            "total_size_mb": 0,
            "extraction_time_seconds": 0,
            "timestamp": None,
        }

    def increment_format(self, format_name: str):
        """Increment count for a format"""
        if format_name in self.stats["images_by_format"]:
            self.stats["images_by_format"][format_name] += 1
            self.stats["total_images"] += 1

    def increment_duplicates(self, count: int = 1):
        """Increment duplicate counter"""
        self.stats["duplicates_found"] += count

    def increment_files_processed(self):
        """Increment files processed counter"""
        self.stats["files_processed"] += 1

    def increment_files_failed(self):
        """Increment files failed counter"""
        self.stats["files_failed"] += 1

    def set_extraction_time(self, seconds: float):
        """Set extraction time"""
        self.stats["extraction_time_seconds"] = seconds

    def set_total_size(self, mb: float):
        """Set total processed size"""
        self.stats["total_size_mb"] = mb

    def set_timestamp(self, timestamp):
        """Set timestamp"""
        self.stats["timestamp"] = timestamp

    def get_stats(self) -> Dict:
        """Get all statistics"""
        return self.stats.copy()

    def get_processing_speed(self) -> float:
        """Get processing speed in MB/s"""
        if self.stats["extraction_time_seconds"] > 0:
            return self.stats["total_size_mb"] / self.stats["extraction_time_seconds"]
        return 0


# ==================== VALIDATION ====================

class ValidationUtils:
    """Validate files and documents"""

    @staticmethod
    def is_valid_docx(file_path: Path) -> bool:
        """Validate if file is a valid DOCX"""
        try:
            from docx import Document

            if file_path.suffix.lower() != ".docx":
                return False

            doc = Document(file_path)
            return True

        except Exception:
            return False

    @staticmethod
    def is_valid_doc(file_path: Path) -> bool:
        """Check if file is a DOC file"""
        return file_path.suffix.lower() == ".doc"

    @staticmethod
    def is_supported_file(file_path: Path) -> bool:
        """Check if file is supported"""
        return file_path.suffix.lower() in SUPPORTED_DOC_FORMATS
