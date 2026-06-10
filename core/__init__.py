"""
D-GITALCODE Document Media Extractor Pro - Core package.

Author: ELKARCH ABDELHAMID | Brand: D-GITALCODE
"""

from core.extractor import (
    BaseExtractor,
    ExtractionResult,
    ExtractionSummary,
    PDFExtractor,
    PowerPointExtractor,
    WordExtractor,
    create_extractor,
    supported_extensions,
)
from core.file_handler import FileHandler
from core.image_utils import ImageInfo, detect_image_format, is_valid_image

__all__ = [
    "BaseExtractor",
    "ExtractionResult",
    "ExtractionSummary",
    "PDFExtractor",
    "PowerPointExtractor",
    "WordExtractor",
    "create_extractor",
    "supported_extensions",
    "FileHandler",
    "ImageInfo",
    "detect_image_format",
    "is_valid_image",
]
