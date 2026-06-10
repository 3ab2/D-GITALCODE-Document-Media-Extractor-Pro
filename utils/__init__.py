"""
D-GITALCODE Document Media Extractor Pro - Utils package.

Author: ELKARCH ABDELHAMID | Brand: D-GITALCODE
"""

from utils.hash_utils import DuplicateDetector, compute_hash
from utils.format_utils import bytes_to_mb, format_duration, format_timestamp
from utils.path_utils import (
    ensure_directory,
    open_in_file_explorer,
    sanitize_filename,
    unique_path,
)

__all__ = [
    "DuplicateDetector",
    "compute_hash",
    "bytes_to_mb",
    "format_duration",
    "format_timestamp",
    "ensure_directory",
    "open_in_file_explorer",
    "sanitize_filename",
    "unique_path",
]
