"""
================================================================================
D-GITALCODE Word Image Extractor
Professional Image Extraction Tool for Microsoft Word Documents
================================================================================

Author: ELKARCH ABDELHAMID
Brand: D-GITALCODE
Description: A sophisticated desktop application for extracting and organizing
             images from DOCX files with automatic classification and reporting.

Version: 1.0.0
Python: 3.11+
License: MIT

Requirements:
    - python-docx
    - Pillow
    - customtkinter

Installation:
    pip install python-docx Pillow customtkinter

================================================================================
"""

import os
import sys
import shutil
import logging
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple, Optional
import traceback
from enum import Enum

import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from PIL import Image
import io


# ==================== CONSTANTS AND ENUMERATIONS ====================

class ImageFormat(Enum):
    """Supported image formats"""
    JPG = (".jpg", "JPG")
    JPEG = (".jpeg", "JPG")
    PNG = (".png", "PNG")
    GIF = (".gif", "GIF")
    BMP = (".bmp", "BMP")
    TIFF = (".tiff", "TIFF")
    WEBP = (".webp", "WEBP")
    SVG = (".svg", "SVG")


SUPPORTED_FORMATS = ["JPG", "PNG", "GIF", "BMP", "TIFF", "WEBP", "SVG"]
OUTPUT_FOLDER_NAME = "Extracted_Images_D-GITALCODE"
REPORT_FILENAME = "Extraction_Report.txt"
LOG_FILENAME = "Extraction_Log.txt"

# Color scheme for modern UI
DARK_BG = "#1a1a1a"
SECONDARY_BG = "#2a2a2a"
ACCENT_COLOR = "#00a8e8"
TEXT_COLOR = "#ffffff"
SUCCESS_COLOR = "#00c853"
ERROR_COLOR = "#ff5252"
WARNING_COLOR = "#ffb300"


# ==================== LOGGING CONFIGURATION ====================

class LoggerSetup:
    """Configure and manage application logging"""
    
    @staticmethod
    def setup_logger(log_dir: Path) -> logging.Logger:
        """Initialize logger with file and console handlers"""
        logger = logging.getLogger("D-GITALCODE")
        logger.setLevel(logging.DEBUG)
        
        # File handler
        log_file = log_dir / LOG_FILENAME
        fh = logging.FileHandler(log_file, encoding='utf-8')
        fh.setLevel(logging.DEBUG)
        
        # Console handler
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        
        # Formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        fh.setFormatter(formatter)
        ch.setFormatter(formatter)
        
        logger.addHandler(fh)
        logger.addHandler(ch)
        
        return logger


logger = None  # Will be initialized later


# ==================== IMAGE EXTRACTION CORE ====================

class ImageExtractor:
    """Core logic for extracting and processing images from DOCX files"""
    
    def __init__(self, docx_path: Path, output_dir: Path, logger_instance):
        """
        Initialize the extractor
        
        Args:
            docx_path: Path to the DOCX file
            output_dir: Directory to save extracted images
            logger_instance: Logger instance
        """
        self.docx_path = Path(docx_path)
        self.output_dir = Path(output_dir)
        self.logger = logger_instance
        self.extraction_stats = {
            "total_images": 0,
            "images_by_format": {},
            "extraction_time": None,
            "timestamp": datetime.now()
        }
        
        # Initialize format counters
        for fmt in SUPPORTED_FORMATS:
            self.extraction_stats["images_by_format"][fmt] = 0
    
    def validate_docx(self) -> bool:
        """Validate that the selected file is a valid DOCX document"""
        try:
            if not self.docx_path.exists():
                self.logger.error(f"File not found: {self.docx_path}")
                return False
            
            if self.docx_path.suffix.lower() != '.docx':
                self.logger.error(f"Invalid file format: {self.docx_path.suffix}")
                return False
            
            # Try to open as DOCX
            doc = Document(self.docx_path)
            self.logger.info(f"Valid DOCX file: {self.docx_path.name}")
            return True
        
        except Exception as e:
            self.logger.error(f"Failed to validate DOCX: {str(e)}")
            return False
    
    def create_output_structure(self) -> bool:
        """Create the output folder structure"""
        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
            
            for fmt in SUPPORTED_FORMATS:
                format_dir = self.output_dir / fmt
                format_dir.mkdir(parents=True, exist_ok=True)
            
            self.logger.info(f"Output structure created at: {self.output_dir}")
            return True
        
        except Exception as e:
            self.logger.error(f"Failed to create output structure: {str(e)}")
            return False
    
    def _get_image_format(self, image_data: bytes) -> Optional[Tuple[str, str]]:
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
            
            if img_format == 'JPEG':
                return ('.jpg', 'JPG')
            elif img_format == 'PNG':
                return ('.png', 'PNG')
            elif img_format == 'GIF':
                return ('.gif', 'GIF')
            elif img_format == 'BMP':
                return ('.bmp', 'BMP')
            elif img_format == 'TIFF':
                return ('.tiff', 'TIFF')
            elif img_format == 'WEBP':
                return ('.webp', 'WEBP')
            else:
                self.logger.warning(f"Unknown image format: {img_format}")
                return None
        
        except Exception as e:
            self.logger.warning(f"Could not detect image format: {str(e)}")
            return None
    
    def extract_images(self, progress_callback=None) -> bool:
        """
        Extract all images from the DOCX file
        
        Args:
            progress_callback: Optional callback for progress updates
            
        Returns:
            True if extraction successful, False otherwise
        """
        try:
            start_time = datetime.now()
            
            if not self.validate_docx():
                return False
            
            if not self.create_output_structure():
                return False
            
            doc = Document(self.docx_path)
            image_counter = 0
            
            # Extract images from document parts
            for rel in doc.part.rels.values():
                if "image" in rel.target_ref:
                    try:
                        image_data = rel.target_part.blob
                        format_info = self._get_image_format(image_data)
                        
                        if format_info:
                            ext, fmt_folder = format_info
                            image_counter += 1
                            
                            # Generate filename
                            filename = f"image_{image_counter}{ext}"
                            filepath = self.output_dir / fmt_folder / filename
                            
                            # Save image
                            with open(filepath, 'wb') as f:
                                f.write(image_data)
                            
                            self.extraction_stats["total_images"] += 1
                            self.extraction_stats["images_by_format"][fmt_folder] += 1
                            
                            self.logger.info(f"Extracted: {filename} ({fmt_folder})")
                            
                            if progress_callback:
                                progress_callback(image_counter)
                    
                    except Exception as e:
                        self.logger.warning(f"Failed to extract image: {str(e)}")
                        continue
            
            end_time = datetime.now()
            self.extraction_stats["extraction_time"] = (end_time - start_time).total_seconds()
            
            self.logger.info(f"Extraction completed: {self.extraction_stats['total_images']} images extracted")
            return True
        
        except Exception as e:
            self.logger.error(f"Extraction failed: {str(e)}\n{traceback.format_exc()}")
            return False
    
    def generate_report(self) -> bool:
        """Generate extraction report"""
        try:
            report_path = self.output_dir / REPORT_FILENAME
            
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write("=" * 80 + "\n")
                f.write("D-GITALCODE WORD IMAGE EXTRACTION REPORT\n")
                f.write("=" * 80 + "\n\n")
                
                f.write(f"Author: ELKARCH ABDELHAMID\n")
                f.write(f"Brand: D-GITALCODE\n")
                f.write(f"Application: D-GITALCODE Word Image Extractor v1.0.0\n\n")
                
                f.write("EXTRACTION DETAILS:\n")
                f.write("-" * 80 + "\n")
                f.write(f"Extraction Date & Time: {self.extraction_stats['timestamp'].strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Source Document: {self.docx_path.name}\n")
                f.write(f"Document Path: {self.docx_path.absolute()}\n")
                f.write(f"Extraction Duration: {self.extraction_stats['extraction_time']:.2f} seconds\n\n")
                
                f.write("IMAGE STATISTICS:\n")
                f.write("-" * 80 + "\n")
                f.write(f"Total Images Extracted: {self.extraction_stats['total_images']}\n\n")
                
                f.write("Images by Format:\n")
                for fmt in SUPPORTED_FORMATS:
                    count = self.extraction_stats["images_by_format"][fmt]
                    if count > 0:
                        f.write(f"  {fmt:8s}: {count:4d} image(s)\n")
                
                f.write("\n" + "=" * 80 + "\n")
                f.write("OUTPUT CONFIGURATION:\n")
                f.write("-" * 80 + "\n")
                f.write(f"Destination Folder: {self.output_dir.absolute()}\n")
                f.write(f"Folder Structure:\n")
                f.write(f"  └── {OUTPUT_FOLDER_NAME}/\n")
                
                for fmt in SUPPORTED_FORMATS:
                    count = self.extraction_stats["images_by_format"][fmt]
                    if count > 0:
                        f.write(f"      ├── {fmt}/ ({count} images)\n")
                
                f.write("\n" + "=" * 80 + "\n")
                f.write("END OF REPORT\n")
                f.write("=" * 80 + "\n")
            
            self.logger.info(f"Report generated: {report_path}")
            return True
        
        except Exception as e:
            self.logger.error(f"Failed to generate report: {str(e)}")
            return False


# ==================== GUI APPLICATION ====================

class D_GITALCODE_Application(ctk.CTk):
    """Main GUI application for D-GITALCODE Word Image Extractor"""
    
    def __init__(self):
        """Initialize the application"""
        super().__init__()
        
        # Configure window
        self.title("D-GITALCODE Word Image Extractor")
        self.geometry("900x700")
        self.resizable(True, True)
        
        # Set appearance
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        # Configure style
        self.configure(fg_color=DARK_BG)
        
        # Application state
        self.selected_file = None
        self.extraction_in_progress = False
        self.logger_instance = None
        
        # Initialize logger
        self._setup_logger()
        
        # Build UI
        self._build_ui()
        
        self.logger_instance.info("Application started")
    
    def _setup_logger(self):
        """Setup logger for the application"""
        try:
            log_dir = Path.home() / ".D-GITALCODE" / "logs"
            log_dir.mkdir(parents=True, exist_ok=True)
            self.logger_instance = LoggerSetup.setup_logger(log_dir)
        except Exception as e:
            print(f"Failed to setup logger: {e}")
    
    def _build_ui(self):
        """Build the user interface"""
        # Main container
        main_frame = ctk.CTkFrame(self, fg_color=DARK_BG)
        main_frame.pack(fill="both", expand=True, padx=0, pady=0)
        
        # Header section
        self._build_header(main_frame)
        
        # Content section
        content_frame = ctk.CTkFrame(main_frame, fg_color=SECONDARY_BG)
        content_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        self._build_file_selection(content_frame)
        self._build_status_area(content_frame)
        self._build_controls(content_frame)
        self._build_progress(content_frame)
        self._build_footer(main_frame)
    
    def _build_header(self, parent):
        """Build the header section"""
        header_frame = ctk.CTkFrame(parent, fg_color=SECONDARY_BG, height=120)
        header_frame.pack(fill="x", padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        # Logo area (placeholder)
        logo_frame = ctk.CTkFrame(header_frame, fg_color=ACCENT_COLOR, height=80)
        logo_frame.pack(fill="x", padx=0, pady=10)
        logo_frame.pack_propagate(False)
        
        logo_label = ctk.CTkLabel(
            logo_frame,
            text="🎨 D-GITALCODE",
            font=("Arial", 32, "bold"),
            text_color=DARK_BG
        )
        logo_label.pack(pady=10)
        
        # Title
        title_label = ctk.CTkLabel(
            header_frame,
            text="Word Image Extractor",
            font=("Arial", 24, "bold"),
            text_color=ACCENT_COLOR
        )
        title_label.pack(pady=5)
        
        author_label = ctk.CTkLabel(
            header_frame,
            text="Developed by ELKARCH ABDELHAMID",
            font=("Arial", 11),
            text_color="#aaaaaa"
        )
        author_label.pack(pady=2)
    
    def _build_file_selection(self, parent):
        """Build file selection section"""
        section_frame = ctk.CTkFrame(parent, fg_color=DARK_BG)
        section_frame.pack(fill="x", pady=15)
        
        label = ctk.CTkLabel(
            section_frame,
            text="Step 1: Select Word Document",
            font=("Arial", 13, "bold"),
            text_color=ACCENT_COLOR
        )
        label.pack(anchor="w", pady=(0, 10))
        
        button_frame = ctk.CTkFrame(section_frame, fg_color=DARK_BG)
        button_frame.pack(fill="x")
        
        self.select_button = ctk.CTkButton(
            button_frame,
            text="📁 Browse Word Document",
            command=self._select_file,
            fg_color=ACCENT_COLOR,
            hover_color="#0088b8",
            text_color=DARK_BG,
            font=("Arial", 11, "bold"),
            height=40
        )
        self.select_button.pack(side="left", padx=5)
        
        self.file_label = ctk.CTkLabel(
            button_frame,
            text="No file selected",
            font=("Arial", 10),
            text_color="#888888"
        )
        self.file_label.pack(side="left", padx=15, fill="x", expand=True)
    
    def _build_status_area(self, parent):
        """Build status display area"""
        section_frame = ctk.CTkFrame(parent, fg_color=DARK_BG)
        section_frame.pack(fill="x", pady=15)
        
        label = ctk.CTkLabel(
            section_frame,
            text="Status Information",
            font=("Arial", 13, "bold"),
            text_color=ACCENT_COLOR
        )
        label.pack(anchor="w", pady=(0, 10))
        
        # Status text box
        self.status_text = ctk.CTkTextbox(
            section_frame,
            height=150,
            fg_color=SECONDARY_BG,
            text_color=TEXT_COLOR,
            font=("Courier", 9)
        )
        self.status_text.pack(fill="both", expand=True)
        self.status_text.insert("1.0", "Ready to extract images...\n")
        self.status_text.configure(state="disabled")
    
    def _build_controls(self, parent):
        """Build control buttons"""
        section_frame = ctk.CTkFrame(parent, fg_color=DARK_BG)
        section_frame.pack(fill="x", pady=15)
        
        label = ctk.CTkLabel(
            section_frame,
            text="Step 2: Start Extraction",
            font=("Arial", 13, "bold"),
            text_color=ACCENT_COLOR
        )
        label.pack(anchor="w", pady=(0, 10))
        
        button_frame = ctk.CTkFrame(section_frame, fg_color=DARK_BG)
        button_frame.pack(fill="x")
        
        self.extract_button = ctk.CTkButton(
            button_frame,
            text="▶️ Start Extraction",
            command=self._start_extraction,
            fg_color=SUCCESS_COLOR,
            hover_color="#00a840",
            text_color=DARK_BG,
            font=("Arial", 11, "bold"),
            height=40,
            state="disabled"
        )
        self.extract_button.pack(side="left", padx=5)
        
        self.clear_button = ctk.CTkButton(
            button_frame,
            text="🔄 Clear",
            command=self._clear_app,
            fg_color="#666666",
            hover_color="#555555",
            text_color=TEXT_COLOR,
            font=("Arial", 11, "bold"),
            height=40
        )
        self.clear_button.pack(side="left", padx=5)
        
        self.open_folder_button = ctk.CTkButton(
            button_frame,
            text="📂 Open Output Folder",
            command=self._open_output_folder,
            fg_color="#666666",
            hover_color="#555555",
            text_color=TEXT_COLOR,
            font=("Arial", 11, "bold"),
            height=40,
            state="disabled"
        )
        self.open_folder_button.pack(side="left", padx=5)
    
    def _build_progress(self, parent):
        """Build progress section"""
        section_frame = ctk.CTkFrame(parent, fg_color=DARK_BG)
        section_frame.pack(fill="x", pady=15)
        
        label = ctk.CTkLabel(
            section_frame,
            text="Progress",
            font=("Arial", 13, "bold"),
            text_color=ACCENT_COLOR
        )
        label.pack(anchor="w", pady=(0, 10))
        
        self.progress_bar = ctk.CTkProgressBar(
            section_frame,
            fg_color=SECONDARY_BG,
            progress_color=ACCENT_COLOR,
            height=20
        )
        self.progress_bar.pack(fill="x", pady=5)
        self.progress_bar.set(0)
        
        self.progress_label = ctk.CTkLabel(
            section_frame,
            text="0 images processed",
            font=("Arial", 10),
            text_color="#888888"
        )
        self.progress_label.pack(anchor="w", pady=5)
    
    def _build_footer(self, parent):
        """Build footer"""
        footer_frame = ctk.CTkFrame(parent, fg_color=SECONDARY_BG, height=40)
        footer_frame.pack(fill="x", side="bottom", padx=0, pady=0)
        footer_frame.pack_propagate(False)
        
        footer_label = ctk.CTkLabel(
            footer_frame,
            text="D-GITALCODE © 2024 | Professional Image Extraction Tool",
            font=("Arial", 9),
            text_color="#666666"
        )
        footer_label.pack(pady=10)
    
    def _select_file(self):
        """Handle file selection"""
        filetypes = [("Word Documents", "*.docx"), ("All Files", "*.*")]
        file_path = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=filetypes
        )
        
        if file_path:
            self.selected_file = file_path
            filename = Path(file_path).name
            self.file_label.configure(text=f"✓ Selected: {filename}")
            self.extract_button.configure(state="normal")
            self._update_status(f"File selected: {filename}\n", "info")
    
    def _start_extraction(self):
        """Start extraction in a separate thread"""
        if not self.selected_file:
            messagebox.showerror("Error", "Please select a Word document first")
            return
        
        if self.extraction_in_progress:
            messagebox.showwarning("Warning", "Extraction already in progress")
            return
        
        # Create output directory
        desktop = Path.home() / "Desktop"
        output_dir = desktop / OUTPUT_FOLDER_NAME
        
        # Disable controls
        self.extraction_in_progress = True
        self.extract_button.configure(state="disabled")
        self.select_button.configure(state="disabled")
        
        # Start extraction in thread
        extraction_thread = threading.Thread(
            target=self._extraction_worker,
            args=(self.selected_file, output_dir),
            daemon=True
        )
        extraction_thread.start()
    
    def _extraction_worker(self, docx_path, output_dir):
        """Worker thread for extraction"""
        try:
            self._update_status("Starting extraction process...\n", "info")
            
            # Create extractor
            extractor = ImageExtractor(
                docx_path=docx_path,
                output_dir=output_dir,
                logger_instance=self.logger_instance
            )
            
            # Extract images
            self._update_status("Extracting images from document...\n", "info")
            
            def progress_update(count):
                self._update_progress(count)
            
            success = extractor.extract_images(progress_callback=progress_update)
            
            if success:
                # Generate report
                self._update_status("Generating extraction report...\n", "info")
                extractor.generate_report()
                
                # Display results
                stats = extractor.extraction_stats
                result_text = f"\n✓ EXTRACTION SUCCESSFUL!\n"
                result_text += f"{'='*50}\n"
                result_text += f"Total Images Extracted: {stats['total_images']}\n"
                result_text += f"Extraction Time: {stats['extraction_time']:.2f} seconds\n\n"
                result_text += f"Images by Format:\n"
                
                for fmt in SUPPORTED_FORMATS:
                    count = stats["images_by_format"][fmt]
                    if count > 0:
                        result_text += f"  {fmt}: {count} image(s)\n"
                
                result_text += f"\nOutput Folder:\n"
                result_text += f"  {output_dir.absolute()}\n"
                
                self._update_status(result_text, "success")
                self.open_folder_button.configure(state="normal")
                
                # Show completion message
                self.after(100, lambda: messagebox.showinfo(
                    "Success",
                    f"Extraction completed!\n\n"
                    f"Total images extracted: {stats['total_images']}\n"
                    f"Location: {output_dir.absolute()}"
                ))
            else:
                self._update_status("✗ Extraction failed. Check status for details.\n", "error")
                self.after(100, lambda: messagebox.showerror(
                    "Error",
                    "Extraction failed. Please check the file and try again."
                ))
        
        except Exception as e:
            self.logger_instance.error(f"Extraction worker error: {str(e)}\n{traceback.format_exc()}")
            self._update_status(f"✗ Error: {str(e)}\n", "error")
            self.after(100, lambda: messagebox.showerror("Error", f"An error occurred:\n{str(e)}"))
        
        finally:
            # Re-enable controls
            self.extraction_in_progress = False
            self.extract_button.configure(state="normal")
            self.select_button.configure(state="normal")
    
    def _update_status(self, message, status_type="info"):
        """Update status text box (thread-safe)"""
        def update():
            self.status_text.configure(state="normal")
            self.status_text.insert("end", message)
            self.status_text.see("end")
            self.status_text.configure(state="disabled")
        
        self.after(0, update)
    
    def _update_progress(self, count):
        """Update progress bar (thread-safe)"""
        def update():
            self.progress_label.configure(text=f"{count} images processed")
        
        self.after(0, update)
    
    def _clear_app(self):
        """Clear application state"""
        self.selected_file = None
        self.file_label.configure(text="No file selected")
        self.extract_button.configure(state="disabled")
        self.open_folder_button.configure(state="disabled")
        self.progress_bar.set(0)
        self.progress_label.configure(text="0 images processed")
        
        self.status_text.configure(state="normal")
        self.status_text.delete("1.0", "end")
        self.status_text.insert("1.0", "Ready to extract images...\n")
        self.status_text.configure(state="disabled")
    
    def _open_output_folder(self):
        """Open output folder in file explorer"""
        try:
            desktop = Path.home() / "Desktop"
            output_dir = desktop / OUTPUT_FOLDER_NAME
            
            if output_dir.exists():
                if sys.platform == "win32":
                    os.startfile(output_dir)
                elif sys.platform == "darwin":
                    os.system(f"open '{output_dir}'")
                else:
                    os.system(f"xdg-open '{output_dir}'")
            else:
                messagebox.showwarning("Not Found", "Output folder not found")
        
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder: {str(e)}")


# ==================== MAIN ENTRY POINT ====================

def main():
    """Main entry point"""
    try:
        app = D_GITALCODE_Application()
        app.mainloop()
    except Exception as e:
        print(f"Fatal error: {e}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
