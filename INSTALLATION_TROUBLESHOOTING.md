# Installation & Troubleshooting Guide

## Quick Installation (2 Minutes)

### Step 1: Install Python 3.11+
- Download from: https://www.python.org/downloads/
- ✓ Check "Add Python to PATH" during installation

### Step 2: Download Project
```bash
git clone https://github.com/3ab2/extractor-images-from-docx.git
cd extractor-images-from-docx
```

### Step 3: Install Dependencies
```bash
python setup.py
```

### Step 4: Run Application
```bash
python main.py
```

---

## Manual Installation

### Method 1: Using setup.py (Recommended)
```bash
python setup.py
```

This script will:
- Check Python version
- Install all dependencies
- Create required directories
- Verify installation

### Method 2: Manual pip install
```bash
pip install -r requirements.txt
```

### Verify Installation
```bash
python test_extraction.py
```

Expected output:
```
✓ ALL TESTS PASSED!
Application is ready for use!
```

---

## Troubleshooting

### Problem: "python" command not found

**Solution (Windows):**
- Right-click Python installer → Run as Administrator
- Select "Add Python to PATH"
- Restart command prompt

**Solution (Mac/Linux):**
```bash
# Try python3 instead
python3 main.py

# Or create alias
alias python=python3
```

---

### Problem: "ModuleNotFoundError: No module named 'X'"

**Solution:**
```bash
# Reinstall all dependencies
pip install -r requirements.txt --upgrade

# Or install specific module
pip install customtkinter
pip install python-docx
pip install Pillow
pip install openpyxl
pip install reportlab
```

---

### Problem: GUI doesn't start / blank window

**Solution:**
1. Check Python version:
   ```bash
   python --version  # Should be 3.11+
   ```

2. Reinstall customtkinter:
   ```bash
   pip uninstall customtkinter
   pip install customtkinter
   ```

3. Try with verbose output:
   ```bash
   python -v main.py
   ```

---

### Problem: PDF report not generated

**Solution:**
```bash
pip install reportlab --upgrade
```

Verify installation:
```python
from reportlab import __version__
print(__version__)
```

---

### Problem: Excel report not generated

**Solution:**
```bash
pip install openpyxl --upgrade
```

Verify installation:
```python
from openpyxl import __version__
print(__version__)
```

---

### Problem: Cannot extract images from DOCX

**Solution:**
1. Verify python-docx is installed:
   ```bash
   pip install python-docx --upgrade
   ```

2. Test with a different DOCX file

3. Check if images are actually embedded in the document

4. Look at extraction logs for detailed error messages

---

### Problem: Extraction is very slow

**Causes & Solutions:**

**Large document (100+ images):**
- This is normal, extraction takes time
- Monitor progress in "⏳ Progress" tab
- GUI remains responsive

**Disk space issue:**
- Check available disk space
- Extracted images need space

**Corrupted document:**
- Try extracting different DOCX file
- Check logs for error details

**System resources:**
- Close unnecessary applications
- Check available RAM

---

### Problem: GUI freezes / becomes unresponsive

**Solution:**
This should NOT happen. If it does:
1. Close the application (Ctrl+C)
2. Restart: `python main.py`
3. Use smaller documents to test
4. Report the issue with document details

---

### Problem: Duplicate detection not working

**Solution:**
1. Verify SHA256 algorithm in config:
   ```python
   # utils.py line ~XX
   HASH_ALGORITHM = "sha256"  # Correct
   ```

2. Check that images are truly identical

3. Reset extractor state if running multiple times

---

### Problem: Output folder not created

**Solution:**
1. Check Desktop exists and is writable:
   ```bash
   # Windows
   cd %USERPROFILE%\Desktop
   
   # Mac/Linux
   cd ~/Desktop
   ```

2. Create manually if needed:
   ```bash
   mkdir "Extracted_Images_D-GITALCODE"
   ```

3. Run extraction again

---

### Problem: Language selection doesn't work

**Solution:**
1. Restart the application
2. Select language from dropdown
3. Start new extraction

---

### Problem: Build to EXE fails

**Solution:**
1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```

2. Run build script:
   ```bash
   python build.py
   ```

3. If still fails, try verbose output:
   ```bash
   python build.py -v
   ```

---

## Performance Optimization

### For Large Batch Processing

**1. Increase chunk size** (config.py):
```python
CHUNK_SIZE = 2048 * 1024  # 2MB instead of 1MB
```

**2. Reduce progress updates** (config.py):
```python
PROGRESS_UPDATE_INTERVAL = 500  # Milliseconds
```

**3. Close unnecessary applications**

**4. Use SSD for output folder** (faster writes)

---

## System Requirements

### Minimum
- Python 3.11+
- Windows 7, macOS 10.13+, or Linux
- 4 GB RAM
- 500 MB free disk space

### Recommended
- Python 3.11+
- Windows 10+, macOS 12+, or Linux
- 8 GB RAM
- 2 GB free disk space
- SSD for better performance

---

## File Permissions

### Ensure write permissions for:
- `~/Desktop/Extracted_Images_D-GITALCODE/` (output folder)
- `~/.D-GITALCODE/logs/` (log files)
- `~/.D-GITALCODE/cache/` (cache folder)

### Fix permissions (Linux/Mac):
```bash
chmod -R 755 ~/.D-GITALCODE
chmod -R 755 ~/Desktop/Extracted_Images_D-GITALCODE
```

---

## Getting Help

### 1. Check Logs
Logs are saved to: `~/.D-GITALCODE/logs/extraction.log`

### 2. Run Test Suite
```bash
python test_extraction.py
```

### 3. Check Documentation
- README.md - Full documentation
- QUICKSTART.md - Quick reference
- PROJECT_STRUCTURE.md - Architecture details
- ADVANCED_USAGE.py - Code examples

### 4. Common Issues Checklist
- [ ] Python 3.11+ installed?
- [ ] All dependencies installed? (run `python test_extraction.py`)
- [ ] DOCX file is valid and contains images?
- [ ] Output folder has write permissions?
- [ ] Sufficient disk space available?
- [ ] No corrupted DOCX file?

---

## Support

**Author:** ELKARCH ABDELHAMID  
**Email:** 3ab2uelkarch2006@gmail.com  
**Brand:** D-GITALCODE  

---

## Version Info

Current Version: **2.0.0**  
Python Required: **3.11+**  
Last Updated: **2024**  

---

**D-GITALCODE © 2024 | Professional Image Extraction Tool**
