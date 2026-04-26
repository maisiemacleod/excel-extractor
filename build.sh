#!/usr/bin/env bash
# Build script for Excel Field Extractor (run on Windows with Python + PyInstaller installed)
# Usage: bash build.sh  (or on Windows: python -m PyInstaller ...)

# Install dependencies first
pip install -r requirements.txt

# Build single-file exe
# --onefile        : pack everything into a single exe
# --noconsole      : hide console window (GUI app)
# --name           : output exe name
pyinstaller \
  --onefile \
  --noconsole \
  --name "ExcelFieldExtractor" \
  excel_extractor.py

echo "Build complete. Executable is at dist/ExcelFieldExtractor.exe"
