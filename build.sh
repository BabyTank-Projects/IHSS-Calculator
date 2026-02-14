#!/bin/bash
# Build script for IHSS Calendar Generator

echo "Installing dependencies..."
pip install -r requirements.txt --break-system-packages

echo "Installing PyInstaller..."
pip install pyinstaller --break-system-packages

echo "Building executable..."
pyinstaller --onefile \
    --windowed \
    --name "IHSS_Calendar_Generator" \
    --icon=NONE \
    --add-data "requirements.txt:." \
    ihsscalculator_enhanced.py

echo "Build complete! Executable is in dist/ folder"
