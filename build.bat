@echo off
REM Build script for IHSS Calendar Generator (Windows)

echo Installing dependencies...
pip install -r requirements.txt

echo Installing PyInstaller...
pip install pyinstaller

echo Building executable...
pyinstaller --onefile ^
    --windowed ^
    --name "IHSS_Calendar_Generator" ^
    --icon=NONE ^
    --add-data "requirements.txt;." ^
    ihsscalculator_enhanced.py

echo Build complete! Executable is in dist\ folder
pause
