@echo off
REM Build script for IHSS Calendar Generator (Windows)
REM Uses python -m pip instead of pip directly for better compatibility

echo Installing dependencies...
python -m pip install -r requirements.txt

echo Installing PyInstaller...
python -m pip install pyinstaller

echo Building executable...
python -m PyInstaller --onefile ^
    --windowed ^
    --name "IHSS_Calendar_Generator" ^
    --icon=NONE ^
    --hidden-import=openpyxl ^
    --hidden-import=openpyxl.styles ^
    --hidden-import=openpyxl.utils ^
    --hidden-import=PIL ^
    --hidden-import=PIL.ImageGrab ^
    --add-data "requirements.txt;." ^
    ihsscalculator_enhanced.py

echo Build complete! Executable is in dist\ folder
pause
