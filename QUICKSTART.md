# Quick Start Guide

## For End Users (No Programming Experience)

### Windows Users:
1. Download `IHSS_Calendar_Generator.exe` from the Releases page
2. Double-click the .exe file to run
3. That's it! No installation needed.

**Note:** Windows may show a security warning because the app isn't from the Microsoft Store. Click "More info" → "Run anyway"

### Mac Users:
1. Download `IHSS_Calendar_Generator` from the Releases page
2. Right-click the file → "Open"
3. If prompted about security, go to System Preferences → Security & Privacy → "Open Anyway"

### First Time Setup:
1. Enter your monthly authorized hours (e.g., 154 hours, 0 minutes)
2. Select your typical workdays
3. Click "Auto-Fill" to see it in action!

---

## For Developers

### Building from Source:

**Windows:**
```batch
git clone https://github.com/yourusername/ihss-calendar.git
cd ihss-calendar
build.bat
```

**Mac/Linux:**
```bash
git clone https://github.com/yourusername/ihss-calendar.git
cd ihss-calendar
chmod +x build.sh
./build.sh
```

The executable will be in the `dist/` folder.

### Development Setup:
```bash
pip install -r requirements.txt
python ihsscalculator_enhanced.py
```

---

## Common Issues

**Q: "Windows protected your PC" warning**  
A: Click "More info" → "Run anyway". This is normal for unsigned executables.

**Q: Screenshot button doesn't work**  
A: Make sure Pillow is installed: `pip install Pillow`

**Q: Application won't start**  
A: Try downloading the latest release or building from source.

---

## Need Help?

- Check the full README.md for detailed documentation
- Click the "Help" button in the app for feature explanations
- Open an issue on GitHub for bugs or questions
