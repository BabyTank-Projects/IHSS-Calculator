# IHSS Hours Calendar Generator

A modern, dark-themed calendar application for IHSS (In-Home Supportive Services) caregivers to manage and track their work hours efficiently.

![IHSS Calendar Generator](screenshot.png)

## Features

### 📅 Calendar Management
- Interactive monthly calendar with customizable pay periods
- Support for full month or split pay periods (1-15, 16-end)
- Color-coded weekly totals (green indicators)
- Real-time hour tracking and validation

### ⏰ Work Time Calculator
- Calculate end times based on start time and daily hours
- Display finish times directly in calendar
- Supports both 12-hour (AM/PM) and 24-hour time formats

### ✨ Auto-Fill
- Automatically distribute authorized hours across selected workdays
- Option to use whole hours for cleaner scheduling
- Customizable workday selection (Sun-Sat)

### 🎯 Exemptions Support
- Live-In Family Care Provider exemption
- Extraordinary Circumstances exemption
- Custom weekly hour override option

### 💾 Export & Screenshot
- Export timesheets to beautifully formatted **Excel files (.xlsx)** with:
  - 🎨 Professional colors and styling
  - 📊 Clean, easy-to-read layout
  - ⏰ **Work Time Calculator data included** (start/end times when used)
  - 📈 Alternating row colors for better readability
  - 🖨️ Print-ready formatting
- Take screenshots of calendar for visual reference
- Timestamped filenames for easy organization
- **Files automatically saved to your Documents folder**

### 🎨 Modern UI
- Dark theme (easy on the eyes)
- Clean, professional interface
- Large, readable text and input fields
- Responsive layout

### ❓ Built-in Help
- Comprehensive help documentation
- Feature explanations and tips
- Usage guidelines

## Download & Installation

### Option 1: Download Executable (No Python Required)
**For Windows Users:**
1. Go to the [Releases](https://github.com/yourusername/ihss-calendar/releases) page
2. Download `IHSS_Calendar_Generator.exe`
3. Double-click to run - no installation needed!

### Option 2: Run from Source (Python Required)
**Requirements:**
- Python 3.8 or higher
- pip (Python package manager)

**Installation:**
```bash
# Clone the repository
git clone https://github.com/yourusername/ihss-calendar.git
cd ihss-calendar

# Install dependencies (includes Pillow for screenshots and openpyxl for Excel export)
pip install -r requirements.txt

# Run the application
python ihsscalculator_enhanced.py
```

## Building the Executable Yourself

If you want to build the executable from source:

**Windows:**
```batch
build.bat
```

**Linux/Mac:**
```bash
chmod +x build.sh
./build.sh
```

The executable will be created in the `dist/` folder.

## Usage Guide

### Getting Started
1. **Set Monthly Hours**: Enter your authorized monthly hours and minutes
2. **Select Exemptions**: Check any applicable exemptions (if applicable)
3. **Choose Workdays**: Select which days you work (leave unchecked for all days)
4. **Select Period**: Choose Full Month or Pay Period 1/2
5. **Auto-Fill or Manual Entry**: 
   - Click "Auto-Fill" to automatically distribute hours
   - Or manually enter hours in each day's cell

### Work Time Calculator
1. Enter your typical start time (e.g., "9:00 AM")
2. Click "Calculate" to see when you'll finish each day
3. End times appear below daily hours in the calendar

### Exporting Data
- **Excel Export**: Click "Export Excel" to save a professionally formatted timesheet
  - Beautiful color-coded headers (blue) and totals (green)
  - Alternating row colors for easy reading
  - Start and end times are automatically included in the export!
  - Files are saved to your **Documents** folder
  - Filename includes date and pay period for easy organization
  - Opens perfectly in Microsoft Excel, Google Sheets, LibreOffice, or Numbers
- **Screenshot**: Click "Screenshot" to capture calendar image
  - Images saved to your **Documents** folder
  - Useful for visual reference when filling out official forms

### Input Formats
Hours can be entered as:
- Whole numbers: `8` (8 hours)
- Decimals: `8.5` (8.5 hours)
- Time format: `8:30` (8 hours 30 minutes)

## Tips

✅ Always verify your schedule matches actual hours worked  
✅ Use Export Excel to get beautifully formatted timesheets with professional colors  
✅ If you use the Work Time Calculator, your start/end times will be included in the Excel export  
✅ Use Screenshot for visual reference when filling out official forms  
✅ Weekly maximum calculations are helpers - follow official IHSS guidelines  
✅ You can mix manual entry with auto-fill by clearing specific days  
✅ Exported files are saved to your Documents folder for easy access
✅ Excel files open in Excel, Google Sheets, LibreOffice, or Numbers

## Technical Details

**Built with:**
- Python 3.x
- Tkinter (GUI framework)
- Pillow (screenshot functionality)
- openpyxl (Excel file generation with formatting)

**Compatible with:**
- Windows 10/11
- macOS 10.14+
- Linux (Ubuntu 20.04+)

## Disclaimer

This is a personal scheduling and calculation tool. Always bill ONLY hours actually worked and follow IHSS rules and laws. The maximum weekly calculation is a simple helper and may not match every IHSS situation. Use official IHSS guidance when in doubt.

## License

MIT License - see LICENSE file for details

## Support

For issues, questions, or feature requests, please open an issue on GitHub.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Changelog

### Version 3.0.0 (Latest) - Excel Export Fixed!
- **CRITICAL FIX**: Excel export now works in the .exe (no more "Missing Library" error!)
- **FIXED**: openpyxl library now properly bundled in executable
- **FIXED**: Updated build scripts with --hidden-import flags
- All features from v1.1.0 now fully functional in the executable
- Users can now export beautiful Excel files with no installation required!

### Version 2.1.0 - Beautiful Excel Export!
- **NEW**: Export to Excel (.xlsx) instead of CSV with beautiful formatting
- **NEW**: Professional color scheme - blue headers, green totals, alternating gray rows
- **NEW**: Work Time Calculator data automatically included in exports (start/end times)
- Clean, print-ready layout perfect for record-keeping
- Better readability with styled columns and borders
- Added openpyxl dependency for Excel generation

### Version 2.0.2 - File Export Fix
- **CRITICAL FIX**: Export CSV and Screenshot now save to Documents folder
- Fixed permission denied errors when running as executable
- Files no longer try to save to protected system directories
- Added robust fallback logic for finding writable directories
- Better error messages showing exact save location

### Version 2.0.0
- **CRITICAL FIX**: Calendar now correctly displays days under proper weekday headers
- Fixed workday selection labels (Su, M, Tu, W, Th, F, Sa)
- Improved calendar cell sizing and spacing
- Added clearer workday abbreviations

### Version 1.0.0
- Initial release
- Calendar management with pay period support
- Work time calculator
- Auto-fill functionality
- CSV export and screenshot features
- Dark theme UI
- Comprehensive help system

## Troubleshooting

**Q: Getting "Missing Library" error when exporting?**  
A: Download the latest version (**v3.0.0**). The library is now included in the .exe - no installation needed!

**Q: Getting "Permission denied" errors when exporting?**  
A: Make sure you're using version 1.0.2 or later. Files now save to your Documents folder automatically.

**Q: Where do my exported files go?**  
A: Excel exports and screenshots are saved to:
- **Windows**: `C:\Users\[YourName]\Documents`
- **Mac**: `/Users/[YourName]/Documents`
- **Linux**: `/home/[YourName]/Documents`

**Q: Can I open the Excel files in Google Sheets?**  
A: Yes! The .xlsx files work perfectly in Microsoft Excel, Google Sheets, LibreOffice Calc, and Apple Numbers.

**Q: "Windows protected your PC" warning**  
A: Click "More info" → "Run anyway". This is normal for unsigned executables.

**Q: I'm running from Python source - "Missing Library" error?**  
A: Install the required libraries: `pip install openpyxl Pillow`

**Q: CSV export instead of Excel?**  
A: If you choose CSV fallback, it still includes all your data and Work Time Calculator info, just without the colors.

---

**Made with ❤️ for IHSS caregivers**
