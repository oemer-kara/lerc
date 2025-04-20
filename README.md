# LERC (Last Epoch Rarity Checker)

A Python utility that extracts item names from screenshots using OCR (Optical Character Recognition) and looks up their reroll chances from a database.

## Features

- **Clipboard OCR Processing**: Automatically detects and processes images from your clipboard
- **Hotkey Support**: Use `Ctrl+Shift+D` to quickly process your clipboard contents
- **Database Lookup**: Checks item names against a CSV database for reroll chances
- **Automatic Dependency Installation**: Installs required Python packages if missing
- **Comprehensive Logging**: Detailed logs in `ocr_tool.log` for troubleshooting

## Requirements

- Windows OS (uses Windows-specific clipboard handling)
- Python 3.6+
- Tesseract OCR (will prompt to install if not found)

## Installation

1. **Install Python dependencies**:
   ```bash
   pip install pillow pytesseract keyboard pywin32
   ```

2. **Install Tesseract OCR**:
   - Download from [UB Mannheim Tesseract](https://github.com/UB-Mannheim/tesseract/wiki)
   - The tool will detect if Tesseract is missing and guide you through installation

3. **Prepare your items database**:
   - Create a CSV file named `items.csv` in the same directory
   - Format: `ItemName,OtherFields...,RerollChance`

## Usage
![ezgif-370c1fdd611986](https://github.com/user-attachments/assets/cb5d1b43-3295-442d-b874-bf7d33e6f167)

1. Take a screenshot of an **item's name** (use Windows `Win+Shift+S` for region capture)
2. Press `Ctrl+Shift+D` to process the image
3. View the results in a popup dialog showing:
   - Extracted item name
   - Reroll chance from database

## Troubleshooting

- Check `ocr_tool.log` for detailed error information
- If OCR fails, verify Tesseract is installed correctly
- For database issues, ensure `items.csv` exists and is properly formatted

## Example Database Format

```
ItemName,Type,Value,Weight,Rarity,RerollChance
GOLD COINS,Currency,1,1,Common,0% (Common)
RUBY,Ring,250,1,Uncommon,10% (Uncommon)
```
Current database can be found [here](https://docs.google.com/spreadsheets/d/e/2PACX-1vRTCibOIdTYtqq6vH10Oerk50_XpixR_rjtf7Ovu59MBib4vgsj24k9nRIjN4Lg21HUJstYk2UfqLWm/pubhtml)

## Creating an Executable (LERC.exe)

### Building the Executable with PyInstaller

1. **Install PyInstaller**:
   ```bash
   pip install pyinstaller
   ```

2. **Build the executable**:
   ```bash
   pyinstaller --onefile --add-data "items.csv;." --icon=lerc.ico --name lerc lerc.py
   ```

3. **Find your executable**:
   - Look in the `dist` folder for `lerc.exe`
   - You can distribute this single file to users
   
### Adding an Icon (Optional)

- Create or find an .ico file for your application
- Include it in the PyInstaller command as shown above with `--icon=lerc.ico`

## Inspiration

https://github.com/0xdreadnaught/le-search

## License

This project is provided as-is without warranty. Feel free to modify for your needs.
