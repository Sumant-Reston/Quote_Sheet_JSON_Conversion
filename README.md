# Quote Sheet JSON Conversion

A simple Python application with a graphical interface to convert specific pallet data from an Excel sheet into a structured JSON file.

---

## What Does This App Do?

This tool extracts pallet information from a pre-formatted Excel file (with data in specific cells) and converts it into a JSON file for easy digital use or integration with other systems.

- **Input:** An Excel file (.xlsx or .xls) with pallet data in fixed cell locations (see below).
- **Output:** A JSON file containing the extracted pallet and build details.

---

## Requirements

- Python 3.x
- [openpyxl](https://pypi.org/project/openpyxl/) (for reading Excel files)

Install dependencies with:

```
pip install openpyxl
```

---

## How to Use

### 1. Run as a Python Script

1. Open a terminal in this project folder.
2. Run:
   ```
   python QSC.py
   ```
3. The app window will open.

### 2. Run as a Windows Executable

1. Build the executable (requires [pyinstaller](https://pypi.org/project/pyinstaller/)):
   ```
   pip install pyinstaller
   pyinstaller --onefile --windowed QSC.py
   ```
2. Find `QSC.exe` in the `dist` folder and double-click to run.

---

## App Workflow

1. **Select Excel File**
   - Click "Select Excel File" and choose your .xlsx file.
2. **Convert to JSON**
   - Click "Convert to JSON" and choose where to save the output .json file.
3. **Done!**
   - Your JSON file is ready for use.

---

## Excel Input Format

The app expects data in specific cells (row/column) of the Excel sheet. Example mapping:

- PalletID: A18
- Length: B18
- Width: C18
- Delivery: B20
- Notched: B21
- HT: B22
- Type: B23
- Build details (TOP, MIDDLE, BOTTOM): rows 27, 29, 31 (columns B–L)

If your Excel file does not match this format, the conversion may not work correctly.

---

## Output Example

See `example_output.json` for a sample of the generated JSON structure.

---

## Support

For questions or issues, please open an issue on this repository.
