# combinefilewithsamename

**Combine multiple Excel workbooks with a popup file selector!**

A simple Python tool for MIS Executives to merge `.xlsx`, `.xls`, `.xlsm`, `.xlsb` files into one workbook — with a beautiful file picker popup.

---

## 🚀 Installation

```bash
pip install combinefilewithsamename
```

---

## 📦 Usage

### Option 1: Command Line
```bash
combinefiles
```

### Option 2: Python Module
```bash
python -m combinefilewithsamename
```

### Option 3: In Python Script
```python
from combinefilewithsamename import run
run()

# Or directly combine files
from combinefilewithsamename import combine_excel_files
combine_excel_files()

# Or get VBA code for Excel
from combinefilewithsamename import inject_vba
inject_vba()
```

---

## ✨ Features

- 📂 **Popup file selector** — choose multiple files easily
- 📊 **All Excel formats** — `.xlsx`, `.xls`, `.xlsm`, `.xlsb`
- 🔄 **Auto rename sheets** — handles duplicate sheet names
- 💾 **Save dialog** — choose where to save combined file
- 📋 **VBA code** — get macro code to use directly in Excel
- ✅ **Summary message** — shows how many files/sheets combined

---

## 🖥️ How It Works

1. Run the tool
2. **Popup appears** → select your Excel files
3. Choose save location
4. Done! All sheets combined into one file

---

## 📋 VBA Option (For Excel Users)

If you prefer running inside Excel:

```python
from combinefilewithsamename import inject_vba
inject_vba()
```

This copies VBA macro code. Then:
1. Open Excel → `Alt + F11`
2. `Insert > Module`
3. Paste code → `F5` to run

---

## 🛠️ Requirements

- Python 3.7+
- openpyxl (auto-installed)
- tkinter (built into Python)

---

## 👨‍💼 Made For

MIS Executives who work with multiple Excel reports daily and need a quick way to combine them!

---

## 📄 License

MIT License
