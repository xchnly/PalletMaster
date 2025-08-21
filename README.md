# 📦 PalletMaster

**PalletMaster** is a professional Python application for generating and managing pallet labels with an intuitive GUI interface.  
It supports importing data from various formats, setting production capacity, and exporting high-quality labels to PDF.

---

## ✨ Features

- 🗃️ Import pallet data from Excel/CSV files
- 🏭 Set daily production capacity limits
- 🖨️ Export pallet labels to PDF with clean, professional layout
- 💾 Save and load application settings via JSON configuration
- 🌍 Full Unicode font support (including Chinese characters)
- 🎨 User-friendly graphical interface
- 📊 Production progress tracking
- 🔢 Automatic serial number generation

---

## 📁 Project Structure

```
PalletMaster/
│
├── 📂 build/                 # Build output (ignored in git)
├── 📂 dist/                  # Executables (ignored in git)
├── 📂 src/                   # Main source code
│   ├── 🐍 Pallet.py          # Main application logic
│   ├── 🎨 UI.py              # User interface components
│   └── 📄 ...                # Additional modules
│
├── 📄 requirements.txt       # Python dependencies
├── 📄 README.md             # Project documentation
├── 🖼️ Pallet.ico            # Application icon
└── 🔤 SourceHanSansSC-Bold.ttf  # Unicode font file
```

> ⚠️ Note: The `build/`, `dist/`, and executable files (`.exe`, `.pkg`) are **not included** in this repository due to GitHub's 100MB file size limit.

---

## 🛠️ Installation & Setup

### 1️⃣ Clone the Repository
```bash
git clone https://github.com/xchnly/PalletMaster.git
cd PalletMaster
```

### 2️⃣ Create Virtual Environment (Recommended)
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python -m venv venv
source venv/bin/activate
```

### 3️⃣ Install Dependencies
```bash
pip install -r requirements.txt
```

### 4️⃣ Run the Application
```bash
python src/Pallet.py
```

---

## 📦 Building Executables

### 🐍 Using PyInstaller
Create a standalone executable with this command:

```bash
pyinstaller --onefile --noconsole --icon=Pallet.ico --add-data "SourceHanSansSC-Bold.ttf;." src/Pallet.py
```

- `--onefile` → Bundle into a single executable
- `--noconsole` → Hide console window (for GUI applications)
- `--icon=Pallet.ico` → Add custom application icon
- `--add-data` → Include additional resources (fonts, etc.)

The generated executable will be available in the `dist/` folder.

---

## 📋 Usage Guide

1. **Import Data** 📥 - Load pallet information from Excel or CSV files
2. **Set Capacity** 🔢 - Define daily production limits
3. **Generate Labels** 🏷️ - Create professional pallet labels
4. **Export PDF** 📄 - Save labels as printable PDF documents
5. **Save Settings** 💾 - Preserve your configuration for future sessions

---

## 🆘 Support

For issues or questions:
1. Check the troubleshooting section below
2. Review the project's GitHub Issues page
3. Contact the development team

---

## 🔄 Version History

- **v1.0.0** (Current) - Initial release with core functionality
- Planned features: Barcode support, cloud integration, batch processing

---

---

## 👥 Contributing

We welcome contributions! Please feel free to submit pull requests or open issues for bugs and feature requests.

---

*Made with ❤️ by the STARLIGH TECHNOLOGY development team*
