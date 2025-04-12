# Mondo Clean Excel

**Mondo Clean Excel** is a lightweight, portable Python application that automatically cleans and reformats messy Excel and CSV files — perfect for healthcare records, administrative data, and inventory sheets.

⚡ **Key Features**
- Clean and normalize headers, remove duplicates, and strip extra whitespace
- Export results as `.xlsx` or `.pdf`
- Auto-adjust column widths and row heights (mimics Excel AutoFit)
- Embedded logo background for professional branding
- Batch processing support — clean multiple files in one click
- "Ta-Da" sound on completion 🎉
- Fully portable — no Python install required (one-click `.exe`)

---

### 📦 Built With
- `tkinter` – lightweight Python GUI
- `pandas` – powerful data processing
- `openpyxl` – Excel manipulation
- `reportlab` – PDF export
- `pyinstaller` – bundled into a portable `.exe`
- `Pillow` – handles embedded images for branding

---

### 🚀 How to Use

#### 🖥 Windows `.exe` version
1. [Download the EXE](#) *(link to your GitHub Releases)*
2. Double-click `MondoClean.exe`
3. Select one or more Excel/CSV files
4. Choose to export as Excel or PDF
5. Done — cleaned files are saved next to your originals

No install. No admin rights. No dependencies.

#### 🐍 Python version
If you'd rather run from source:

```bash
git clone https://github.com/ArmandoSNHU/Excel-Auto-Cleaner.git
cd Excel-Auto-Cleaner
pip install -r requirements.txt
python excel_cleaner_gui.py
