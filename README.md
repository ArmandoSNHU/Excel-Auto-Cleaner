# 🧼 Excel Auto Cleaner

A simple drag-and-drop desktop tool that cleans up messy Excel or CSV files. Built with Python, it removes duplicates, trims spaces, standardizes headers, and gives you the option to download the cleaned file as a new Excel or PDF.

---

## ✅ Features

- Drag-and-drop interface (file picker included)
- Cleans:
  - Empty rows
  - Duplicate rows
  - Extra spaces in text
  - Inconsistent column headers
- Export options:
  - Cleaned Excel file (`.xlsx`)
  - PDF summary (`.pdf`)

---

## 📦 Requirements

Install the required libraries with pip:

```bash
pip install pandas openpyxl reportlab
