# Certificate Generator 🎓

A modern, fast, and attractive bulk certificate generator. This tool automatically reads names from an Excel sheet (or CSV file) and inserts them centered into a PDF certificate template where a placeholder keyword exists.

It features two modern user interfaces:
1. **CustomTkinter Desktop GUI** - A native desktop application window with a dark theme.
2. **Streamlit Web GUI** - A browser-based interface supporting drag-and-drop file uploads, dynamic excel previews, and ZIP downloads.

---

## 🛠️ Installation

1. Ensure you have Python 3.8+ installed.
2. Install the required libraries using pip:
   ```bash
   pip install -r requirements.txt
   ```

---

## 🚀 How to Run

### Option 1: CustomTkinter Desktop App (Recommended)
This runs as a standalone desktop window on your system.
```bash
python gui.py
```
- **Features**: Clean dark layout, native file browsing dialogs, sliders for font size, dropdown menus for font selection, and non-blocking multi-threaded processing.

### Option 2: Streamlit Web Dashboard
This runs as a modern, interactive web dashboard in your browser.
```bash
streamlit run app.py
```
- **Features**: Drag-and-drop file uploaders, live spreadsheet preview, configuration sidebar, progress feedback, and single-click ZIP download of all generated certificates.

### Option 3: Command Line (CLI)
You can still run it via the command line. Put your `template.pdf` and `names.xlsx` in the root folder, check that the column containing names is headed `"Name"`, and run:
```bash
python script.py
```

---

## 📝 Document Formatting Tips

- **Name Column**: Make sure your Excel/CSV names file contains a column header labeled `"Name"` (case-insensitive). If no `"Name"` column is found, the script will fallback to the first column.
- **Certificate ID**:
  - The script will automatically scan for a `"Certificate ID"` column. If it exists, it will use the existing IDs.
  - If it is missing, the script will generate unique IDs (e.g. `CERT-A1B2C3D4`) and write them back into the source Excel/CSV file under a new `"Certificate ID"` column.
- **Placeholders**:
  - **Name**: Add the placeholder text `NAME_PLACEHOLDER` inside your PDF template where you want the name to be printed.
  - **QR Code**: Add the placeholder text `QR_PLACEHOLDER` in the PDF template where you want the QR code image to be inserted.
  - **ID Text**: Add the placeholder text `ID_PLACEHOLDER` in the PDF template where you want the Certificate ID text to be printed.
  - *Fallback*: If `QR_PLACEHOLDER` and `ID_PLACEHOLDER` are not found in the template, they will automatically fallback to a standard position in the bottom-left corner of the certificate.
- **Text Color**: Keep placeholder text colors white (or matching the background color) so that the raw placeholder text isn't visible behind the printed text/images.

