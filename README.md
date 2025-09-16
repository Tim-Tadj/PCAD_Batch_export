# PCAD Batch Export

PCAD Batch Export automates the repetitive report workflow inside PowerCad-5. It drives the desktop application to export project reports in batches, converts the generated PDFs to Word documents, and extracts key cable data into a consolidated CSV so that engineers can review and share results quickly.

## Features
- Automates PowerCad-5 to open each `.QPJ` project and print the configured report bundle to PDF.
- Optionally converts every PDF into a Word document based on the provided template and appends them into `concatenated.docx`.
- Parses each PDF to capture load, breaker, capacity, and impedance values; produces `0cable_info.csv` with pass/fail checks.
- Simple Qt interface with material theme for picking folders, enabling the steps you need, and tracking progress.

## Prerequisites
- Windows 10/11 with PowerCad-5 installed and licensed. The application must be running and the main window titled `PowerCad-5 - Version*` must be accessible.
- Python 3.12 (see `.python-version`).
- Poppler utilities (`pdftoppm.exe` and `pdfinfo.exe`) are bundled in the repository; ensure they remain alongside `batch_export.py` or on your `PATH`.

## Installation
1. Install [uv](https://github.com/astral-sh/uv) if it is not already available. One quick option on Windows is:
   ```powershell
   pip install uv
   ```
2. Create and activate a virtual environment managed by uv:
   ```powershell
   uv venv
   .\.venv\Scripts\Activate.ps1
   ```
3. Install the Python dependencies with uv:
   ```powershell
   uv pip install PySide6 pywinauto PyPDF2 pdf2image pillow python-docx docxcompose qt-material keyboard comtypes pywin32
   ```
   > `pdf2image` relies on the Poppler executables included in this repository; no additional install is required if you run the app from here.
4. (Optional) Build a frozen executable with cx_Freeze:
   ```powershell
   uv run python setup.py build
   ```
## Usage
1. Launch the UI:
   ```powershell
   uv run python batch_export.py
   ```
2. In PowerCad-5, ensure no modal dialogs are open and that the project window is visible.
3. In the PCAD Batch Export window:
   - **PCAD Folder**: select the directory that contains the `.QPJ` files you wish to process.
   - **PDF Folder**: choose the destination directory for generated PDFs and subsequent outputs.
   - Enable one or more tasks:
     - `PCAD Batch Export` - controls PowerCad-5 to export each project to PDF (existing PDFs in the output folder can be deleted for a clean run).
     - `Convert PDFs to Word` - crops each PDF page, drops it into a Word document based on `template.docx`, and produces a combined `concatenated.docx`.
     - `Convert Info to CSV` - parses cable metrics from each PDF and writes `0cable_info.csv`.
4. Click **Execute**. The progress bar reflects each stage; tasks run sequentially even when multiple are selected.
5. When the status reads "Complete", review the generated files in your output folder.

## Outputs
- Individual report PDFs named after the source `.QPJ` files.
- Word exports per cable plus a merged `concatenated.docx` (when conversion is enabled).
- `0cable_info.csv` summarising demand, breaker rating, cable capacity, impedance limits, and pass/fail status.

## Troubleshooting
- If the app cannot interact with PowerCad-5, confirm the program is launched under the same user session and that the window title matches `PowerCad-5 - Version*`.
- Keep the mouse and keyboard idle while the export is running; user interaction with PowerCad may interrupt automation.
- `pdf2image` errors typically mean the Poppler executables are missing; check that `pdfinfo.exe` and `pdftoppm.exe` remain in the project root.

## Development
- UI theming is controlled by `color-theme.xml`; adjust or replace the file to tweak the qt-material theme.
- Word generation uses `template.docx` for per-cable pages and `default.docx` as the base for the concatenated document. Modify these templates to change branding or layout.
- Packaging is handled by `setup.py` with cx_Freeze. Update the `build_exe_options` there if additional resources are needed.

