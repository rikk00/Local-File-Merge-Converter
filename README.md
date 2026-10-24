# Windows File Converter

This tool adds **Convert to → Markdown** and **Convert to → PDF** to the File Explorer context menu for files.

## Setup

1. Install Python for Windows. The installer finds `pythonw.exe` on `PATH` or through the `py` launcher.
2. Right-click `install.bat` and select **Run as administrator**, then approve the UAC prompt.
3. The installer installs `requirements.txt`, registers the **Convert to** context menu, and adds **Merge as PDF** to File Explorer's **Send to** menu.

`install.bat` must remain beside `converter.py`, `merge_pdf.py`, and `requirements.txt`.

## Supported PDF inputs

Markdown conversion uses MarkItDown and its optional format integrations. PDF output supports common image formats through Pillow and Word, Excel, and PowerPoint formats through Microsoft Office automation. Microsoft Office must be installed for Office document conversion.

The output is placed next to the source with the same base name. Existing output files are not overwritten. Selecting Markdown for an `.md` file or PDF for a `.pdf` file leaves the source unchanged.

## Merge PDFs, images, and Office documents

`merge_pdf.py` combines selected PDF, JPG/PNG, Word (`.doc`, `.docx`), and Excel (`.xls`, `.xlsx`) files. Office conversion requires Microsoft Office.

After installation, select multiple files in Explorer and choose **Send to → Merge as PDF**.
