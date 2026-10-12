# PDF Merge in Explorer

![Platform](https://img.shields.io/badge/platform-Windows-blue)
![Python](https://img.shields.io/badge/python-3.10%2B-green)
![File Manager](https://img.shields.io/badge/file--manager-Explorer-orange)
![License](https://img.shields.io/badge/license-GPL%20v3-blue)

**English** | [Deutsch](README.de.md)

Merge files into one PDF or convert them to PDF one by one, straight from Windows Explorer – via the **"Send to"** menu.

Windows counterpart of [Linux-PDF-Merge-in-Nemo](https://github.com/robitschmatthias-ui/Linux-PDF-Merge-in-Nemo), with the same three modes.

---

## Features

- **Three entries under "Send to":**
  - **PDF Quick Merge** – merge right away, no questions. Saved as `Merge.pdf` (or `Merge2.pdf`, `Merge3.pdf` …) in the folder of the source files
  - **PDF Advanced Merge** – set the order by drag & drop (or ▲/▼), choose page size and orientation. Suggested file name = name of the first file in the list
  - **Convert to PDF** – each selected file becomes its own PDF in the same folder (images fitted to A4). No merging
- **Supported formats:** PDF, images (JPG, PNG, TIFF, BMP, WebP, GIF), Office documents (DOC, DOCX, ODT, XLSX, ODS, PPTX, ODP), text (TXT, RTF)
- **Office files via Microsoft Office** – Word, Excel and PowerPoint create the PDF themselves, so the layout matches the original exactly. Without Office, LibreOffice is used if installed
- **Fit to A4** (Advanced) – PDF and Office pages are scaled to A4 and centered. Images that reach at least 200 DPI on A4 fill the page; smaller images keep their original size, centered, instead of being blown up and pixelated
- **Orientation** (Advanced) – keep original, force portrait or landscape (pages in the wrong orientation are rotated by 90° – also combined with A4)
- **Open PDF after merging** (Advanced) – optionally opens the result in the default PDF program
- **Nothing gets overwritten** – "Convert to PDF" saves as `Name (2).pdf` if the PDF already exists; Advanced Merge asks before overwriting
- **Explorer order** – files are sorted by name the way Explorer does (`File2` before `File10`)

## Requirements

| Component | Purpose | Installation |
|---|---|---|
| Python 3.10+ | runs the tool | automatically via winget if missing |
| `pypdf`, `img2pdf`, `pillow`, `pywin32` | PDF processing, images, Office automation | automatically (pip, in its own venv) |
| Microsoft Office **or** LibreOffice | only for Office documents | not automatic – PDFs and images work without |
| Git | to download the repository | `winget install --id Git.Git -e` |

## Installation

In **Windows PowerShell**:

```powershell
git clone https://github.com/robitschmatthias-ui/Windows-PDF-Merge-in-Explorer.git
cd Windows-PDF-Merge-in-Explorer
.\install\install.bat
```

Or double-click `install\install.bat` in Explorer.

The installer creates a Python venv inside the project folder and three shortcuts in the "Send to" folder. **No administrator rights** are needed. Keep the project folder where it is afterwards – the shortcuts point to it.

## Usage

1. Select files in Explorer
2. Right-click → **Send to** → **PDF Quick Merge** / **PDF Advanced Merge** / **Convert to PDF**

> Windows 11: "Send to" is in the classic menu – right-click → **Show more options**, or **Shift + right-click** directly.

### PDF Advanced Merge

In the dialog:
- Sort the files by **drag & drop** or with **▲ Up / ▼ Down**
- Choose the **page size** (keep original or fit to A4)
- Choose the **orientation** (original, portrait, landscape)
- Optional: **open the PDF after merging**

Then choose where to save – the name of the first file is suggested.

## Known limitations

### Digital signatures
When merging, signature objects are kept, but their checksum no longer matches. PDF programs will show the signature as **invalid**. Do not merge PDFs whose signature has to stay valid.

### Password-protected files
PDFs with an open password and password-protected Office files are skipped and listed in the final message. PDFs with permissions-only protection (no open password) work.

### Links and form fields when fitting to A4
When a page is scaled to A4, clickable links and form fields stay at their old position and no longer line up exactly with the scaled content. Without fitting to A4 this does not happen.

### Microsoft Office in the background
For Office files the tool starts Word/Excel/PowerPoint invisibly and quits them at the end. Macros in the documents are not run. If PowerPoint is already open with your own presentations, it is left running.

## Log

Errors are logged to `%LOCALAPPDATA%\pdf-merge\pdf-merge.log`.

## Uninstall

Run `install\uninstall.bat` – it removes the three "Send to" entries. Delete the project folder by hand afterwards.

---

## License

This project is licensed under the [GNU General Public License v3.0](LICENSE) – modified versions must stay open source as well.
