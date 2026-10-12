# CLAUDE.md – PDF Merge in Explorer (Windows)

## Binding decisions – do not ask again

Made by the repository owner on 2026-10-12:

- **Separate repository** for the Windows version. The Linux version stays in [Linux-PDF-Merge-in-Nemo](https://github.com/robitschmatthias-ui/Linux-PDF-Merge-in-Nemo); both offer the same three modes (quick, advanced, convert).
- **Public repository, in English**: code, comments, UI, installer output and `README.md` in English; `README.de.md` as German translation (same pattern as PDF-OCR-Converter). The Linux repository stays German.
- **Explorer integration via "Send to"** (shortcuts in the SendTo folder), not via registry context menus: "Send to" passes all selected files in one call; a classic context menu entry starts the program once per file.
- **Office files: Microsoft Office via COM, LibreOffice as fallback.**
- **"Convert to PDF" never overwrites**: existing PDF → `Name (2).pdf`, `Name (3).pdf` …, listed in the final message.
- **A4 + portrait/landscape**: rotate mismatched pages by 90° first, then fit to A4.
- **Advanced "Fit to A4" for images**: images reaching at least 200 DPI on A4 fill the page; smaller ones stay at original size, centered (same rule as the Linux version).
- **Commit author**: GitHub noreply address `272321947+robitschmatthias-ui@users.noreply.github.com`, no co-author lines.

## Technical notes

- `install/*.ps1` ASCII only – Windows PowerShell 5.1 reads files without BOM as ANSI.
- `.bat` files with CRLF, everything else LF (see `.gitattributes`).
- COM objects are bound to their thread: `OfficeConverter` is only used in the worker thread (`ui.run_with_progress` calls `CoInitialize` there) and is closed there after the progress window has closed – on some PCs Word takes ~35 s to quit.
- PowerPoint shares its instance with the user: `close()` quits an Office application only if no documents are open in it.
- No hard time limits for conversions; the progress window shows the current file.
