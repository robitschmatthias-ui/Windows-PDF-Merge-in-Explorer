"""PDF Merge for Windows Explorer ("Send to").

Usage:  pythonw pdf_merge.py --quick|--advanced|--convert FILE...

  --quick     merge right away -> Merge.pdf (Merge2.pdf, ...) in the folder
              of the first file
  --advanced  choose order, page size and orientation, then where to save
  --convert   convert each file to its own PDF in the same folder (images
              fitted to A4); existing PDFs are never overwritten
"""

from __future__ import annotations

import ctypes
import functools
import logging
import logging.handlers
import os
import shutil
import sys
import tempfile
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from office import OFFICE_EXT, OfficeConverter
from pdf_ops import (IMAGE_EXT, PDF_EXT, PdfError, image_to_pdf, is_image, is_pdf,
                     merge, open_pdf, postprocess, write_atomic)
from ui import AdvancedDialog, run_with_progress

APP_NAME = "PDF Merge"
LOG_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "pdf-merge"
ICON = Path(__file__).resolve().parent.parent / "pdf-merge-icon.ico"

log = logging.getLogger("pdf_merge")


def setup_logging() -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    handler = logging.handlers.RotatingFileHandler(
        LOG_DIR / "pdf-merge.log", maxBytes=1_000_000, backupCount=2, encoding="utf-8"
    )
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    logging.basicConfig(level=logging.INFO, handlers=[handler])


# ── Files ───────────────────────────────────────────────────────────────────

def explorer_order(paths: list[Path]) -> list[Path]:
    """Sorts like Explorer does (File2 before File10).

    "Send to" does not pass the files in display order; it starts with the
    file that was right-clicked.
    """
    cmp = ctypes.windll.shlwapi.StrCmpLogicalW
    cmp.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
    return sorted(paths, key=functools.cmp_to_key(lambda a, b: cmp(str(a), str(b))))


def collect_files(args: list[str]) -> list[Path]:
    files = []
    for arg in args:
        p = Path(arg)
        # Lock files of Office (~$name.docx) and LibreOffice (.~lock.name#)
        if p.name.startswith("~$") or p.name.startswith(".~lock."):
            continue
        if p.is_file():
            files.append(p.resolve())
    return explorer_order(files)


def free_name(directory: Path, stem: str, pattern: str) -> Path:
    """pattern "{stem} ({n})" -> report.pdf, report (2).pdf, ...
    pattern "{stem}{n}"      -> Merge.pdf, Merge2.pdf, ..."""
    candidate = directory / f"{stem}.pdf"
    n = 2
    while candidate.exists():
        candidate = directory / (pattern.format(stem=stem, n=n) + ".pdf")
        n += 1
    return candidate


def file_kind(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in PDF_EXT:
        return "pdf"
    if ext in IMAGE_EXT:
        return "image"
    if ext in OFFICE_EXT:
        return "office"
    # Unknown extension: detect by content
    if is_pdf(path):
        return "pdf"
    if is_image(path):
        return "image"
    return "unknown"


def to_pdf(path: Path, dst: Path, office: OfficeConverter, image_a4: str | None) -> None:
    kind = file_kind(path)
    if kind == "image":
        image_to_pdf(path, dst, a4=image_a4)
    elif kind == "office":
        office.convert(path, dst)
    else:
        raise PdfError("file type not supported")


def convert_for_merge(files, workdir, office, report, start, span, image_a4=None):
    """Returns (pdf_readers, errors) in the order of files.

    Every file is checked on its own: a broken or password-protected PDF is
    skipped and reported instead of aborting the whole merge.
    """
    pdfs, errors = [], []
    total = len(files)
    for i, f in enumerate(files):
        report(start + i * span // total, f"File {i + 1} of {total}: {f.name}")
        try:
            if file_kind(f) == "pdf":
                pdfs.append(open_pdf(f))
            else:
                dst = workdir / f"{i + 1}_{f.stem}.pdf"
                to_pdf(f, dst, office, image_a4)
                pdfs.append(open_pdf(dst))
        except Exception as exc:
            log.warning("%s: %s", f, exc)
            errors.append(f"{f.name} [{exc}]")
    return pdfs, errors


def lines(items: list[str]) -> str:
    return "\n".join(items)


# ── Modes ───────────────────────────────────────────────────────────────────

def quick_merge(root, files, workdir, office) -> None:
    output = free_name(files[0].parent, "Merge", "{stem}{n}")

    def job(report):
        pdfs, errors = convert_for_merge(files, workdir, office, report, 0, 90)
        if not pdfs:
            return errors, False
        report(95, "Merging PDF...")
        write_atomic(merge(pdfs), output)
        return errors, True

    try:
        errors, created = run_with_progress(root, "PDF Quick Merge", "Processing files...",
                                            job, after_done=office.close)
    except Exception as exc:
        log.exception("Quick merge failed")
        messagebox.showerror("Error", f"The PDF could not be created.\n\nDetails:\n{exc}")
        return

    if not created:
        messagebox.showerror("Error", "The PDF could not be created.\n\n"
                             f"Failed files:\n{lines(errors)}")
    elif errors:
        messagebox.showwarning("Partly successful",
                               "The PDF was created, but these files could not be "
                               f"converted:\n\n{lines(errors)}\n\nSaved as:\n{output}")
    else:
        messagebox.showinfo("Done!", f"PDF created:\n{output}")


def advanced_merge(root, files, workdir, office) -> None:
    options = AdvancedDialog(root, files).show()
    if options is None:
        return
    first = options.files[0]
    target = filedialog.asksaveasfilename(
        parent=root,
        title="Choose location and file name",
        initialdir=str(first.parent),
        initialfile=f"{first.stem}.pdf",
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
    )
    if not target:
        return
    output = Path(target)
    if output.suffix.lower() != ".pdf":
        output = output.with_name(output.name + ".pdf")

    # With "fit to A4", images go straight onto an A4 page using the 200 DPI
    # rule; postprocess() then leaves those pages at scale 1.
    image_a4 = "keep-small" if options.size == "a4" else None

    def job(report):
        pdfs, errors = convert_for_merge(options.files, workdir, office, report, 0, 70, image_a4)
        if not pdfs:
            return errors, False
        report(75, "Merging PDF...")
        writer = merge(pdfs)

        def page_progress(i, total):
            report(75 + i * 20 // total, f"Adjusting page {i} of {total}...")

        postprocess(writer, options.size, options.orient, page_progress)
        report(97, "Saving PDF...")
        write_atomic(writer, output)
        return errors, True

    try:
        errors, created = run_with_progress(root, "PDF Advanced Merge", "Processing files...",
                                            job, after_done=office.close)
    except Exception as exc:
        log.exception("Advanced merge failed")
        messagebox.showerror("Error", f"The PDF could not be created.\n\nDetails:\n{exc}")
        return

    if not created:
        messagebox.showerror("Error", "No usable files found.\n\n"
                             f"Failed files:\n{lines(errors)}")
        return
    if errors:
        messagebox.showwarning("Partly successful",
                               "The PDF was created, but these files could not be "
                               f"converted:\n\n{lines(errors)}\n\nSaved as:\n{output}")
    else:
        messagebox.showinfo("Done!", f"PDF created:\n{output}")
    if options.open_after:
        os.startfile(output)


def batch_convert(root, files, workdir, office) -> None:
    def job(report):
        success, renamed, skipped, errors = 0, [], [], []
        total = len(files)
        for i, f in enumerate(files):
            report(i * 100 // total, f"File {i + 1} of {total}: {f.name}")
            if file_kind(f) == "pdf":
                skipped.append(f.name)
                continue
            # Never overwrite an existing PDF: report.pdf -> report (2).pdf
            output = free_name(f.parent, f.stem, "{stem} ({n})")
            tmp = workdir / f"{i + 1}_{f.stem}.pdf"
            try:
                to_pdf(f, tmp, office, image_a4="fill")
                # shutil.move instead of os.replace: target may be on another drive
                shutil.move(str(tmp), str(output))
            except Exception as exc:
                log.warning("%s: %s", f, exc)
                errors.append(f"{f.name} [{exc}]")
                continue
            success += 1
            if output.name != f"{f.stem}.pdf":
                renamed.append(f"{f.name} → {output.name}")
        report(100)
        return success, renamed, skipped, errors

    try:
        success, renamed, skipped, errors = run_with_progress(
            root, "Convert to PDF", "Converting files...", job, after_done=office.close)
    except Exception as exc:
        log.exception("Conversion failed")
        messagebox.showerror("Error", f"Conversion failed.\n\nDetails:\n{exc}")
        return

    if success == 0 and not errors:
        messagebox.showinfo("Note", "All files already are PDFs – nothing to convert."
                            f"\n\n{lines(skipped)}")
    elif success and not errors:
        msg = f"{success} file(s) converted."
        if renamed:
            msg += f"\n\nSaved under a new name (PDF already existed):\n{lines(renamed)}"
        if skipped:
            msg += f"\n\nSkipped (already PDF):\n{lines(skipped)}"
        messagebox.showinfo("Done!", msg)
    elif success:
        msg = f"{success} file(s) converted.\n\nFailed:\n{lines(errors)}"
        if renamed:
            msg += f"\n\nSaved under a new name:\n{lines(renamed)}"
        if skipped:
            msg += f"\n\nSkipped:\n{lines(skipped)}"
        messagebox.showwarning("Partly successful", msg)
    else:
        messagebox.showerror("Error", f"No file converted.\n\n{lines(errors)}")


# ── Entry point ─────────────────────────────────────────────────────────────

MODES = {"--quick": quick_merge, "--advanced": advanced_merge, "--convert": batch_convert}


def main(argv: list[str]) -> int:
    setup_logging()
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # sharp rendering with display scaling
    except Exception:
        pass

    root = tk.Tk()
    root.withdraw()
    root.title(APP_NAME)
    if ICON.exists():
        root.iconbitmap(default=str(ICON))

    mode = argv[1] if len(argv) > 1 else ""
    if mode not in MODES:
        messagebox.showerror("Error", f"Unknown mode: {mode}\n\nUsage:\n"
                             "  pdf_merge.py --quick FILES...\n"
                             "  pdf_merge.py --advanced FILES...\n"
                             "  pdf_merge.py --convert FILES...")
        return 1

    files = collect_files(argv[2:])
    log.info("Start %s with %d file(s)", mode, len(files))
    if not files:
        messagebox.showerror("Error", "No usable files selected.")
        return 1

    with tempfile.TemporaryDirectory(prefix="pdf-merge-", ignore_cleanup_errors=True) as tmp:
        workdir = Path(tmp)
        office = OfficeConverter(workdir)
        MODES[mode](root, files, workdir, office)
    root.destroy()
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv))
    except Exception:
        log.exception("Unexpected error")
        raise
